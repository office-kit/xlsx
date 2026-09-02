// Stylesheet pool. The workbook holds dedup pools per style component (fonts /
// fills / borders / numFmts / cellXfs / cellStyleXfs); cells reference entries
// by index via a `styleId` (an index into cellXfs).
//
// All adds run through `add*` free functions that look up the pool via
// `stableStringify`-keyed maps so the same logical Font / Fill / etc. added
// 1000× lands in a single pool slot.

import { OpenXmlSchemaError } from '../utils/exceptions';
import { stableStringify } from '../utils/stable-stringify';
import type { Alignment } from './alignment';
import type { Border } from './borders';
import { DEFAULT_BORDER } from './borders';
import type { Fill } from './fills';
import { DEFAULT_EMPTY_FILL, DEFAULT_GRAY_FILL } from './fills';
import type { Font } from './fonts';
import { DEFAULT_FONT } from './fonts';
import { BUILTIN_FORMATS_MAX_SIZE, builtinFormatId } from './numbers';
import type { Protection } from './protection';

/**
 * One entry in the cellXfs / cellStyleXfs pool. Represents the union of
 * indexes-into-other-pools that a cell or named style points at.
 *
 * Mirrors openpyxl's CellStyle (styles/cell_style.py). The TS port keeps it as
 * a plain readonly object — no class, no methods. The Stylesheet `addCellXf`
 * allocates the index; cells store only that.
 */
export interface CellXf {
  readonly fontId: number;
  readonly fillId: number;
  readonly borderId: number;
  readonly numFmtId: number;
  /** Index into cellStyleXfs — only set for direct-cell xfs that point at a NamedStyle. */
  readonly xfId?: number;
  readonly alignment?: Alignment;
  readonly protection?: Protection;
  readonly applyFont?: boolean;
  readonly applyFill?: boolean;
  readonly applyBorder?: boolean;
  readonly applyNumberFormat?: boolean;
  readonly applyAlignment?: boolean;
  readonly applyProtection?: boolean;
  readonly pivotButton?: boolean;
  readonly quotePrefix?: boolean;
}

export interface Stylesheet {
  fonts: Font[];
  fills: Fill[];
  borders: Border[];
  /** numFmtId → format code (custom IDs only; built-ins implicit). */
  numFmts: Map<number, string>;
  cellXfs: CellXf[];
  cellStyleXfs: CellXf[];
  /** Named styles (Excel's "Cell Styles" gallery; populated by addNamedStyle). */
  namedStyles?: Array<import('./named-styles').StylesheetNamedStyle>;
  /**
   * The tail of `<styleSheet>` this model doesn't cover — `<tableStyles>`,
   * `<colors>` and `<extLst>` (which is where Excel keeps its x14/x15 slicer
   * and timeline styles). Captured verbatim in document order so re-saving a
   * loaded workbook keeps them; CT_Stylesheet puts all three after `<dxfs>`,
   * so the writer appends them there.
   */
  stylesXmlTail?: Array<import('../xml/tree').XmlNode>;

  // Internal dedup maps. Underscore-prefixed so JSON / structuredClone
  // serialisation can choose to skip them; never part of the public API.
  _fontIdByKey: Map<string, number>;
  _fillIdByKey: Map<string, number>;
  _borderIdByKey: Map<string, number>;
  _xfIdByKey: Map<string, number>;
  _styleXfIdByKey: Map<string, number>;
  _numFmtIdByCode: Map<string, number>;
  _namedStyleByName?: Map<string, import('./named-styles').StylesheetNamedStyle>;
}

/**
 * Build a fresh Stylesheet pre-populated with Excel's required default entries.
 * Mirrors openpyxl's empty Stylesheet: fonts: [DEFAULT_FONT] (index 0) fills:
 * [DEFAULT_EMPTY_FILL,
 *             DEFAULT_GRAY_FILL]          (indices 0, 1 — required)
 * borders: [DEFAULT_BORDER] (index 0) cellXfs: empty (indices allocated on
 * demand)
 */
export function makeStylesheet(): Stylesheet {
  const fontKey = stableStringify(DEFAULT_FONT);
  const fill0Key = stableStringify(DEFAULT_EMPTY_FILL);
  const fill1Key = stableStringify(DEFAULT_GRAY_FILL);
  const borderKey = stableStringify(DEFAULT_BORDER);

  return {
    fonts: [DEFAULT_FONT],
    fills: [DEFAULT_EMPTY_FILL, DEFAULT_GRAY_FILL],
    borders: [DEFAULT_BORDER],
    numFmts: new Map(),
    cellXfs: [],
    cellStyleXfs: [],
    _fontIdByKey: new Map([[fontKey, 0]]),
    _fillIdByKey: new Map([
      [fill0Key, 0],
      [fill1Key, 1],
    ]),
    _borderIdByKey: new Map([[borderKey, 0]]),
    _xfIdByKey: new Map(),
    _styleXfIdByKey: new Map(),
    _numFmtIdByCode: new Map(),
  };
}

// ---- pool add helpers ------------------------------------------------------

/** Add a Font to the pool, returning its 0-based index. Idempotent. */
export function addFont(ss: Stylesheet, font: Font): number {
  return addToPool(font, ss.fonts, ss._fontIdByKey);
}

/** Add a Fill to the pool, returning its 0-based index. Idempotent. */
export function addFill(ss: Stylesheet, fill: Fill): number {
  return addToPool(fill, ss.fills, ss._fillIdByKey);
}

/** Add a Border to the pool, returning its 0-based index. Idempotent. */
export function addBorder(ss: Stylesheet, border: Border): number {
  return addToPool(border, ss.borders, ss._borderIdByKey);
}

/**
 * Resolve a number-format string to its numFmtId.
 * - Built-in codes return their canonical OOXML ID.
 * - Otherwise the custom code is registered (and allocated an ID
 * ≥ {@link BUILTIN_FORMATS_MAX_SIZE}). Idempotent.
 */
export function addNumFmt(ss: Stylesheet, formatCode: string): number {
  const builtin = builtinFormatId(formatCode);
  if (builtin !== undefined) return builtin;
  const cached = ss._numFmtIdByCode.get(formatCode);
  if (cached !== undefined) return cached;
  const id = BUILTIN_FORMATS_MAX_SIZE + ss.numFmts.size;
  ss.numFmts.set(id, formatCode);
  ss._numFmtIdByCode.set(formatCode, id);
  return id;
}

/** Add a CellXf to the cellXfs pool, returning its 0-based index. Idempotent. */
export function addCellXf(ss: Stylesheet, xf: CellXf): number {
  validateCellXfRefs(ss, xf, /* isStyle */ false);
  return addToPool(xf, ss.cellXfs, ss._xfIdByKey);
}

/** Add a CellXf to the cellStyleXfs pool. Idempotent. */
export function addCellStyleXf(ss: Stylesheet, xf: CellXf): number {
  validateCellXfRefs(ss, xf, /* isStyle */ true);
  return addToPool(xf, ss.cellStyleXfs, ss._styleXfIdByKey);
}

// ---- internals -------------------------------------------------------------

const addToPool = <T>(value: T, pool: T[], byKey: Map<string, number>): number => {
  const key = stableStringify(value);
  const cached = byKey.get(key);
  if (cached !== undefined) return cached;
  const id = pool.length;
  pool.push(value);
  byKey.set(key, id);
  return id;
};

const validateCellXfRefs = (ss: Stylesheet, xf: CellXf, isStyle: boolean): void => {
  if (xf.fontId < 0 || xf.fontId >= ss.fonts.length) {
    throw new OpenXmlSchemaError(`CellXf.fontId ${xf.fontId} out of range [0, ${ss.fonts.length})`);
  }
  if (xf.fillId < 0 || xf.fillId >= ss.fills.length) {
    throw new OpenXmlSchemaError(`CellXf.fillId ${xf.fillId} out of range [0, ${ss.fills.length})`);
  }
  if (xf.borderId < 0 || xf.borderId >= ss.borders.length) {
    throw new OpenXmlSchemaError(`CellXf.borderId ${xf.borderId} out of range [0, ${ss.borders.length})`);
  }
  // numFmtId is permissive: built-ins or any registered custom ID.
  if (!Number.isInteger(xf.numFmtId) || xf.numFmtId < 0) {
    throw new OpenXmlSchemaError(`CellXf.numFmtId must be a non-negative integer; got ${xf.numFmtId}`);
  }
  if (!isStyle && xf.xfId !== undefined) {
    if (xf.xfId < 0 || xf.xfId >= ss.cellStyleXfs.length) {
      throw new OpenXmlSchemaError(`CellXf.xfId ${xf.xfId} out of range [0, ${ss.cellStyleXfs.length})`);
    }
  }
};

/** Returns the currently registered numFmt entries (built-ins are implicit and not included). */
export function getCustomNumFmts(ss: Stylesheet): ReadonlyArray<{ id: number; code: string }> {
  const out: Array<{ id: number; code: string }> = [];
  for (const [id, code] of ss.numFmts.entries()) out.push({ id, code });
  // numFmts is keyed by id; sort for deterministic output.
  out.sort((a, b) => a.id - b.id);
  return out;
}

/** Read-only snapshot of every Font entry in the pool, indexed by id. */
export function listFonts(ss: Stylesheet): ReadonlyArray<Font> {
  return ss.fonts;
}

/** Read-only snapshot of every Fill entry in the pool, indexed by id. */
export function listFills(ss: Stylesheet): ReadonlyArray<Fill> {
  return ss.fills;
}

/** Read-only snapshot of every Border entry in the pool, indexed by id. */
export function listBorders(ss: Stylesheet): ReadonlyArray<Border> {
  return ss.borders;
}

/** Read-only snapshot of every CellXf entry in the cellXfs pool. */
export function listCellXfs(ss: Stylesheet): ReadonlyArray<CellXf> {
  return ss.cellXfs;
}

/** Read-only snapshot of every CellStyleXf entry (named-style xfs). */
export function listCellStyleXfs(ss: Stylesheet): ReadonlyArray<CellXf> {
  return ss.cellStyleXfs;
}

/**
 * Convenience: build the default `cellXfs[0]` Excel emits — points at the
 * workbook's font 0 / fill 0 / border 0 / numFmtId 0 (General).
 */
export function defaultCellXf(): CellXf {
  return Object.freeze({ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 });
}
