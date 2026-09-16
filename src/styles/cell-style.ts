// Cell ↔ Stylesheet bridge.
//
// A cell's `styleId` is an index into `Workbook.styles.cellXfs`. Each CellXf
// points at slots in the font / fill / border / numFmt pools. To "apply" a Font
// to a cell we therefore:
//   1. read the cell's current CellXf (or `defaultCellXf` if it points
//      at a slot that hasn't been allocated yet — common right after
//      `makeCell` since `cellXfs` starts empty)
//   2. resolve / register the new component in its pool
//   3. build a new CellXf that carries the new id + the matching
//      `apply*` flag (Excel needs the flag to honour the override over
//      the underlying NamedStyle)
//   4. dedup that CellXf via `addCellXf` and write the returned index
//      back to `c.styleId`
//
// Following the no-classes rule, this module is just a flat list of free
// functions; the workbook is passed in so callers don't need to thread the
// stylesheet manually.

import type { Cell } from '../cell/cell.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import type { Workbook } from '../workbook/workbook.js';
import type { RangeRef } from '../utils/coordinate.js';
import { parseRange } from '../worksheet/cell-range.js';
import { ensureCell, type Worksheet } from '../worksheet/worksheet.js';
import type { Alignment, HorizontalAlignment, VerticalAlignment } from './alignment.js';
import { alignmentToCss, makeAlignment } from './alignment.js';
import type { Border, SideStyle } from './borders.js';
import { borderToCss, DEFAULT_BORDER, makeBorder, makeSide } from './borders.js';
import type { Color } from './colors.js';
import { makeColor } from './colors.js';
import type { Fill } from './fills.js';
import { DEFAULT_EMPTY_FILL, fillToCss, makePatternFill } from './fills.js';
import type { Font, UnderlineStyle } from './fonts.js';
import { DEFAULT_FONT, fontToCss, makeFont } from './fonts.js';
import { ensureBuiltinStyle } from './named-styles.js';
import { builtinFormatCode } from './numbers.js';
import type { Protection } from './protection.js';
import { DEFAULT_PROTECTION } from './protection.js';
import {
  addBorder,
  addCellXf,
  addFill,
  addFont,
  addNumFmt,
  type CellXf,
  defaultCellXf,
  type Stylesheet,
} from './stylesheet.js';

/** Default General number format code (numFmtId 0). */
const GENERAL_FORMAT_CODE = 'General';

/** Resolve a cell's current CellXf, falling back to defaults when unset. */
function currentXf(ss: Stylesheet, c: Cell): CellXf {
  return ss.cellXfs[c.styleId] ?? defaultCellXf();
}

// ---- read accessors --------------------------------------------------------

export function getCellFont(wb: Workbook, c: Cell): Font {
  const xf = currentXf(wb.styles, c);
  return wb.styles.fonts[xf.fontId] ?? DEFAULT_FONT;
}

export function getCellFill(wb: Workbook, c: Cell): Fill {
  const xf = currentXf(wb.styles, c);
  return wb.styles.fills[xf.fillId] ?? DEFAULT_EMPTY_FILL;
}

export function getCellBorder(wb: Workbook, c: Cell): Border {
  const xf = currentXf(wb.styles, c);
  return wb.styles.borders[xf.borderId] ?? DEFAULT_BORDER;
}

export function getCellAlignment(wb: Workbook, c: Cell): Alignment {
  return currentXf(wb.styles, c).alignment ?? {};
}

export function getCellProtection(wb: Workbook, c: Cell): Protection {
  return currentXf(wb.styles, c).protection ?? DEFAULT_PROTECTION;
}

/**
 * Returns the cell's number-format **code** (e.g. `"0.00"`, `"General"`).
 * Built-in IDs resolve through `builtinFormatCode`; custom IDs come from the
 * workbook's numFmts map.
 */
export function getCellNumberFormat(wb: Workbook, c: Cell): string {
  const id = currentXf(wb.styles, c).numFmtId;
  const builtin = builtinFormatCode(id);
  if (builtin !== undefined) return builtin;
  return wb.styles.numFmts.get(id) ?? GENERAL_FORMAT_CODE;
}

/**
 * Aggregate `fontToCss` + `fillToCss` + `borderToCss` + `alignmentToCss` for a
 * cell into a single CSS-property record. Resolves the cell's `styleId` against
 * the workbook stylesheet, then merges the four partials. On key collision the
 * priority is alignment > border > fill > font (alignment is most specific,
 * font is the broad default). A fully-default cell (`styleId === 0` with empty
 * pools) returns `{}`.
 */
export function cellStyleToCss(wb: Workbook, c: Cell): Record<string, string> {
  // Don't pay the resolve cost when the cell points at the default xf and the
  // stylesheet is still in its initial state.
  if (c.styleId === 0) {
    const xf = wb.styles.cellXfs[0];
    if (!xf || (xf.fontId === 0 && xf.fillId === 0 && xf.borderId === 0 && xf.alignment === undefined)) {
      return {};
    }
  }
  const font = getCellFont(wb, c);
  const fill = getCellFill(wb, c);
  const border = getCellBorder(wb, c);
  const alignment = getCellAlignment(wb, c);
  return {
    ...fontToCss(font),
    ...fillToCss(fill),
    ...borderToCss(border),
    ...alignmentToCss(alignment),
  };
}

// ---- write accessors -------------------------------------------------------

/**
 * Reserve cellXfs[0] for the implicit default xf when the pool is empty.
 * Excel's `<c>` elements without an `s=` attribute resolve to `cellXfs[0]`, so
 * the first time a caller styles any cell we need to make sure that slot stays
 * the default — otherwise unstyled cells in the same sheet would inherit the
 * freshly added styled xf.
 *
 * Idempotent: calling this on a non-empty pool is a no-op.
 */
const reserveDefaultXfSlot = (wb: Workbook): void => {
  if (wb.styles.cellXfs.length === 0) addCellXf(wb.styles, defaultCellXf());
};

/**
 * Replace one field on the cell's CellXf. Centralises the dedup + styleId
 * update so each `setCell*` is a single dispatch.
 */
function applyXfPatch(wb: Workbook, c: Cell, patch: Partial<CellXf>): void {
  reserveDefaultXfSlot(wb);
  const next: CellXf = { ...currentXf(wb.styles, c), ...patch };
  c.styleId = addCellXf(wb.styles, next);
}

/**
 * Replace the cell's font outright. `font` has to be complete: anything it
 * omits is omitted from the `<font>` record too, and each viewer then falls
 * back to its own default rather than the workbook's. Reach for {@link
 * patchCellFont} unless you really mean "this font and nothing inherited".
 */
export function setCellFont(wb: Workbook, c: Cell, font: Font): void {
  const fontId = addFont(wb.styles, font);
  applyXfPatch(wb, c, { fontId, applyFont: true });
}

export function setCellFill(wb: Workbook, c: Cell, fill: Fill): void {
  const fillId = addFill(wb.styles, fill);
  applyXfPatch(wb, c, { fillId, applyFill: true });
}

export function setCellBorder(wb: Workbook, c: Cell, border: Border): void {
  const borderId = addBorder(wb.styles, border);
  applyXfPatch(wb, c, { borderId, applyBorder: true });
}

export function setCellAlignment(wb: Workbook, c: Cell, alignment: Alignment): void {
  applyXfPatch(wb, c, { alignment, applyAlignment: true });
}

export function setCellProtection(wb: Workbook, c: Cell, protection: Protection): void {
  applyXfPatch(wb, c, { protection, applyProtection: true });
}

/**
 * Set the cell's number format by its **code** string. Built-in codes resolve
 * to their canonical id; custom codes are registered via `addNumFmt`.
 */
export function setCellNumberFormat(wb: Workbook, c: Cell, formatCode: string): void {
  const numFmtId = addNumFmt(wb.styles, formatCode);
  applyXfPatch(wb, c, { numFmtId, applyNumberFormat: true });
}

/**
 * Copy the source cell's `styleId` to the target cell. Both cells share the
 * same workbook stylesheet, so the styled appearance carries over without
 * allocating a new xf entry. Pass cells from different workbooks via {@link
 * cloneCellStyle} if you need a deep copy across workbooks.
 */
export function copyCellStyle(_wb: Workbook, source: Cell, target: Cell): void {
  target.styleId = source.styleId;
}

/**
 * Reset a cell back to the default (unstyled) appearance — equivalent to
 * Excel's "Clear Formatting" command. After the call, the cell inherits the
 * workbook's default font / fill / border / alignment / protection /
 * numberFormat. The underlying xf pool is **not** shrunk (Excel doesn't bother
 * either; the orphaned xf is harmless).
 */
export function clearCellStyle(_wb: Workbook, c: Cell): void {
  c.styleId = 0;
}

/**
 * Range-level shortcut for {@link clearCellStyle}. Walks every cell actually
 * present in the range and resets its `styleId` to 0; cells that don't exist
 * yet are **not** materialised (no-op for sparse regions, unlike the styled
 * `setRange*` family which has to create cells to make the patch observable).
 */
export function clearRangeStyle(wb: Workbook, ws: Worksheet, range: RangeRef): void {
  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  for (let r = minRow; r <= maxRow; r++) {
    const row = ws.rows.get(r);
    if (!row) continue;
    for (let c = minCol; c <= maxCol; c++) {
      const cell = row.get(c);
      if (cell) clearCellStyle(wb, cell);
    }
  }
}

/**
 * Deep-copy the source cell's full xf (font / fill / border / alignment /
 * protection / numberFormat) into a possibly-different workbook. Returns the
 * new styleId in the target workbook.
 */
export function cloneCellStyle(
  sourceWb: Workbook,
  source: Cell,
  targetWb: Workbook,
  target: Cell,
): number {
  reserveDefaultXfSlot(targetWb);
  const srcXf = currentXf(sourceWb.styles, source);
  const srcFont = sourceWb.styles.fonts[srcXf.fontId] ?? DEFAULT_FONT;
  const srcFill = sourceWb.styles.fills[srcXf.fillId] ?? DEFAULT_EMPTY_FILL;
  const srcBorder = sourceWb.styles.borders[srcXf.borderId] ?? DEFAULT_BORDER;
  const srcNumFmt = builtinFormatCode(srcXf.numFmtId)
    ?? sourceWb.styles.numFmts.get(srcXf.numFmtId)
    ?? GENERAL_FORMAT_CODE;
  const fontId = addFont(targetWb.styles, srcFont);
  const fillId = addFill(targetWb.styles, srcFill);
  const borderId = addBorder(targetWb.styles, srcBorder);
  const numFmtId = addNumFmt(targetWb.styles, srcNumFmt);
  const next: { -readonly [K in keyof CellXf]?: CellXf[K] } = {
    fontId,
    fillId,
    borderId,
    numFmtId,
  };
  if (srcXf.applyFont) next.applyFont = true;
  if (srcXf.applyFill) next.applyFill = true;
  if (srcXf.applyBorder) next.applyBorder = true;
  if (srcXf.applyNumberFormat) next.applyNumberFormat = true;
  if (srcXf.alignment !== undefined) {
    next.alignment = srcXf.alignment;
    next.applyAlignment = true;
  }
  if (srcXf.protection !== undefined) {
    next.protection = srcXf.protection;
    next.applyProtection = true;
  }
  target.styleId = addCellXf(targetWb.styles, next as CellXf);
  return target.styleId;
}

/**
 * The multi-axis style spec shared by {@link setCellStyle}, {@link
 * setRangeStyle} and {@link registerCellStyle}. Every axis is independent, so
 * pass any subset. A `font` here replaces the whole font rather than merging
 * into what the cell already had, because a style spec has to describe the same
 * appearance wherever it is applied. Use {@link patchCellFont} to change one
 * field and keep the rest.
 */
export interface CellStyleSpec {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  protection?: Protection;
  numberFormat?: string;
}

/** Resolve a style spec to the CellXf fields it sets, registering each component in its pool. */
const buildXfPatch = (wb: Workbook, spec: CellStyleSpec): { -readonly [K in keyof CellXf]?: CellXf[K] } => {
  const patch: { -readonly [K in keyof CellXf]?: CellXf[K] } = {};
  if (spec.font !== undefined) {
    patch.fontId = addFont(wb.styles, spec.font);
    patch.applyFont = true;
  }
  if (spec.fill !== undefined) {
    patch.fillId = addFill(wb.styles, spec.fill);
    patch.applyFill = true;
  }
  if (spec.border !== undefined) {
    patch.borderId = addBorder(wb.styles, spec.border);
    patch.applyBorder = true;
  }
  if (spec.alignment !== undefined) {
    patch.alignment = spec.alignment;
    patch.applyAlignment = true;
  }
  if (spec.protection !== undefined) {
    patch.protection = spec.protection;
    patch.applyProtection = true;
  }
  if (spec.numberFormat !== undefined) {
    patch.numFmtId = addNumFmt(wb.styles, spec.numberFormat);
    patch.applyNumberFormat = true;
  }
  return patch;
};

/**
 * Register a style and return its `styleId`, ready to pass to `setCell`,
 * `setCellByCoord` or `appendRow`. Build each distinct look once and let the
 * write carry the formatting, instead of running a styling pass over every cell
 * afterwards.
 *
 * The id names a **complete** style, not a patch: it starts from the workbook
 * default, so an axis missing from `spec` renders as the default even when the
 * target cell previously had something there. {@link setCellStyle} is the
 * patch-an-existing-cell counterpart.
 */
export function registerCellStyle(wb: Workbook, spec: CellStyleSpec): number {
  reserveDefaultXfSlot(wb);
  return addCellXf(wb.styles, { ...defaultCellXf(), ...buildXfPatch(wb, spec) });
}

/**
 * Build a single CellXf id from a multi-axis style spec, then apply it to every
 * cell in `range`. The xf is registered once per style shape, so a 1000-cell
 * range allocates one xf, much faster than looping `setCellStyle` per cell.
 */
export function setRangeStyle(wb: Workbook, ws: Worksheet, range: RangeRef, spec: CellStyleSpec): void {
  const patch = buildXfPatch(wb, spec);
  if (Object.keys(patch).length === 0) return;
  reserveDefaultXfSlot(wb);

  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  // Pre-register the xf for each existing cell — Excel dedupes by value, so the
  // inner xf-pool ends up the same shape for cells already carrying part of the
  // patch as for blanks.
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const cell = ensureCell(ws, r, c);
      const next: CellXf = { ...currentXf(wb.styles, cell), ...patch };
      cell.styleId = addCellXf(wb.styles, next);
    }
  }
}

/**
 * Combined cell-style setter. Each axis is independent — pass any subset and
 * the corresponding `applyXxx` flags get set on the underlying CellXf. Avoids
 * 5+ separate stylesheet round-trips when a caller wants to style a single cell
 * across multiple axes (Excel dedupes the resulting xf record on every call).
 */
export function setCellStyle(wb: Workbook, c: Cell, spec: CellStyleSpec): void {
  const patch = buildXfPatch(wb, spec);
  if (Object.keys(patch).length === 0) return;
  applyXfPatch(wb, c, patch as Partial<CellXf>);
}

// ---- fill presets -------------------------------------------------------

/**
 * Set the cell's background to a solid color. Accepts a hex string
 * (`'FFAAFFAA'`) or a partial `Color` object (`{ theme: 4, tint: 0.4 }`).
 * Equivalent to `setCellFill(wb, c, makePatternFill({ patternType: 'solid',
 * fgColor: makeColor(...) }))`.
 */
export function setCellBackgroundColor(wb: Workbook, c: Cell, color: string | Partial<Color>): void {
  const colorObj = typeof color === 'string' ? makeColor({ rgb: color }) : makeColor(color);
  setCellFill(wb, c, makePatternFill({ patternType: 'solid', fgColor: colorObj }));
}

/** Strip the cell's background fill, returning it to the default. */
export function clearCellBackground(wb: Workbook, c: Cell): void {
  setCellFill(wb, c, DEFAULT_EMPTY_FILL);
}

/**
 * Range-level shortcut for `setCellBackgroundColor`. Each cell in the range
 * gets the same solid pattern fill via `setRangeStyle`, so the fill pool dedups
 * to a single entry across the whole range.
 */
export function setRangeBackgroundColor(
  wb: Workbook,
  ws: Worksheet,
  range: RangeRef,
  color: string | Partial<Color>,
): void {
  const colorObj = typeof color === 'string' ? makeColor({ rgb: color }) : makeColor(color);
  setRangeStyle(wb, ws, range, {
    fill: makePatternFill({ patternType: 'solid', fgColor: colorObj }),
  });
}

/**
 * Range-level shortcut for `setCellFont`. Replaces the whole Font on every cell
 * in the range; there is no range-level merge, so build the complete font you
 * want rather than passing `makeFont({ bold: true })` and losing the workbook
 * default.
 */
export function setRangeFont(
  wb: Workbook,
  ws: Worksheet,
  range: RangeRef,
  font: Font,
): void {
  setRangeStyle(wb, ws, range, { font });
}

/**
 * Range-level shortcut for `setCellNumberFormat`. Stamps the same format-code
 * onto every cell in the range; the numFmt pool dedups the code so callers
 * don't pay per-cell pool churn.
 */
export function setRangeNumberFormat(
  wb: Workbook,
  ws: Worksheet,
  range: RangeRef,
  formatCode: string,
): void {
  setRangeStyle(wb, ws, range, { numberFormat: formatCode });
}

/**
 * Range-level shortcut for `setCellProtection`. Stamps the same Protection
 * (locked / hidden) onto every cell in the range. Pass a full `Protection`
 * value or a partial — partials default missing fields to `false` per Excel's
 * `<protection>` semantics.
 *
 * Common usage: `setRangeProtection(wb, ws, 'B2:B100', { locked: false })` to
 * leave just an input column editable when the sheet is protected.
 */
export function setRangeProtection(
  wb: Workbook,
  ws: Worksheet,
  range: RangeRef,
  protection: Protection | Partial<Protection>,
): void {
  // Funnel partials through the Protection factory so frozen invariant holds.
  const p: Protection = Object.isFrozen(protection)
    ? (protection as Protection)
    : { locked: protection.locked ?? false, hidden: protection.hidden ?? false };
  setRangeStyle(wb, ws, range, { protection: p });
}

/**
 * Range-level shortcut for `wrapCellText`. Toggles "Wrap Text" on every cell in
 * the range while preserving each cell's existing alignment (horizontal /
 * vertical / textRotation / indent are not touched). Empty cells in the range
 * are materialised so the alignment patch is observable on round-trip.
 */
export function setRangeWrapText(wb: Workbook, ws: Worksheet, range: RangeRef, on = true): void {
  reserveDefaultXfSlot(wb);
  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const cell = ensureCell(ws, r, c);
      wrapCellText(wb, cell, on);
    }
  }
}

/**
 * Range-level Alignment setter. Two modes:
 *
 *   - `mode: 'merge'` (default) — each cell's existing alignment is
 *     preserved; the supplied partial overlays it. Use this when you
 *     want to set just `horizontal` or `vertical` without wiping the
 *     other axes.
 *   - `mode: 'replace'` — each cell's alignment is **wholly replaced**
 *     by the supplied value. Indent / textRotation / wrapText that
 *     weren't supplied are dropped.
 *
 * Empty cells in the range are materialised so the patch is observable on
 * round-trip.
 */
export function setRangeAlignment(
  wb: Workbook,
  ws: Worksheet,
  range: RangeRef,
  alignment: Partial<Alignment>,
  mode: 'merge' | 'replace' = 'merge',
): void {
  reserveDefaultXfSlot(wb);
  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  if (mode === 'replace') {
    setRangeStyle(wb, ws, range, { alignment: makeAlignment(alignment) });
    return;
  }
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const cell = ensureCell(ws, r, c);
      const cur = currentXf(wb.styles, cell).alignment;
      setCellAlignment(wb, cell, mergeAlignment(cur, alignment));
    }
  }
}

// ---- font presets -------------------------------------------------------

const mergeFont = (current: Font, patch: Partial<Font>): Font => makeFont({ ...current, ...patch });

/**
 * Merge `patch` over the cell's current font, so the fields left out keep the
 * value they already had (the workbook default, on an unstyled cell). This is
 * the way to change several font fields at once: `setCellFont(wb, c,
 * makeFont({ bold: true }))` is a legal call that registers a font carrying no
 * name and no size, and Excel, LibreOffice and Sheets each substitute a
 * different one.
 *
 * The single-field setters below are this function with one field filled in.
 */
export function patchCellFont(wb: Workbook, c: Cell, patch: Partial<Font>): void {
  setCellFont(wb, c, mergeFont(getCellFont(wb, c), patch));
}

/** Toggle bold on a cell. Preserves other font fields. */
export function setBold(wb: Workbook, c: Cell, on = true): void {
  patchCellFont(wb, c, { bold: on });
}

/** Toggle italic on a cell. */
export function setItalic(wb: Workbook, c: Cell, on = true): void {
  patchCellFont(wb, c, { italic: on });
}

/** Toggle strike-through on a cell. */
export function setStrikethrough(wb: Workbook, c: Cell, on = true): void {
  patchCellFont(wb, c, { strike: on });
}

/**
 * Set the underline style. Pass `false` to drop underline; pass `'single' |
 * 'double' | 'singleAccounting' | 'doubleAccounting'` to apply that style; pass
 * `true` for the most common single-line.
 */
export function setUnderline(
  wb: Workbook,
  c: Cell,
  style: UnderlineStyle | boolean = 'single',
): void {
  const cur = getCellFont(wb, c);
  // Strip the existing underline by spreading then overwriting; makeFont
  // ignores `underline: undefined` so passing nothing for the off-case drops it
  // entirely.
  const { underline: _drop, ...rest } = cur;
  if (style === false) {
    setCellFont(wb, c, makeFont(rest));
    return;
  }
  const u = style === true ? 'single' : style;
  setCellFont(wb, c, makeFont({ ...rest, underline: u }));
}

/** Set the font size in points (e.g. 14). Preserves other fields. */
export function setFontSize(wb: Workbook, c: Cell, size: number): void {
  patchCellFont(wb, c, { size });
}

/** Set the font family name (e.g. "Arial"). Preserves other fields. */
export function setFontName(wb: Workbook, c: Cell, name: string): void {
  patchCellFont(wb, c, { name });
}

/**
 * Set the font color. Accepts a hex string ("FFAA0033") or a partial `Color`
 * object (`{ theme: 4, tint: 0.4 }`). Preserves other font fields.
 */
export function setFontColor(wb: Workbook, c: Cell, color: string | Partial<Color>): void {
  const colorObj = typeof color === 'string' ? makeColor({ rgb: color }) : makeColor(color);
  patchCellFont(wb, c, { color: colorObj });
}

// ---- alignment presets --------------------------------------------------

const mergeAlignment = (current: Alignment | undefined, patch: Partial<Alignment>): Alignment => {
  return makeAlignment({ ...current, ...patch });
};

/**
 * Center a cell horizontally + vertically. Mirrors Excel's "Merge & Center" UI
 * button (without the merge — see {@link mergeCells} for that). Preserves any
 * other alignment fields already present.
 */
export function centerCell(wb: Workbook, c: Cell): void {
  const cur = currentXf(wb.styles, c).alignment;
  setCellAlignment(wb, c, mergeAlignment(cur, { horizontal: 'center', vertical: 'center' }));
}

/** Toggle "Wrap Text" on a cell, preserving other alignment fields. */
export function wrapCellText(wb: Workbook, c: Cell, wrap = true): void {
  const cur = currentXf(wb.styles, c).alignment;
  setCellAlignment(wb, c, mergeAlignment(cur, { wrapText: wrap }));
}

/** Set the horizontal alignment in isolation. */
export function alignCellHorizontal(wb: Workbook, c: Cell, horizontal: HorizontalAlignment): void {
  const cur = currentXf(wb.styles, c).alignment;
  setCellAlignment(wb, c, mergeAlignment(cur, { horizontal }));
}

/** Set the vertical alignment in isolation. */
export function alignCellVertical(wb: Workbook, c: Cell, vertical: VerticalAlignment): void {
  const cur = currentXf(wb.styles, c).alignment;
  setCellAlignment(wb, c, mergeAlignment(cur, { vertical }));
}

/**
 * Rotate the cell's text. `degrees` accepts 0..180 (clockwise) or 255 for
 * Excel's "vertical stacked" mode. Mirrors the rotate icons in the alignment
 * ribbon.
 */
export function rotateCellText(wb: Workbook, c: Cell, degrees: number): void {
  const cur = currentXf(wb.styles, c).alignment;
  setCellAlignment(wb, c, mergeAlignment(cur, { textRotation: degrees }));
}

/** Set or clear the indent level (0..255). */
export function indentCell(wb: Workbook, c: Cell, levels: number): void {
  const cur = currentXf(wb.styles, c).alignment;
  setCellAlignment(wb, c, mergeAlignment(cur, { indent: levels }));
}

// ---- format presets -----------------------------------------------------
//
// Mirror Excel's "Format Cells → Number → Category" panel. Each preset builds
// the exact format-code Excel ships with, then routes through the existing
// `setCellNumberFormat` so the dedup pool sees the same string every time.

/**
 * Format a cell as currency. Produces one of:
 * - default → `"$#,##0.00"` (US dollar, 2 decimals)
 * - `{ symbol: "€" }` → `"€#,##0.00"`
 * - `{ symbol: "¥", decimals: 0 }` → `"¥#,##0"`
 * - `{ accounting: true }` → `"_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-"`
 * (Excel's "Accounting" subtype with right-aligned symbol).
 */
export function setCellAsCurrency(
  wb: Workbook,
  c: Cell,
  opts: { symbol?: string; decimals?: number; accounting?: boolean } = {},
): void {
  const symbol = opts.symbol ?? '$';
  const decimals = opts.decimals ?? 2;
  const decTail = decimals > 0 ? `.${'0'.repeat(decimals)}` : '';
  const code = opts.accounting
    ? `_-${symbol}* #,##0${decTail}_-;-${symbol}* #,##0${decTail}_-;_-${symbol}* "-"${'?'.repeat(decimals)}_-;_-@_-`
    : `${symbol}#,##0${decTail}`;
  setCellNumberFormat(wb, c, code);
}

/**
 * Format a cell as a percentage. `decimals` defaults to 0 → `"0%"`; `decimals:
 * 2` → `"0.00%"`. The cell value is multiplied by 100 by Excel during display.
 */
export function setCellAsPercent(wb: Workbook, c: Cell, decimals = 0): void {
  if (!Number.isInteger(decimals) || decimals < 0) {
    throw new OpenXmlSchemaError(`setCellAsPercent: decimals must be a non-negative integer; got ${decimals}`);
  }
  const code = decimals === 0 ? '0%' : `0.${'0'.repeat(decimals)}%`;
  setCellNumberFormat(wb, c, code);
}

/**
 * Format a cell as a date. `format` defaults to Excel's default
 * locale-independent ISO-style date `"yyyy-mm-dd"`. Common alternatives:
 * `"m/d/yyyy"`, `"dd-mmm-yy"`, `"yyyy-mm-dd hh:mm:ss"`.
 */
export function setCellAsDate(wb: Workbook, c: Cell, format = 'yyyy-mm-dd'): void {
  setCellNumberFormat(wb, c, format);
}

/**
 * Format a cell as a thousands-separated number. `decimals` defaults to 0 →
 * `"#,##0"`; `decimals: 2` → `"#,##0.00"`.
 */
export function setCellAsNumber(wb: Workbook, c: Cell, decimals = 0): void {
  if (!Number.isInteger(decimals) || decimals < 0) {
    throw new OpenXmlSchemaError(`setCellAsNumber: decimals must be a non-negative integer; got ${decimals}`);
  }
  const code = decimals === 0 ? '#,##0' : `#,##0.${'0'.repeat(decimals)}`;
  setCellNumberFormat(wb, c, code);
}

// ---- table-header preset ------------------------------------------------

/**
 * Apply Excel's stock "table header" formatting to a range: bold white text on
 * a dark fill, plus a thick bottom border. Override any axis via `opts` — pass
 * `bold: false` to drop the bold, or `fillColor: 'FF305496'` for a different
 * shade. Defaults match Excel's "Table Style Medium 2" header row.
 */
export function formatAsHeader(
  wb: Workbook,
  ws: Worksheet,
  range: RangeRef,
  opts: {
    fillColor?: string | Partial<Color>;
    fontColor?: string | Partial<Color>;
    bold?: boolean;
    bottomBorder?: SideStyle | false;
    bottomBorderColor?: string | Partial<Color>;
  } = {},
): void {
  const fillColor = opts.fillColor === undefined ? 'FF305496' : typeof opts.fillColor === 'string' ? makeColor({ rgb: opts.fillColor }) : makeColor(opts.fillColor);
  const fontColor = opts.fontColor === undefined ? 'FFFFFFFF' : typeof opts.fontColor === 'string' ? makeColor({ rgb: opts.fontColor }) : makeColor(opts.fontColor);
  const bold = opts.bold ?? true;
  const borderStyle = opts.bottomBorder ?? 'medium';
  const borderColorObj = opts.bottomBorderColor === undefined ? undefined : typeof opts.bottomBorderColor === 'string' ? makeColor({ rgb: opts.bottomBorderColor }) : makeColor(opts.bottomBorderColor);

  const fillColorObj = typeof fillColor === 'string' ? makeColor({ rgb: fillColor }) : fillColor;
  const fontColorObj = typeof fontColor === 'string' ? makeColor({ rgb: fontColor }) : fontColor;

  const styleOpts: CellStyleSpec = {
    font: makeFont({ bold, color: fontColorObj }),
    fill: makePatternFill({ patternType: 'solid', fgColor: fillColorObj }),
  };
  if (borderStyle !== false) {
    const side = makeSide({ style: borderStyle, ...(borderColorObj ? { color: borderColorObj } : {}) });
    styleOpts.border = makeBorder({ bottom: side });
  }
  setRangeStyle(wb, ws, range, styleOpts);
}

// ---- named / built-in style application --------------------------------

/**
 * Apply a built-in Excel style ("Heading 1" / "Total" / "Good" / "Bad" /
 * "Calculation" / etc.) to a single cell. Registers the built-in on the
 * Stylesheet (idempotent) and points the cell's xf at it via `xfId` while
 * inheriting the matching font/fill/border/ numFmt ids so the cell renders
 * correctly on its own.
 *
 * Throws when `name` isn't in {@link BUILTIN_NAMED_STYLES}; use {@link
 * applyNamedStyle} for user-registered styles.
 */
export function applyBuiltinStyle(wb: Workbook, c: Cell, name: string): void {
  const xfId = ensureBuiltinStyle(wb.styles, name);
  applyNamedStyleByXfId(wb, c, xfId);
}

/**
 * Apply a NamedStyle that's already registered on the workbook (via
 * `addNamedStyle` or `ensureBuiltinStyle`) to a single cell, by name.
 */
export function applyNamedStyle(wb: Workbook, c: Cell, name: string): void {
  const entry = wb.styles._namedStyleByName?.get(name);
  if (entry === undefined) {
    throw new OpenXmlSchemaError(`applyNamedStyle: no named style "${name}" registered`);
  }
  applyNamedStyleByXfId(wb, c, entry.xfId);
}

const applyNamedStyleByXfId = (wb: Workbook, c: Cell, xfId: number): void => {
  const styleXf = wb.styles.cellStyleXfs[xfId];
  if (!styleXf) {
    throw new OpenXmlSchemaError(`applyNamedStyle: cellStyleXfs[${xfId}] missing`);
  }
  // Excel only renders the cellStyleXf's font / fill / border when the cellXf
  // carries the matching `apply*` flag — even a `xfId` reference alone doesn't
  // make Excel inherit the values. So we copy every component the named style
  // touches *and* set the corresponding apply* flag, mirroring what real Excel
  // writes for `s=` referenced cells using a built-in named style.
  const patch: { -readonly [K in keyof CellXf]?: CellXf[K] } = {
    xfId,
    fontId: styleXf.fontId,
    fillId: styleXf.fillId,
    borderId: styleXf.borderId,
    numFmtId: styleXf.numFmtId,
    applyFont: true,
    applyFill: true,
    applyBorder: true,
    applyNumberFormat: true,
  };
  if (styleXf.alignment !== undefined) {
    patch.alignment = styleXf.alignment;
    patch.applyAlignment = true;
  }
  if (styleXf.protection !== undefined) {
    patch.protection = styleXf.protection;
    patch.applyProtection = true;
  }
  applyXfPatch(wb, c, patch as Partial<CellXf>);
};

// ---- border presets ----------------------------------------------------

/**
 * Apply the same {@link SideStyle} to all four edges of a single cell. Optional
 * color via hex string or `Color` partial. Equivalent to `setCellBorder(wb, c,
 * makeBorder({ left, right, top, bottom: side }))` with all four sides
 * identical.
 */
export function setCellBorderAll(
  wb: Workbook,
  c: Cell,
  opts: { style: SideStyle; color?: string | Partial<Color> } = { style: 'thin' },
): void {
  const colorObj = opts.color === undefined ? undefined : typeof opts.color === 'string' ? makeColor({ rgb: opts.color }) : makeColor(opts.color);
  const side = makeSide({ style: opts.style, ...(colorObj ? { color: colorObj } : {}) });
  setCellBorder(wb, c, makeBorder({ left: side, right: side, top: side, bottom: side }));
}

/**
 * Draw an outer border around a rectangular range. Cells on the perimeter
 * receive a partial border (only the edges that face outside the range); inner
 * cells are unaffected unless `inner` is provided, in which case every cell in
 * the range receives a border combining its perimeter edges with the `inner`
 * style for the inside edges.
 */
export function setRangeBorderBox(
  wb: Workbook,
  ws: Worksheet,
  range: RangeRef,
  opts: { style: SideStyle; color?: string | Partial<Color>; inner?: SideStyle } = { style: 'thin' },
): void {
  const { minRow, maxRow, minCol, maxCol } = parseRange(range);
  const colorObj = opts.color === undefined ? undefined : typeof opts.color === 'string' ? makeColor({ rgb: opts.color }) : makeColor(opts.color);
  const outer = makeSide({ style: opts.style, ...(colorObj ? { color: colorObj } : {}) });
  const inner =
    opts.inner !== undefined
      ? makeSide({ style: opts.inner, ...(colorObj ? { color: colorObj } : {}) })
      : undefined;
  for (let r = minRow; r <= maxRow; r++) {
    for (let col = minCol; col <= maxCol; col++) {
      const onTop = r === minRow;
      const onBottom = r === maxRow;
      const onLeft = col === minCol;
      const onRight = col === maxCol;
      // Skip cells that wouldn't get any styling (interior cells when no
      // inner).
      if (!inner && !onTop && !onBottom && !onLeft && !onRight) continue;
      const sides: { -readonly [K in keyof Border]?: Border[K] } = {};
      const top = onTop ? outer : inner;
      const bottom = onBottom ? outer : inner;
      const left = onLeft ? outer : inner;
      const right = onRight ? outer : inner;
      if (top !== undefined) sides.top = top;
      if (bottom !== undefined) sides.bottom = bottom;
      if (left !== undefined) sides.left = left;
      if (right !== undefined) sides.right = right;
      const cell = ensureCell(ws, r, col);
      setCellBorder(wb, cell, makeBorder(sides));
    }
  }
}
