// Excel's built-in table styles (TableStyleLight1…Dark11) and how a table's
// cells look under them.
//
// A workbook never carries the built-in definitions: Excel keeps them in the
// application, so they are rebuilt here. Fills, fonts and boldness were read
// back from Excel's own rendering of every built-in style (Range.DisplayFormat
// on a sample table with each style option on and off); borders, which Excel
// does not report there, follow its gallery. Colours stay theme-relative, so
// a workbook with another theme recolours its tables the way Excel does.
//
// Styles come in runs of seven: the first of each run is neutral (text and
// background shades), the other six use accents 1–6. Dark8–11 pair a header
// accent with a body accent instead.

import type { TableDefinition } from '@office-kit/xlsx/worksheet';
import { parseRangeAddress, type Range } from './address.ts';
import type { StrokeStyle } from './render-style.ts';
import { resolveColor, type ThemePalette } from './theme.ts';

/** A theme colour reference: SpreadsheetML theme index plus Excel's tint. */
export interface ThemeRef {
  readonly theme: number;
  readonly tint: number;
}

type LineKind = 'thin' | 'medium' | 'double';

export interface Line {
  readonly kind: LineKind;
  readonly color: ThemeRef;
}

interface ElementBorders {
  readonly top?: Line;
  readonly bottom?: Line;
  readonly left?: Line;
  readonly right?: Line;
  readonly insideH?: Line;
  readonly insideV?: Line;
}

interface Element {
  readonly fill?: ThemeRef;
  readonly font?: ThemeRef;
  readonly bold?: boolean;
  readonly border?: ElementBorders;
}

/** The parts of a style, lowest precedence first (ECMA-376 §18.8.40 lists the same order). */
const ELEMENT_ORDER = ['whole', 'colStripe1', 'colStripe2', 'rowStripe1', 'rowStripe2', 'lastCol', 'firstCol', 'header', 'total'] as const;
type ElementName = (typeof ELEMENT_ORDER)[number];
type StyleSpec = Partial<Record<ElementName, Element>>;

const BG = 0;
const TX = 1;
const ACCENT1 = 4;

// Excel's stored tints are not round numbers; these are the values its
// built-in styles use, so colours match Excel to the last digit.
const T80 = 0.7999816888943144;
const T60 = 0.5999938962981048;
const T40 = 0.3999755851924192;
const D15 = -0.1499984740745262;
const D25 = -0.249977111117893;
const D35 = -0.3499862666707358;
const D50 = -0.499984740745262;
const L15 = 0.1499984740745262;
const L25 = 0.249977111117893;
const L45 = 0.4499954222235861;
const L50 = 0.499984740745262;

const ref = (theme: number, tint = 0): ThemeRef => ({ theme, tint });
const thin = (color: ThemeRef): Line => ({ kind: 'thin', color });
const medium = (color: ThemeRef): Line => ({ kind: 'medium', color });
const double = (color: ThemeRef): Line => ({ kind: 'double', color });

/**
 * Colours of one style in a run. `accent` is 1–6, or 0 for the neutral first
 * style, whose tints Excel swaps for fixed text/background shades.
 */
function accentPalette(accent: number) {
  const neutral = accent === 0;
  const c = (tint = 0) => ref(neutral ? TX : ACCENT1 + accent - 1, tint);
  return {
    solid: c(),
    lighter80: neutral ? ref(BG, D15) : c(T80),
    lighter60: neutral ? ref(BG, D35) : c(T60),
    lighter40: neutral ? ref(TX, L50) : c(T40),
    darker25: c(neutral ? 0 : D25),
  };
}

const bold: Element = { bold: true };
const white = ref(BG);
const black = ref(TX);

function light(group: 0 | 1 | 2, accent: number): StyleSpec {
  const p = accentPalette(accent);
  if (group === 0) {
    return {
      whole: { font: p.darker25, border: { top: thin(p.solid), bottom: thin(p.solid) } },
      header: { bold: true, border: { bottom: thin(p.solid) } },
      total: { bold: true, border: { top: double(p.solid) } },
      firstCol: bold,
      lastCol: bold,
      rowStripe1: { fill: p.lighter80 },
      colStripe1: { fill: p.lighter80 },
    };
  }
  if (group === 1) {
    const line = thin(p.solid);
    return {
      whole: { border: { top: line, bottom: line, left: line, right: line } },
      header: { fill: p.solid, font: white, bold: true },
      total: { bold: true, border: { top: double(p.solid) } },
      firstCol: bold,
      lastCol: bold,
      rowStripe1: { border: { top: line, bottom: line } },
      colStripe1: { border: { left: line, right: line } },
    };
  }
  const line = thin(p.solid);
  return {
    whole: { border: { top: line, bottom: line, left: line, right: line, insideH: line, insideV: line } },
    header: { bold: true, border: { bottom: line } },
    total: { bold: true, border: { top: double(p.solid) } },
    firstCol: bold,
    lastCol: bold,
    rowStripe1: { fill: p.lighter80 },
    colStripe1: { fill: p.lighter80 },
  };
}

function mediumStyle(group: 0 | 1 | 2 | 3, accent: number): StyleSpec {
  const p = accentPalette(accent);
  const strong: Element = { fill: p.solid, font: white, bold: true };
  switch (group) {
    case 0: {
      const line = thin(p.lighter40);
      return {
        whole: { border: { top: line, bottom: line, left: line, right: line, insideH: line } },
        header: strong,
        total: { bold: true, border: { top: double(p.solid) } },
        firstCol: bold,
        lastCol: bold,
        rowStripe1: { fill: p.lighter80 },
        colStripe1: { fill: p.lighter80 },
      };
    }
    case 1: {
      const line = thin(white);
      return {
        whole: { fill: p.lighter80, border: { insideH: line, insideV: line } },
        header: { ...strong, border: { bottom: medium(white) } },
        total: { ...strong, border: { top: medium(white) } },
        firstCol: strong,
        lastCol: strong,
        rowStripe1: { fill: p.lighter60 },
        colStripe1: { fill: p.lighter60 },
      };
    }
    case 2:
      return {
        whole: { border: { top: medium(black), bottom: medium(black), insideH: thin(black) } },
        header: { ...strong, border: { bottom: medium(black) } },
        total: { border: { top: double(black) } },
        firstCol: strong,
        lastCol: strong,
        rowStripe1: { fill: ref(BG, D15) },
        colStripe1: { fill: ref(BG, D15) },
      };
    case 3: {
      const line = thin(p.lighter40);
      return {
        whole: { fill: p.lighter80, border: { top: line, bottom: line, left: line, right: line, insideH: line, insideV: line } },
        header: { bold: true },
        total: { bold: true, border: { top: double(p.solid) } },
        firstCol: bold,
        lastCol: bold,
        rowStripe1: { fill: p.lighter60 },
        colStripe1: { fill: p.lighter60 },
      };
    }
  }
}

function dark(accent: number): StyleSpec {
  const neutral = accent === 0;
  const c = (tint: number) => ref(neutral ? TX : ACCENT1 + accent - 1, tint);
  const body = neutral ? ref(TX, L45) : c(0);
  const band = neutral ? ref(TX, L25) : c(D25);
  const total = neutral ? ref(TX, L15) : c(D50);
  return {
    whole: { fill: body, font: white },
    header: { fill: black, font: white, bold: true, border: { bottom: medium(white) } },
    total: { fill: total, font: white, bold: true, border: { top: medium(white) } },
    firstCol: { fill: band, bold: true, border: { right: medium(white) } },
    lastCol: { fill: band, bold: true, border: { left: medium(white) } },
    rowStripe1: { fill: band },
    colStripe1: { fill: band },
  };
}

/** Dark8–11: a header accent over a lighter body accent (Dark8 is neutral). */
function darkPair(header: number, body: number): StyleSpec {
  const p = accentPalette(body);
  return {
    whole: { fill: p.lighter80 },
    header: { fill: header === 0 ? black : ref(ACCENT1 + header - 1), font: white },
    total: { fill: p.lighter80, bold: true, border: { top: double(black) } },
    firstCol: bold,
    lastCol: bold,
    rowStripe1: { fill: p.lighter60 },
    colStripe1: { fill: p.lighter60 },
  };
}

const BUILT_IN = /^TableStyle(Light|Medium|Dark)(\d{1,2})$/;
const DARK_PAIRS: ReadonlyArray<readonly [number, number]> = [
  [0, 0],
  [2, 1],
  [4, 3],
  [6, 5],
];

function builtInSpec(name: string): StyleSpec | undefined {
  const m = BUILT_IN.exec(name);
  if (!m?.[1] || !m[2]) return undefined;
  const n = Number(m[2]) - 1;
  const accent = n % 7;
  switch (m[1]) {
    case 'Light':
      return n < 21 ? light((Math.floor(n / 7) as 0 | 1 | 2), accent) : undefined;
    case 'Medium':
      return n < 28 ? mediumStyle((Math.floor(n / 7) as 0 | 1 | 2 | 3), accent) : undefined;
    default: {
      if (n < 7) return dark(accent);
      const pair = DARK_PAIRS[n - 7];
      return pair ? darkPair(pair[0], pair[1]) : undefined;
    }
  }
}

/** Every built-in style, in the order Excel's gallery lists them. */
export const BUILT_IN_TABLE_STYLES: readonly string[] = [
  ...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
  ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
  ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`),
];

/** What a table style adds to one cell; cell formatting still wins over it. */
export interface TableLook {
  readonly fill?: string;
  readonly color?: string;
  readonly bold?: boolean;
  readonly top?: StrokeStyle;
  readonly bottom?: StrokeStyle;
  readonly left?: StrokeStyle;
  readonly right?: StrokeStyle;
}

function stroke(line: Line, palette: ThemePalette): StrokeStyle {
  const color = resolveColor(line.color, palette, '#000000');
  switch (line.kind) {
    case 'thin':
      return { width: 1, color, dash: [], double: false };
    case 'medium':
      return { width: 2, color, dash: [], double: false };
    case 'double':
      return { width: 3, color, dash: [], double: true };
  }
}

/** Where a cell sits in its table; each flag selects or edges a style element. */
export interface CellPlace {
  readonly header: boolean;
  readonly total: boolean;
  readonly firstCol: boolean;
  readonly lastCol: boolean;
  /** Body row band (0 = first stripe, 1 = second), undefined outside the body. */
  readonly rowBand: 0 | 1 | undefined;
  readonly colBand: 0 | 1;
  readonly tableTop: boolean;
  readonly tableBottom: boolean;
}

export interface TableOptions {
  readonly rowStripes: boolean;
  readonly colStripes: boolean;
  readonly firstCol: boolean;
  readonly lastCol: boolean;
}

/** A style's contribution to one cell, still theme-relative. */
export interface RawLook {
  readonly fill?: ThemeRef;
  readonly font?: ThemeRef;
  readonly bold?: boolean;
  readonly top?: Line;
  readonly bottom?: Line;
  readonly left?: Line;
  readonly right?: Line;
}

/** Combine the style's elements that apply at `place`, as Excel layers them. */
export function rawLookAt(styleName: string, opts: TableOptions, place: CellPlace): RawLook | undefined {
  const spec = builtInSpec(styleName);
  if (!spec) return undefined;
  const look: { -readonly [K in keyof RawLook]: RawLook[K] } = {};
  const apply = (el: Element | undefined, edges: { top: boolean; bottom: boolean; left: boolean; right: boolean }) => {
    if (!el) return;
    if (el.fill) look.fill = el.fill;
    if (el.font) look.font = el.font;
    if (el.bold) look.bold = true;
    const b = el.border;
    if (!b) return;
    const top = edges.top ? b.top : b.insideH;
    const bottom = edges.bottom ? b.bottom : b.insideH;
    const left = edges.left ? b.left : b.insideV;
    const right = edges.right ? b.right : b.insideV;
    if (top) look.top = top;
    if (bottom) look.bottom = bottom;
    if (left) look.left = left;
    if (right) look.right = right;
  };
  const body = !place.header && !place.total;
  const cellEdges = { top: true, bottom: true, left: true, right: true };
  for (const name of ELEMENT_ORDER) {
    switch (name) {
      case 'whole':
        apply(spec.whole, { top: place.tableTop, bottom: place.tableBottom, left: place.firstCol, right: place.lastCol });
        break;
      case 'colStripe1':
      case 'colStripe2':
        if (opts.colStripes && body && place.colBand === (name === 'colStripe1' ? 0 : 1)) apply(spec[name], cellEdges);
        break;
      case 'rowStripe1':
      case 'rowStripe2':
        if (opts.rowStripes && place.rowBand === (name === 'rowStripe1' ? 0 : 1)) apply(spec[name], cellEdges);
        break;
      case 'lastCol':
        if (opts.lastCol && place.lastCol) apply(spec.lastCol, { top: place.tableTop, bottom: place.tableBottom, left: true, right: true });
        break;
      case 'firstCol':
        if (opts.firstCol && place.firstCol) apply(spec.firstCol, { top: place.tableTop, bottom: place.tableBottom, left: true, right: true });
        break;
      case 'header':
        if (place.header) apply(spec.header, { top: true, bottom: true, left: place.firstCol, right: place.lastCol });
        break;
      case 'total':
        if (place.total) apply(spec.total, { top: true, bottom: true, left: place.firstCol, right: place.lastCol });
        break;
    }
  }
  return Object.keys(look).length > 0 ? look : undefined;
}

function resolveLook(raw: RawLook | undefined, palette: ThemePalette): TableLook | undefined {
  if (!raw) return undefined;
  const look: { -readonly [K in keyof TableLook]: TableLook[K] } = {};
  if (raw.fill) look.fill = resolveColor(raw.fill, palette, '#FFFFFF');
  if (raw.font) look.color = resolveColor(raw.font, palette, '#000000');
  if (raw.bold) look.bold = true;
  if (raw.top) look.top = stroke(raw.top, palette);
  if (raw.bottom) look.bottom = stroke(raw.bottom, palette);
  if (raw.left) look.left = stroke(raw.left, palette);
  if (raw.right) look.right = stroke(raw.right, palette);
  return look;
}

interface CompiledTable {
  readonly range: Range;
  readonly headerRows: number;
  readonly totalRows: number;
  readonly style: string;
  readonly opts: TableOptions;
  /** Looks per place key, filled on first use: a table has only a few dozen distinct places. */
  readonly looks: Map<number, TableLook | undefined>;
}

export function hasBuiltInStyle(name: string | undefined): boolean {
  return name !== undefined && builtInSpec(name) !== undefined;
}

export function tableOptions(def: TableDefinition): TableOptions {
  const info = def.styleInfo;
  return {
    rowStripes: info?.showRowStripes === true,
    colStripes: info?.showColumnStripes === true,
    firstCol: info?.showFirstColumn === true,
    lastCol: info?.showLastColumn === true,
  };
}

function compile(def: TableDefinition): CompiledTable | undefined {
  const style = def.styleInfo?.name;
  const range = parseRangeAddress(def.ref)?.range;
  if (!style || !range || !builtInSpec(style)) return undefined;
  return { range, headerRows: def.headerRowCount ?? 1, totalRows: def.totalsRowCount ?? 0, style, opts: tableOptions(def), looks: new Map() };
}

/** Where (row, col) sits in a table covering `range`. */
export function placeOf(t: { readonly range: Range; readonly headerRows: number; readonly totalRows: number }, row: number, col: number): CellPlace {
  const { range } = t;
  const header = row < range.r1 + t.headerRows;
  const total = row > range.r2 - t.totalRows;
  const bodyIndex = row - range.r1 - t.headerRows;
  return {
    header,
    total,
    firstCol: col === range.c1,
    lastCol: col === range.c2,
    rowBand: header || total ? undefined : ((bodyIndex % 2) as 0 | 1),
    colBand: ((col - range.c1) % 2) as 0 | 1,
    tableTop: row === range.r1,
    tableBottom: row === range.r2,
  };
}

function placeKey(p: CellPlace): number {
  return (
    (p.header ? 1 : 0) |
    (p.total ? 2 : 0) |
    (p.firstCol ? 4 : 0) |
    (p.lastCol ? 8 : 0) |
    (p.rowBand === undefined ? 0 : p.rowBand === 0 ? 16 : 32) |
    (p.colBand === 1 ? 64 : 0) |
    (p.tableTop ? 128 : 0) |
    (p.tableBottom ? 256 : 0)
  );
}

/**
 * The table-style look of each cell on a sheet, or undefined when no table
 * uses a style this module knows. Per cell it is a scan of the sheet's tables
 * (a handful) and a memoised lookup, so painting stays proportional to the
 * visible cells.
 */
export function tableLooks(tables: readonly TableDefinition[], palette: ThemePalette): ((row: number, col: number) => TableLook | undefined) | undefined {
  const compiled = tables.flatMap((def) => compile(def) ?? []);
  if (compiled.length === 0) return undefined;
  return (row, col) => {
    for (const t of compiled) {
      const r = t.range;
      if (row < r.r1 || row > r.r2 || col < r.c1 || col > r.c2) continue;
      const place = placeOf(t, row, col);
      const key = placeKey(place);
      if (t.looks.has(key)) return t.looks.get(key);
      const look = resolveLook(rawLookAt(t.style, t.opts, place), palette);
      t.looks.set(key, look);
      return look;
    }
    return undefined;
  };
}
