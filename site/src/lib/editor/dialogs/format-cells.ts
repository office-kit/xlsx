// State and apply logic behind Format Cells. The dialog starts from the
// active cell's format and, like Excel, applies only what the user changed, so
// formatting a mixed selection does not flatten the attributes left alone.

import type { Cell, CellValue } from '@office-kit/xlsx/cell';
import { makeCell } from '@office-kit/xlsx/cell';
import type { Alignment, Border, Color, Font, FontPatch, HorizontalAlignment, PatternType, Side, SideStyle, UnderlineStyle, VerticalAlignment } from '@office-kit/xlsx/styles';
import { getCellAlignment, getCellBorder, getCellDisplayText, getCellFill, getCellFont, getCellProtection, makeBorder, makeColor } from '@office-kit/xlsx/styles';
import { createWorkbook } from '@office-kit/xlsx/workbook';
import { mergeCells, unmergeCells } from '@office-kit/xlsx/worksheet';
import { isWholeColumns, isWholeRows, MAX_COL, MAX_ROW, rangesIntersect, toBoundaries } from '../core/address.ts';
import { getCellAt } from '../core/cells.ts';
import type { EditorController } from '../core/controller.svelte.ts';
import { applyStyle, transformStyle, type EdgeInfo, type StylePatch } from '../core/format.ts';
import { buildNumberCode, DATE_CODES, FRACTION_CODES, numberFormatCategory, SPECIAL_CODES, TIME_CODES, type FormatCategory } from '../core/number-formats.ts';
import { resolveColor, type ThemePalette } from '../core/theme.ts';

// ---- Number ------------------------------------------------------------------

export const CURRENCY_SYMBOLS = ['$', '¥', '€', '£', ''] as const;

export interface NumberState {
  category: FormatCategory;
  decimals: number;
  thousands: boolean;
  symbol: string;
  negative: 0 | 1 | 2 | 3;
  /** Selected entry of the Date / Time / Fraction / Special type lists. */
  listCode: string;
  /** The Custom category's Type field. */
  customCode: string;
}

export function listFor(category: FormatCategory): readonly string[] {
  switch (category) {
    case 'date':
      return DATE_CODES;
    case 'time':
      return TIME_CODES;
    case 'fraction':
      return FRACTION_CODES;
    case 'special':
      return SPECIAL_CODES;
    default:
      return [];
  }
}

export function numberCode(s: NumberState): string {
  switch (s.category) {
    case 'general':
      return 'General';
    case 'text':
      return '@';
    case 'number':
      return buildNumberCode({ category: 'number', decimals: s.decimals, thousands: s.thousands, negative: s.negative });
    case 'currency':
      return buildNumberCode({ category: 'currency', decimals: s.decimals, symbol: s.symbol, negative: s.negative });
    case 'accounting':
      return buildNumberCode({ category: 'accounting', decimals: s.decimals, symbol: s.symbol });
    case 'percentage':
    case 'scientific':
      return buildNumberCode({ category: s.category, decimals: s.decimals });
    case 'date':
    case 'time':
    case 'fraction':
    case 'special':
      return s.listCode;
    case 'custom':
      return s.customCode;
  }
}

/** Recover the dialog's options from a format code; anything the options can't rebuild exactly is Custom. */
export function numberStateFor(code: string): NumberState {
  const base: NumberState = { category: 'custom', decimals: 2, thousands: false, symbol: '$', negative: 0, listCode: '', customCode: code };
  const cat = numberFormatCategory(code);
  if (cat === 'general' || cat === 'text') return { ...base, category: cat };
  const listed = (['date', 'time', 'fraction', 'special'] as const).find((c) => listFor(c).includes(code));
  if (listed) return { ...base, category: listed, listCode: code };
  const decimals = /\.(0+)/.exec(code.split(';')[0] ?? '')?.[1]?.length ?? 0;
  const candidates: NumberState[] = [];
  for (const negative of [0, 1, 2, 3] as const) {
    candidates.push({ ...base, category: 'number', decimals, thousands: false, negative });
    candidates.push({ ...base, category: 'number', decimals, thousands: true, negative });
    for (const symbol of CURRENCY_SYMBOLS) candidates.push({ ...base, category: 'currency', decimals, symbol, negative });
  }
  for (const symbol of CURRENCY_SYMBOLS) candidates.push({ ...base, category: 'accounting', decimals, symbol });
  candidates.push({ ...base, category: 'percentage', decimals }, { ...base, category: 'scientific', decimals });
  return candidates.find((c) => numberCode(c) === code) ?? base;
}

/**
 * Formats sample values through a private workbook, so previewing codes never
 * registers number formats in the document being edited.
 */
export class FormatSampler {
  readonly #wb = createWorkbook();

  constructor(date1904: boolean) {
    this.#wb.date1904 = date1904;
  }

  /** The text Excel would show; undefined when the code is not a valid format. */
  format(code: string, value: CellValue): string | undefined {
    try {
      const styleId = transformStyle(this.#wb, 0, { numFmt: code }, ALL_EDGES);
      return getCellDisplayText(this.#wb, makeCell(1, 1, value, styleId));
    } catch {
      // The library rejects malformed codes; the dialog reports that as Excel does.
      return undefined;
    }
  }
}

const ALL_EDGES: EdgeInfo = { top: true, bottom: true, left: true, right: true };

/** The value a sample shows: a formula's cached result, otherwise the cell's value. */
export function sampleValue(cell: Cell | undefined): CellValue {
  const v = cell?.value ?? null;
  if (v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula') {
    const cached = v.cachedValue;
    return typeof cached === 'number' || typeof cached === 'string' || typeof cached === 'boolean' ? cached : null;
  }
  return v;
}

// ---- Alignment ---------------------------------------------------------------

export interface AlignState {
  horizontal: HorizontalAlignment;
  vertical: VerticalAlignment;
  indent: number;
  wrap: boolean;
  shrink: boolean;
  merge: boolean;
  /** -90..90 degrees, as the dialog shows it. */
  degrees: number;
  /** Stacked vertical text (textRotation 255). */
  stacked: boolean;
}

/** OOXML stores downward angles as 91..180 (ECMA-376 §18.8.1 textRotation). */
export function rotationToDegrees(r: number): number {
  return r > 90 && r <= 180 ? 90 - r : r;
}

export function degreesToRotation(d: number): number {
  return d < 0 ? 90 - d : d;
}

// ---- Font --------------------------------------------------------------------

export type FontStyle = 'regular' | 'italic' | 'bold' | 'boldItalic';

export interface FontState {
  name: string;
  style: FontStyle;
  size: number;
  underline: UnderlineStyle;
  /** RRGGBB, or null for Automatic. */
  color: string | null;
  strike: boolean;
  superscript: boolean;
  subscript: boolean;
}

export const FONT_NAMES = ['Aptos', 'Aptos Narrow', 'Aptos Display', 'Calibri', 'Calibri Light', 'Arial', 'Helvetica', 'Times New Roman', 'Cambria', 'Georgia', 'Verdana', 'Courier New', 'Consolas', '游ゴシック', '游明朝', 'メイリオ', 'ＭＳ Ｐゴシック', 'ヒラギノ角ゴシック'];
export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

// ---- Border ------------------------------------------------------------------

export const EDGES = ['top', 'insideH', 'bottom', 'left', 'insideV', 'right', 'diagUp', 'diagDown'] as const;
export type Edge = (typeof EDGES)[number];

/** Line styles in the order of Excel's Style box (two columns, top to bottom). */
/** Excel's Line Style list, read top to bottom in two columns of seven. */
export const LINE_STYLES: readonly SideStyle[] = ['none', 'hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double'];

export interface BorderState {
  lineStyle: SideStyle;
  /** RRGGBB, or null for Automatic. */
  lineColor: string | null;
  edges: Record<Edge, Side | null>;
  /** Edges the user set or cleared; the rest stay as they are on each cell. */
  touched: Edge[];
}

/** Stroke geometry for drawing a line style in SVG previews. */
export function strokeOf(style: SideStyle | undefined): { width: number; dash: string; double: boolean } {
  switch (style) {
    case 'hair':
      return { width: 1, dash: '1 1', double: false };
    case 'dotted':
      return { width: 1, dash: '2 2', double: false };
    case 'dashDotDot':
      return { width: 1, dash: '6 2 2 2 2 2', double: false };
    case 'dashDot':
      return { width: 1, dash: '6 2 2 2', double: false };
    case 'dashed':
      return { width: 1, dash: '4 2', double: false };
    case 'mediumDashDotDot':
      return { width: 2, dash: '7 2 2 2 2 2', double: false };
    case 'slantDashDot':
      return { width: 2, dash: '8 2 3 2', double: false };
    case 'mediumDashDot':
      return { width: 2, dash: '7 2 2 2', double: false };
    case 'mediumDashed':
      return { width: 2, dash: '6 3', double: false };
    case 'medium':
      return { width: 2, dash: '', double: false };
    case 'thick':
      return { width: 3, dash: '', double: false };
    case 'double':
      return { width: 1, dash: '', double: true };
    default:
      return { width: 1, dash: '', double: false };
  }
}

function borderPatch(state: BorderState): StylePatch['border'] {
  const touched = new Set(state.touched);
  if (touched.size === 0) return undefined;
  const { edges } = state;
  return (cur, e) => {
    const b: { -readonly [K in keyof Border]: Border[K] } = { ...cur };
    const put = (k: 'top' | 'bottom' | 'left' | 'right', side: Side | null) => {
      if (side) b[k] = side;
      else delete b[k];
    };
    if (e.top ? touched.has('top') : touched.has('insideH')) put('top', e.top ? edges.top : edges.insideH);
    if (e.bottom ? touched.has('bottom') : touched.has('insideH')) put('bottom', e.bottom ? edges.bottom : edges.insideH);
    if (e.left ? touched.has('left') : touched.has('insideV')) put('left', e.left ? edges.left : edges.insideV);
    if (e.right ? touched.has('right') : touched.has('insideV')) put('right', e.right ? edges.right : edges.insideV);
    if (touched.has('diagUp') || touched.has('diagDown')) {
      const diag = edges.diagUp ?? edges.diagDown;
      if (diag) b.diagonal = diag;
      else delete b.diagonal;
      if (edges.diagUp) b.diagonalUp = true;
      else delete b.diagonalUp;
      if (edges.diagDown) b.diagonalDown = true;
      else delete b.diagonalDown;
    }
    return makeBorder(b);
  };
}

// ---- Fill & Protection -----------------------------------------------------------

export const PATTERNS: readonly PatternType[] = ['solid', 'gray0625', 'gray125', 'lightGray', 'mediumGray', 'darkGray', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis'];

export interface FillState {
  /** Background colour (RRGGBB), or null for No Color. */
  background: string | null;
  /** 'solid' when no pattern is chosen. */
  pattern: PatternType;
  patternColor: string | null;
}

export interface ProtectionState {
  locked: boolean;
  hidden: boolean;
}

// ---- whole dialog ----------------------------------------------------------------

export interface FormatCellsState {
  number: NumberState;
  align: AlignState;
  font: FontState;
  border: BorderState;
  fill: FillState;
  protection: ProtectionState;
}

function hexOf(color: Color | undefined, palette: ThemePalette): string | null {
  const css = resolveColor(color, palette, '');
  return css ? css.slice(1).toUpperCase() : null;
}

function fontStateOf(font: Font, palette: ThemePalette): FontState {
  return {
    name: font.name ?? 'Aptos Narrow',
    style: font.bold && font.italic ? 'boldItalic' : font.bold ? 'bold' : font.italic ? 'italic' : 'regular',
    size: font.size ?? 11,
    underline: font.underline ?? 'none',
    color: hexOf(font.color, palette),
    strike: font.strike === true,
    superscript: font.vertAlign === 'superscript',
    subscript: font.vertAlign === 'subscript',
  };
}

/** The workbook's Normal style font, which the Font tab's "Normal font" box restores. */
export function normalFontState(ctl: EditorController): FontState {
  const doc = ctl.doc;
  return fontStateOf(getCellFont(doc.wb, makeCell(1, 1, null, 0)), doc.styles.palette);
}

export function sameFont(a: FontState, b: FontState): boolean {
  return (Object.keys(a) as Array<keyof FontState>).every((k) => a[k] === b[k]);
}

export function initialState(ctl: EditorController): FormatCellsState {
  const doc = ctl.doc;
  const wb = doc.wb;
  const palette = doc.styles.palette;
  const { row, col } = doc.selection.active;
  const cell = getCellAt(doc.ws, row, col) ?? makeCell(row, col, null, ctl.defaultStyleAt(row, col));
  const font = getCellFont(wb, cell);
  const align = getCellAlignment(wb, cell);
  const border = getCellBorder(wb, cell);
  const fill = getCellFill(wb, cell);
  const protection = getCellProtection(wb, cell);
  const rotation = align.textRotation ?? 0;
  const single = doc.selection.ranges.length === 1 && doc.selection.ranges.every((r) => r.r1 === r.r2 && r.c1 === r.c2);
  const pattern = fill.kind === 'pattern' ? fill : undefined;
  const solid = pattern?.patternType === 'solid';
  const hasPattern = pattern !== undefined && pattern.patternType !== undefined && pattern.patternType !== 'none' && !solid;
  return {
    number: numberStateFor(doc.styles.get(cell.styleId).numFmt),
    align: {
      horizontal: align.horizontal ?? 'general',
      vertical: align.vertical ?? 'bottom',
      indent: align.indent ?? 0,
      wrap: align.wrapText === true,
      shrink: align.shrinkToFit === true,
      merge: doc.merges.at(row, col) !== undefined,
      degrees: rotation === 255 ? 0 : rotationToDegrees(rotation),
      stacked: rotation === 255,
    },
    font: fontStateOf(font, palette),
    border: {
      lineStyle: 'thin',
      lineColor: null,
      // Inside edges have no single-cell equivalent; the dialog starts them blank.
      edges: {
        top: single ? (border.top ?? null) : null,
        bottom: single ? (border.bottom ?? null) : null,
        left: single ? (border.left ?? null) : null,
        right: single ? (border.right ?? null) : null,
        insideH: null,
        insideV: null,
        diagUp: border.diagonalUp ? (border.diagonal ?? null) : null,
        diagDown: border.diagonalDown ? (border.diagonal ?? null) : null,
      },
      touched: [],
    },
    fill: {
      background: solid ? hexOf(pattern?.fgColor, palette) : hasPattern ? hexOf(pattern?.bgColor, palette) : null,
      pattern: hasPattern && pattern?.patternType ? pattern.patternType : 'solid',
      patternColor: hasPattern ? hexOf(pattern?.fgColor, palette) : null,
    },
    protection: { locked: protection.locked !== false, hidden: protection.hidden === true },
  };
}

function colorOf(hex: string | null, auto: Color): Color {
  return hex ? makeColor({ rgb: hex }) : auto;
}

/** Only the attributes that differ from the dialog's starting state. */
export function buildPatch(start: FormatCellsState, s: FormatCellsState): StylePatch {
  const patch: { -readonly [K in keyof StylePatch]: StylePatch[K] } = {};
  const code = numberCode(s.number);
  if (code !== numberCode(start.number)) patch.numFmt = code;

  // Undefined fields are removed from the cell's alignment by transformStyle.
  const alignment: { -readonly [K in keyof Alignment]?: Alignment[K] | undefined } = {};
  const a = s.align;
  const a0 = start.align;
  if (a.horizontal !== a0.horizontal) alignment.horizontal = a.horizontal === 'general' ? undefined : a.horizontal;
  if (a.vertical !== a0.vertical) alignment.vertical = a.vertical === 'bottom' ? undefined : a.vertical;
  if (a.indent !== a0.indent) alignment.indent = a.indent || undefined;
  if (a.wrap !== a0.wrap) alignment.wrapText = a.wrap || undefined;
  if (a.shrink !== a0.shrink) alignment.shrinkToFit = a.shrink || undefined;
  if (a.degrees !== a0.degrees || a.stacked !== a0.stacked) {
    const rotation = a.stacked ? 255 : degreesToRotation(a.degrees);
    alignment.textRotation = rotation || undefined;
  }
  if (Object.keys(alignment).length > 0) patch.alignment = alignment;

  const font: { -readonly [K in keyof FontPatch]: FontPatch[K] } = {};
  const f = s.font;
  const f0 = start.font;
  if (f.name !== f0.name) {
    font.name = f.name;
    font.scheme = undefined;
  }
  if (f.style !== f0.style) {
    font.bold = f.style === 'bold' || f.style === 'boldItalic' ? true : undefined;
    font.italic = f.style === 'italic' || f.style === 'boldItalic' ? true : undefined;
  }
  if (f.size !== f0.size) font.size = f.size;
  if (f.underline !== f0.underline) font.underline = f.underline === 'none' ? undefined : f.underline;
  if (f.color !== f0.color) font.color = colorOf(f.color, makeColor({ theme: 1 }));
  if (f.strike !== f0.strike) font.strike = f.strike || undefined;
  if (f.superscript !== f0.superscript || f.subscript !== f0.subscript) font.vertAlign = f.superscript ? 'superscript' : f.subscript ? 'subscript' : undefined;
  if (Object.keys(font).length > 0) patch.font = font;

  const fl = s.fill;
  const fl0 = start.fill;
  if (fl.background !== fl0.background || fl.pattern !== fl0.pattern || fl.patternColor !== fl0.patternColor) {
    if (fl.pattern === 'solid') {
      patch.fill = fl.background ? { kind: 'pattern', patternType: 'solid', fgColor: makeColor({ rgb: fl.background }) } : null;
    } else {
      patch.fill = {
        kind: 'pattern',
        patternType: fl.pattern,
        fgColor: colorOf(fl.patternColor, makeColor({ indexed: 64 })),
        ...(fl.background ? { bgColor: makeColor({ rgb: fl.background }) } : {}),
      };
    }
  }

  if (s.protection.locked !== start.protection.locked || s.protection.hidden !== start.protection.hidden) {
    patch.protection = { locked: s.protection.locked, hidden: s.protection.hidden };
  }

  const border = borderPatch(s.border);
  if (border) patch.border = border;
  return patch;
}

/** Apply the dialog as one undo step: the style patch plus a merge/unmerge when that box changed. */
export function applyFormatCells(ctl: EditorController, start: FormatCellsState, s: FormatCellsState): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const patch = buildPatch(start, s);
  const mergeChanged = s.align.merge !== start.align.merge;
  if (Object.keys(patch).length === 0 && !mergeChanged) return;
  doc.transact('Format Cells', (tx) => {
    if (mergeChanged) tx.sheet(ws, 'mergedCells');
    for (const range of doc.selection.ranges) {
      tx.cells(ws, range);
      if (range.r1 === 1 && range.r2 === MAX_ROW) tx.sheet(ws, 'columnDimensions');
      if (range.c1 === 1 && range.c2 === MAX_COL) tx.sheet(ws, 'rowDimensions');
      if (Object.keys(patch).length > 0) applyStyle(doc.wb, ws, range, patch);
      if (!mergeChanged || isWholeColumns(range) || isWholeRows(range)) continue;
      for (const existing of ws.mergedCells.slice()) {
        if (rangesIntersect({ r1: existing.minRow, c1: existing.minCol, r2: existing.maxRow, c2: existing.maxCol }, range)) unmergeCells(ws, existing);
      }
      if (s.align.merge && (range.r1 !== range.r2 || range.c1 !== range.c2)) mergeCells(ws, toBoundaries(range));
    }
  });
}
