// Conditional formatting: evaluation for the painter and the edits behind
// Home ▸ Conditional Formatting.
//
// The painter asks for an overlay per visible cell on every frame, so the
// expensive parts are paid once per model version: `conditionalOverlay` is
// rebuilt whenever `doc.version` changes, each rule's statistics (sorted
// values, average, duplicate counts) are computed lazily the first time a
// cell inside the rule asks, walking only the populated cells of its sqref,
// and finished overlays are memoised per cell. Rules are bucketed by row band
// on demand, so a cell outside every rule costs one Map lookup.
//
// The library keeps colorScale / dataBar / iconSet rules as their verbatim
// inner XML; `parseVisualRule` reads the handful of elements the renderer
// needs and `visualRuleXml` writes them back in the same shape.

import type { Cell } from '@office-kit/xlsx/cell';
import type { Color, DifferentialStyle } from '@office-kit/xlsx/styles';
import { addDxf, makeDifferentialStyle, makeFont, makePatternFill } from '@office-kit/xlsx/styles';
import type { Workbook } from '@office-kit/xlsx/workbook';
import type { CellIsOperator, CfvoType, ConditionalFormatting, ConditionalFormattingRule, TextOperator, TimePeriod, Worksheet } from '@office-kit/xlsx/worksheet';
import { makeCfRule, makeConditionalFormatting } from '@office-kit/xlsx/worksheet';
import { coerceToText, fromStorageFormula, toStorageFormula, translateFormula, type CalcEngine, type CalcScalar } from '../calc/index.ts';
import type { CellOverlay } from '../grid/paint.ts';
import type { Range } from './address.ts';
import { fromBoundaries, inRange, MAX_COL, parseRangeAddress, rangeAddress, rangesIntersect, toBoundaries } from './address.ts';
import { forEachCellInRange } from './cells.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';
import type { StyleResolver } from './render-style.ts';
import { resolveColor, type ThemePalette } from './theme.ts';

export type OverlayFn = (row: number, col: number, cell: Cell | undefined) => CellOverlay | undefined;

// ---- visual rule payloads ---------------------------------------------------

export interface CfvoSpec {
  readonly type: CfvoType;
  readonly val?: string;
  /** Icon sets: `gte="0"` makes the threshold strict (`>` instead of `>=`). */
  readonly gte?: boolean;
}

export type VisualRule =
  | { readonly kind: 'colorScale'; readonly cfvos: readonly CfvoSpec[]; readonly colors: readonly Color[] }
  | {
      readonly kind: 'dataBar';
      readonly cfvos: readonly CfvoSpec[];
      readonly color: Color;
      readonly showValue: boolean;
      readonly minLength: number;
      readonly maxLength: number;
      /**
       * Excel 2010+ bars (an x14 extension is attached): `min`/`max` behave as
       * autoMin/autoMax, so all-positive data grows from zero, and lengths span
       * the whole cell. Read-only: the editor writes classic bars.
       */
      readonly extended: boolean;
    }
  | {
      readonly kind: 'iconSet';
      readonly iconSet: string;
      readonly cfvos: readonly CfvoSpec[];
      readonly reverse: boolean;
      readonly showValue: boolean;
    };

const ATTR_RE = /([\w:]+)="([^"]*)"/g;

function decodeXml(text: string): string {
  return text.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (_, e: string) => {
    switch (e) {
      case 'lt':
        return '<';
      case 'gt':
        return '>';
      case 'amp':
        return '&';
      case 'quot':
        return '"';
      case 'apos':
        return "'";
    }
    return String.fromCodePoint(e.startsWith('#x') ? Number.parseInt(e.slice(2), 16) : Number.parseInt(e.slice(1), 10));
  });
}

function encodeXml(text: string): string {
  return text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
}

function attrs(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of tag.matchAll(ATTR_RE)) if (m[1] !== undefined && m[2] !== undefined) out[m[1]] = decodeXml(m[2]);
  return out;
}

const CFVO_TYPES: ReadonlySet<string> = new Set(['min', 'max', 'num', 'percent', 'percentile', 'formula']);

function parseCfvos(xml: string): CfvoSpec[] {
  const out: CfvoSpec[] = [];
  for (const m of xml.matchAll(/<(?:\w+:)?cfvo\b[^>]*>/g)) {
    const a = attrs(m[0]);
    const type = a['type'] ?? 'min';
    // x14 data bars write autoMin/autoMax; the base part's min/max is the closest reading.
    const mapped = type === 'autoMin' ? 'min' : type === 'autoMax' ? 'max' : type;
    if (!CFVO_TYPES.has(mapped)) continue;
    out.push({
      type: mapped as CfvoType,
      ...(a['val'] !== undefined ? { val: a['val'] } : {}),
      ...(a['gte'] === '0' || a['gte'] === 'false' ? { gte: false } : {}),
    });
  }
  return out;
}

function parseColors(xml: string): Color[] {
  const out: Color[] = [];
  for (const m of xml.matchAll(/<(?:\w+:)?color\b[^>]*>/g)) {
    const a = attrs(m[0]);
    const color: { -readonly [K in keyof Color]: Color[K] } = {};
    if (a['rgb'] !== undefined) color.rgb = a['rgb'].toUpperCase();
    if (a['theme'] !== undefined) color.theme = Number(a['theme']);
    if (a['indexed'] !== undefined) color.indexed = Number(a['indexed']);
    if (a['tint'] !== undefined) color.tint = Number(a['tint']);
    out.push(color);
  }
  return out;
}

const flag = (v: string | undefined, fallback: boolean): boolean => (v === undefined ? fallback : v === '1' || v === 'true');

/** The colour scale / data bar / icon set settings of a visual rule, or undefined for other rule types. */
export function parseVisualRule(rule: ConditionalFormattingRule): VisualRule | undefined {
  const xml = rule.innerXml ?? '';
  const root = /<(?:\w+:)?(colorScale|dataBar|iconSet)\b[^>]*>/.exec(xml);
  if (rule.type === 'colorScale') {
    const cfvos = parseCfvos(xml);
    const colors = parseColors(xml);
    if (cfvos.length < 2 || colors.length !== cfvos.length) return undefined;
    return { kind: 'colorScale', cfvos, colors };
  }
  if (rule.type === 'dataBar') {
    const a = attrs(root?.[0] ?? '');
    const cfvos = parseCfvos(xml);
    const extended = /<(?:\w+:)?extLst\b/.test(xml);
    return {
      kind: 'dataBar',
      cfvos: cfvos.length >= 2 ? cfvos : [{ type: 'min' }, { type: 'max' }],
      color: parseColors(xml)[0] ?? { rgb: 'FF638EC6' },
      showValue: flag(a['showValue'], true),
      minLength: Number(a['minLength'] ?? (extended ? 0 : 10)),
      maxLength: Number(a['maxLength'] ?? (extended ? 100 : 90)),
      extended,
    };
  }
  if (rule.type === 'iconSet') {
    const a = attrs(root?.[0] ?? '');
    const iconSet = a['iconSet'] ?? '3TrafficLights1';
    const cfvos = parseCfvos(xml);
    return {
      kind: 'iconSet',
      iconSet,
      cfvos: cfvos.length > 0 ? cfvos : defaultIconCfvos(iconCount(iconSet)),
      reverse: flag(a['reverse'], false),
      showValue: flag(a['showValue'], true),
    };
  }
  return undefined;
}

function cfvoXml(c: CfvoSpec): string {
  const val = c.val !== undefined ? ` val="${encodeXml(c.val)}"` : '';
  const gte = c.gte === false ? ' gte="0"' : '';
  return `<cfvo type="${c.type}"${val}${gte}/>`;
}

function colorXml(c: Color): string {
  if (c.theme !== undefined) return `<color theme="${c.theme}"${c.tint ? ` tint="${c.tint}"` : ''}/>`;
  if (c.indexed !== undefined) return `<color indexed="${c.indexed}"/>`;
  return `<color rgb="${c.rgb ?? 'FF000000'}"/>`;
}

export function visualRuleXml(v: VisualRule): string {
  switch (v.kind) {
    case 'colorScale':
      return `<colorScale>${v.cfvos.map(cfvoXml).join('')}${v.colors.map(colorXml).join('')}</colorScale>`;
    case 'dataBar': {
      const len = v.minLength !== 10 || v.maxLength !== 90 ? ` minLength="${v.minLength}" maxLength="${v.maxLength}"` : '';
      return `<dataBar${len}${v.showValue ? '' : ' showValue="0"'}>${v.cfvos.map(cfvoXml).join('')}${colorXml(v.color)}</dataBar>`;
    }
    case 'iconSet':
      return `<iconSet iconSet="${v.iconSet}"${v.reverse ? ' reverse="1"' : ''}${v.showValue ? '' : ' showValue="0"'}>${v.cfvos.map(cfvoXml).join('')}</iconSet>`;
  }
}

// ---- icon sets --------------------------------------------------------------

const RED = '#E5534B';
const YELLOW = '#F2B705';
const GREEN = '#3D9A47';
const GRAY = '#7F7F7F';
const BLACK = '#262626';
const BLUE = '#3A78C3';

interface Icon {
  readonly glyph: string;
  readonly color: string;
}

const icon = (glyph: string, color: string): Icon => ({ glyph, color });

/** Lowest threshold first, as the cfvos are ordered. */
export const ICON_SETS: Readonly<Record<string, readonly Icon[]>> = {
  '3Arrows': [icon('↓', RED), icon('→', YELLOW), icon('↑', GREEN)],
  '3ArrowsGray': [icon('↓', GRAY), icon('→', GRAY), icon('↑', GRAY)],
  '3Flags': [icon('⚑', RED), icon('⚑', YELLOW), icon('⚑', GREEN)],
  '3TrafficLights1': [icon('●', RED), icon('●', YELLOW), icon('●', GREEN)],
  '3TrafficLights2': [icon('●', RED), icon('●', YELLOW), icon('●', GREEN)],
  '3Signs': [icon('◆', RED), icon('▲', YELLOW), icon('●', GREEN)],
  '3Symbols': [icon('✖', RED), icon('!', YELLOW), icon('✔', GREEN)],
  '3Symbols2': [icon('✖', RED), icon('!', YELLOW), icon('✔', GREEN)],
  '3Triangles': [icon('▼', RED), icon('▬', YELLOW), icon('▲', GREEN)],
  '3Stars': [icon('☆', YELLOW), icon('⯪', YELLOW), icon('★', YELLOW)],
  '4Arrows': [icon('↓', RED), icon('↘', YELLOW), icon('↗', YELLOW), icon('↑', GREEN)],
  '4ArrowsGray': [icon('↓', GRAY), icon('↘', GRAY), icon('↗', GRAY), icon('↑', GRAY)],
  '4RedToBlack': [icon('●', BLACK), icon('●', GRAY), icon('●', '#F4A6A6'), icon('●', RED)],
  '4Rating': [icon('▂', BLUE), icon('▄', BLUE), icon('▆', BLUE), icon('█', BLUE)],
  '4TrafficLights': [icon('●', BLACK), icon('●', RED), icon('●', YELLOW), icon('●', GREEN)],
  '5Arrows': [icon('↓', RED), icon('↘', YELLOW), icon('→', YELLOW), icon('↗', YELLOW), icon('↑', GREEN)],
  '5ArrowsGray': [icon('↓', GRAY), icon('↘', GRAY), icon('→', GRAY), icon('↗', GRAY), icon('↑', GRAY)],
  '5Quarters': [icon('○', BLACK), icon('◔', BLACK), icon('◑', BLACK), icon('◕', BLACK), icon('●', BLACK)],
  '5Rating': [icon('▁', BLUE), icon('▂', BLUE), icon('▄', BLUE), icon('▆', BLUE), icon('█', BLUE)],
  '5Boxes': [icon('□', BLUE), icon('◱', BLUE), icon('◧', BLUE), icon('◩', BLUE), icon('■', BLUE)],
};

export function iconCount(iconSet: string): number {
  return ICON_SETS[iconSet]?.length ?? Number(iconSet[0] ?? 3);
}

/** Excel's default thresholds for a new icon set: equal percent bands. */
export function defaultIconCfvos(n: number): CfvoSpec[] {
  return Array.from({ length: n }, (_, i) => ({ type: 'percent' as const, val: String(Math.round((i * 100) / n)) }));
}

// ---- differential formats ---------------------------------------------------

/** The editable subset of a DXF the dialogs expose, colours as `#RRGGBB`. */
export interface CfFormat {
  readonly fill?: string;
  readonly color?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
}

type MutableOverlay = { -readonly [K in keyof CellOverlay]: CellOverlay[K] };

function dxfsOf(wb: Workbook): ReadonlyArray<DifferentialStyle> {
  // The stylesheet carries `dxfs` only once a DXF exists; the field is not on
  // the public `Stylesheet` type, so widen it with the optional member.
  const ss: Workbook['styles'] & { dxfs?: DifferentialStyle[] } = wb.styles;
  return ss.dxfs ?? [];
}

/**
 * The fill colour of a DXF. Inside a DXF a solid fill conventionally keeps its
 * colour in `bgColor` (Excel writes `<patternFill><bgColor/>` with no
 * pattern type), so that wins over `fgColor`.
 */
function dxfFillColor(dxf: DifferentialStyle, palette: ThemePalette): string | undefined {
  const fill = dxf.fill;
  if (!fill) return undefined;
  if (fill.kind === 'gradient') {
    const first = fill.stops[0];
    return first ? resolveColor(first.color, palette, '#FFFFFF') : undefined;
  }
  if (fill.patternType === 'none') return undefined;
  const color = fill.bgColor ?? fill.fgColor;
  return color ? resolveColor(color, palette, '#FFFFFF') : undefined;
}

function dxfOverlay(dxf: DifferentialStyle | undefined, palette: ThemePalette): CellOverlay {
  const out: MutableOverlay = {};
  if (!dxf) return out;
  const fill = dxfFillColor(dxf, palette);
  if (fill) out.fill = fill;
  const font = dxf.font;
  if (font) {
    if (font.color) out.color = resolveColor(font.color, palette, '#000000');
    if (font.bold !== undefined) out.bold = font.bold;
    if (font.italic !== undefined) out.italic = font.italic;
    if (font.strike !== undefined) out.strike = font.strike;
    if (font.underline !== undefined) out.underline = font.underline !== 'none';
  }
  return out;
}

/** Read a DXF back into the dialog's format model (for editing an existing rule). */
export function cfFormatOf(wb: Workbook, dxfId: number | undefined, palette: ThemePalette): CfFormat {
  const o = dxfOverlay(dxfId === undefined ? undefined : dxfsOf(wb)[dxfId], palette);
  return {
    ...(o.fill !== undefined ? { fill: o.fill } : {}),
    ...(o.color !== undefined ? { color: o.color } : {}),
    ...(o.bold ? { bold: true } : {}),
    ...(o.italic ? { italic: true } : {}),
    ...(o.underline ? { underline: true } : {}),
    ...(o.strike ? { strike: true } : {}),
  };
}

const argb = (hex: string): Color => ({ rgb: `FF${hex.replace('#', '').toUpperCase()}` });

/**
 * Register `format` as a DXF and return its id. The DXF pool is append-only
 * with de-duplication like the other style pools, so it is not part of the
 * undo snapshot: an orphaned entry after undo is harmless.
 */
export function registerCfFormat(wb: Workbook, format: CfFormat): number {
  const font = makeFont({
    ...(format.color !== undefined ? { color: argb(format.color) } : {}),
    ...(format.bold !== undefined ? { bold: format.bold } : {}),
    ...(format.italic !== undefined ? { italic: format.italic } : {}),
    ...(format.strike !== undefined ? { strike: format.strike } : {}),
    ...(format.underline !== undefined ? { underline: format.underline ? 'single' : 'none' } : {}),
  });
  const hasFont = format.color !== undefined || format.bold !== undefined || format.italic !== undefined || format.strike !== undefined || format.underline !== undefined;
  const dxf = makeDifferentialStyle({
    ...(hasFont ? { font } : {}),
    ...(format.fill !== undefined ? { fill: makePatternFill({ patternType: 'solid', fgColor: argb(format.fill), bgColor: argb(format.fill) }) } : {}),
  });
  return addDxf(wb.styles, dxf);
}

// ---- statistics -------------------------------------------------------------

interface Stats {
  /** Numeric values, ascending. */
  readonly sorted: Float64Array;
  readonly mean: number;
  /** Sample standard deviation (Excel's STDEV), 0 for fewer than two values. */
  readonly stdDev: number;
}

function percentileInc(sorted: Float64Array, p: number): number {
  if (sorted.length === 0) return 0;
  const pos = Math.min(1, Math.max(0, p)) * (sorted.length - 1);
  const lo = Math.floor(pos);
  const a = sorted[lo] ?? 0;
  const b = sorted[Math.min(lo + 1, sorted.length - 1)] ?? a;
  return a + (b - a) * (pos - lo);
}

function dupKey(v: CalcScalar): string | undefined {
  if (v === null || v === '') return undefined;
  if (typeof v === 'number') return `n${v}`;
  if (typeof v === 'boolean') return `b${v}`;
  if (typeof v === 'string') return `s${v.toLowerCase()}`;
  return `e${v.code}`;
}

const isError = (v: CalcScalar): v is Extract<CalcScalar, { kind: 'error' }> => typeof v === 'object' && v !== null;

const isBlankText = (v: CalcScalar): boolean => v === null || (typeof v === 'string' && v.trim() === '');

/** Text Excel's SEARCH sees for a value (numbers in General form). */
function textOf(v: CalcScalar): string {
  if (v === null) return '';
  const t = coerceToText(v);
  return typeof t === 'string' ? t : `${t.code}`;
}

/** Excel's comparison order for cellIs: numbers < text < booleans; text compares case-insensitively; blank acts as 0 or "". */
function compare(a: CalcScalar, b: CalcScalar): number | undefined {
  if (isError(a) || isError(b)) return undefined;
  const rank = (v: Exclude<CalcScalar, { kind: 'error' }>): number => (typeof v === 'number' || v === null ? 0 : typeof v === 'string' ? 1 : 2);
  const av = a === null ? (typeof b === 'string' ? '' : 0) : a;
  const bv = b === null ? (typeof a === 'string' ? '' : 0) : b;
  const ra = rank(av);
  const rb = rank(bv);
  if (ra !== rb) return ra - rb;
  if (typeof av === 'string' && typeof bv === 'string') {
    const x = av.toLowerCase();
    const y = bv.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return Number(av) - Number(bv);
}

const truthy = (v: CalcScalar): boolean => v === true || (typeof v === 'number' && v !== 0);

// ---- dates ------------------------------------------------------------------

const MS_PER_DAY = 86_400_000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

function serialToDate(serial: number, date1904: boolean): Date {
  return new Date((date1904 ? EPOCH_1904 : EPOCH_1900) + Math.floor(serial) * MS_PER_DAY);
}

/** Whether day `d` (a serial) falls in `period` relative to day `today`. Weeks start on Sunday, as in Excel. */
export function inTimePeriod(period: TimePeriod, d: number, today: number, date1904: boolean): boolean {
  const day = Math.floor(d);
  const t = Math.floor(today);
  const weekStart = t - serialToDate(t, date1904).getUTCDay();
  const monthIndex = (s: number): number => {
    const date = serialToDate(s, date1904);
    return date.getUTCFullYear() * 12 + date.getUTCMonth();
  };
  switch (period) {
    case 'today':
      return day === t;
    case 'yesterday':
      return day === t - 1;
    case 'tomorrow':
      return day === t + 1;
    case 'last7Days':
      return day >= t - 6 && day <= t;
    case 'thisWeek':
      return day >= weekStart && day <= weekStart + 6;
    case 'lastWeek':
      return day >= weekStart - 7 && day < weekStart;
    case 'nextWeek':
      return day >= weekStart + 7 && day <= weekStart + 13;
    case 'thisMonth':
      return monthIndex(day) === monthIndex(t);
    case 'lastMonth':
      return monthIndex(day) === monthIndex(t) - 1;
    case 'nextMonth':
      return monthIndex(day) === monthIndex(t) + 1;
  }
}

// ---- colours ----------------------------------------------------------------

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '');
  return [Number.parseInt(h.slice(0, 2), 16), Number.parseInt(h.slice(2, 4), 16), Number.parseInt(h.slice(4, 6), 16)];
}

function lerpColor(a: string, b: string, t: number): string {
  const [r1, g1, b1] = hexToRgb(a);
  const [r2, g2, b2] = hexToRgb(b);
  const mix = (x: number, y: number) =>
    Math.round(x + (y - x) * t)
      .toString(16)
      .padStart(2, '0');
  return `#${mix(r1, r2)}${mix(g1, g2)}${mix(b1, b2)}`.toUpperCase();
}

// ---- compiled rules ---------------------------------------------------------

type RuleEval = (row: number, col: number, v: CalcScalar) => CellOverlay | null;

interface CompiledRule {
  readonly ranges: readonly Range[];
  readonly stopIfTrue: boolean;
  readonly evaluate: RuleEval;
}

interface Context {
  readonly ws: Worksheet;
  readonly calc: CalcEngine;
  readonly palette: ThemePalette;
  readonly dxfs: ReadonlyArray<DifferentialStyle>;
  readonly valueAt: (row: number, col: number) => CalcScalar;
}

const BUCKET_ROWS = 256;
const MEMO_LIMIT = 200_000;

function stripEq(f: string): string {
  return f.startsWith('=') ? f.slice(1) : f;
}

function compileRule(ctx: Context, rule: ConditionalFormattingRule, ranges: readonly Range[]): CompiledRule | undefined {
  const first = ranges[0];
  if (!first) return undefined;
  const sheet = ctx.ws.title;
  // Relative references in a rule's formulas are written for the top-left
  // cell of the first sqref range and shift with the cell being tested.
  const formulaAt = (f: string, row: number, col: number): CalcScalar =>
    ctx.calc.evaluate(translateFormula(stripEq(f), row - first.r1, col - first.c1), sheet, row, col);

  let statsCache: Stats | undefined;
  const stats = (): Stats => {
    if (statsCache) return statsCache;
    const values: number[] = [];
    for (const r of ranges) {
      forEachCellInRange(ctx.ws, r, (cell) => {
        const v = ctx.valueAt(cell.row, cell.col);
        if (typeof v === 'number') values.push(v);
      });
    }
    const sorted = Float64Array.from(values).sort();
    let sum = 0;
    for (const v of sorted) sum += v;
    const mean = sorted.length > 0 ? sum / sorted.length : 0;
    let sq = 0;
    for (const v of sorted) sq += (v - mean) ** 2;
    statsCache = { sorted, mean, stdDev: sorted.length > 1 ? Math.sqrt(sq / (sorted.length - 1)) : 0 };
    return statsCache;
  };

  let dupCache: Map<string, number> | undefined;
  const dupCounts = (): Map<string, number> => {
    if (dupCache) return dupCache;
    const counts = new Map<string, number>();
    for (const r of ranges) {
      forEachCellInRange(ctx.ws, r, (cell) => {
        const k = dupKey(ctx.valueAt(cell.row, cell.col));
        if (k !== undefined) counts.set(k, (counts.get(k) ?? 0) + 1);
      });
    }
    dupCache = counts;
    return counts;
  };

  const cfvoValue = (c: CfvoSpec): number | undefined => {
    const s = stats();
    const min = s.sorted[0] ?? 0;
    const max = s.sorted[s.sorted.length - 1] ?? 0;
    const num = (): number | undefined => {
      if (c.val === undefined) return undefined;
      const n = Number(c.val);
      if (c.val.trim() !== '' && Number.isFinite(n)) return n;
      const v = formulaAt(c.val, first.r1, first.c1);
      return typeof v === 'number' ? v : undefined;
    };
    switch (c.type) {
      case 'min':
        return min;
      case 'max':
        return max;
      case 'num':
      case 'formula':
        return num();
      case 'percent': {
        const p = num();
        return p === undefined ? undefined : min + ((max - min) * p) / 100;
      }
      case 'percentile': {
        const p = num();
        return p === undefined ? undefined : percentileInc(s.sorted, p / 100);
      }
    }
  };

  const fmt = dxfOverlay(rule.dxfId === undefined ? undefined : ctx.dxfs[rule.dxfId], ctx.palette);
  const when = (test: (row: number, col: number, v: CalcScalar) => boolean): RuleEval => (row, col, v) => (test(row, col, v) ? fmt : null);
  const visual = parseVisualRule(rule);

  let evaluate: RuleEval | undefined;
  switch (rule.type) {
    case 'cellIs': {
      const [f1, f2] = rule.formulas;
      if (f1 === undefined) return undefined;
      const op = rule.operator ?? 'equal';
      evaluate = when((row, col, v) => {
        if (isError(v)) return false;
        const c1 = compare(v, formulaAt(f1, row, col));
        if (c1 === undefined) return false;
        switch (op) {
          case 'equal':
            return c1 === 0;
          case 'notEqual':
            return c1 !== 0;
          case 'greaterThan':
            return c1 > 0;
          case 'greaterThanOrEqual':
            return c1 >= 0;
          case 'lessThan':
            return c1 < 0;
          case 'lessThanOrEqual':
            return c1 <= 0;
          case 'between':
          case 'notBetween': {
            if (f2 === undefined) return false;
            const c2 = compare(v, formulaAt(f2, row, col));
            if (c2 === undefined) return false;
            // Excel accepts the bounds in either order.
            const lo = compare(formulaAt(f1, row, col), formulaAt(f2, row, col)) ?? 0;
            const inside = lo <= 0 ? c1 >= 0 && c2 <= 0 : c1 <= 0 && c2 >= 0;
            return op === 'between' ? inside : !inside;
          }
        }
        return false;
      });
      break;
    }
    case 'expression': {
      const f = rule.formulas[0];
      if (f === undefined) return undefined;
      evaluate = when((row, col) => truthy(formulaAt(f, row, col)));
      break;
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith': {
      const needle = (rule.text ?? '').toLowerCase();
      const type = rule.type;
      evaluate = when((_r, _c, v) => {
        if (isError(v)) return false;
        const hay = textOf(v).toLowerCase();
        if (type === 'containsText') return hay.includes(needle);
        if (type === 'notContainsText') return !hay.includes(needle);
        if (type === 'beginsWith') return hay.startsWith(needle);
        return hay.endsWith(needle);
      });
      break;
    }
    case 'containsBlanks':
      evaluate = when((_r, _c, v) => isBlankText(v));
      break;
    case 'notContainsBlanks':
      evaluate = when((_r, _c, v) => !isBlankText(v));
      break;
    case 'containsErrors':
      evaluate = when((_r, _c, v) => isError(v));
      break;
    case 'notContainsErrors':
      evaluate = when((_r, _c, v) => !isError(v));
      break;
    case 'timePeriod': {
      const period = rule.timePeriod;
      if (!period) return undefined;
      const today = ctx.calc.now();
      const date1904 = ctx.calc.date1904;
      evaluate = when((_r, _c, v) => typeof v === 'number' && inTimePeriod(period, v, today, date1904));
      break;
    }
    case 'duplicateValues':
    case 'uniqueValues': {
      const wantDup = rule.type === 'duplicateValues';
      evaluate = when((_r, _c, v) => {
        const k = dupKey(v);
        if (k === undefined) return false;
        return (dupCounts().get(k) ?? 0) > 1 === wantDup;
      });
      break;
    }
    case 'top10': {
      let threshold: number | undefined;
      const bottom = rule.bottom === true;
      const thresholdOf = (): number | undefined => {
        if (threshold !== undefined) return threshold;
        const s = stats().sorted;
        if (s.length === 0) return undefined;
        const rank = rule.rank ?? 10;
        const n = Math.min(s.length, Math.max(1, rule.percent ? Math.floor((s.length * rank) / 100) : rank));
        threshold = bottom ? s[n - 1] : s[s.length - n];
        return threshold;
      };
      evaluate = when((_r, _c, v) => {
        if (typeof v !== 'number') return false;
        const th = thresholdOf();
        return th !== undefined && (bottom ? v <= th : v >= th);
      });
      break;
    }
    case 'aboveAverage': {
      const above = rule.aboveAverage !== false;
      const equal = rule.equalAverage === true;
      const k = rule.stdDev ?? 0;
      evaluate = when((_r, _c, v) => {
        if (typeof v !== 'number') return false;
        const s = stats();
        if (s.sorted.length === 0) return false;
        const limit = s.mean + (above ? 1 : -1) * k * s.stdDev;
        if (equal && v === limit) return true;
        return above ? v > limit : v < limit;
      });
      break;
    }
    case 'colorScale': {
      if (visual?.kind !== 'colorScale') return undefined;
      const colors = visual.colors.map((c) => resolveColor(c, ctx.palette, '#FFFFFF'));
      let points: number[] | undefined;
      evaluate = (_r, _c, v) => {
        if (typeof v !== 'number') return null;
        points ??= visual.cfvos.map((c) => cfvoValue(c) ?? 0);
        const pts = points;
        const last = pts.length - 1;
        if (v <= (pts[0] ?? 0)) return { fill: colors[0] ?? '#FFFFFF' };
        if (v >= (pts[last] ?? 0)) return { fill: colors[last] ?? '#FFFFFF' };
        for (let i = 0; i < last; i++) {
          const a = pts[i] ?? 0;
          const b = pts[i + 1] ?? 0;
          if (v <= b) return { fill: lerpColor(colors[i] ?? '#FFFFFF', colors[i + 1] ?? '#FFFFFF', b === a ? 1 : (v - a) / (b - a)) };
        }
        return { fill: colors[last] ?? '#FFFFFF' };
      };
      break;
    }
    case 'dataBar': {
      if (visual?.kind !== 'dataBar') return undefined;
      const color = resolveColor(visual.color, ctx.palette, '#638EC6');
      const minLen = visual.minLength / 100;
      const maxLen = visual.maxLength / 100;
      let bounds: [number, number] | undefined;
      evaluate = (_r, _c, v) => {
        if (typeof v !== 'number') return null;
        if (!bounds) {
          const minCfvo = visual.cfvos[0] ?? { type: 'min' };
          const maxCfvo = visual.cfvos[1] ?? { type: 'max' };
          const lo = cfvoValue(minCfvo) ?? 0;
          const hi = cfvoValue(maxCfvo) ?? 0;
          bounds = visual.extended ? [minCfvo.type === 'min' ? Math.min(0, lo) : lo, maxCfvo.type === 'max' ? Math.max(0, hi) : hi] : [lo, hi];
        }
        const [lo, hi] = bounds;
        const hide = visual.showValue ? {} : { hideValue: true };
        if (lo < 0 && hi > 0) {
          // Mixed signs: bars grow from an axis placed at zero, negatives in red.
          const axis = -lo / (hi - lo);
          if (v >= 0) return { bar: { start: axis, end: axis + (1 - axis) * Math.min(1, v / hi), color }, ...hide };
          return { bar: { start: axis - axis * Math.min(1, v / lo), end: axis, color: '#FF0000' }, ...hide };
        }
        const t = hi === lo ? 1 : Math.min(1, Math.max(0, (v - lo) / (hi - lo)));
        return { bar: { start: 0, end: minLen + (maxLen - minLen) * t, color }, ...hide };
      };
      break;
    }
    case 'iconSet': {
      if (visual?.kind !== 'iconSet') return undefined;
      const icons = ICON_SETS[visual.iconSet] ?? ICON_SETS['3TrafficLights1'] ?? [];
      let thresholds: number[] | undefined;
      evaluate = (_r, _c, v) => {
        if (typeof v !== 'number') return null;
        thresholds ??= visual.cfvos.map((c) => cfvoValue(c) ?? 0);
        let idx = 0;
        for (let i = 1; i < thresholds.length; i++) {
          const th = thresholds[i] ?? 0;
          const strict = visual.cfvos[i]?.gte === false;
          if (strict ? v > th : v >= th) idx = i;
        }
        const pick = icons[visual.reverse ? icons.length - 1 - idx : idx];
        if (!pick) return null;
        return { icon: pick, ...(visual.showValue ? {} : { hideValue: true }) };
      };
      break;
    }
  }
  if (!evaluate) return undefined;
  return { ranges, stopIfTrue: rule.stopIfTrue === true, evaluate };
}

/** Lower-priority rules only fill what higher-priority rules left unset. */
function mergeInto(target: MutableOverlay, add: CellOverlay): void {
  if (target.fill === undefined && add.fill !== undefined) target.fill = add.fill;
  if (target.color === undefined && add.color !== undefined) target.color = add.color;
  if (target.bold === undefined && add.bold !== undefined) target.bold = add.bold;
  if (target.italic === undefined && add.italic !== undefined) target.italic = add.italic;
  if (target.strike === undefined && add.strike !== undefined) target.strike = add.strike;
  if (target.underline === undefined && add.underline !== undefined) target.underline = add.underline;
  if (target.bar === undefined && add.bar !== undefined) target.bar = add.bar;
  if (target.icon === undefined && add.icon !== undefined) target.icon = add.icon;
  if (add.hideValue) target.hideValue = true;
}

export function sqrefRanges(cf: ConditionalFormatting): Range[] {
  return cf.sqref.ranges.map(fromBoundaries);
}

/** Build the per-cell overlay for the sheet's rules; undefined when the sheet has none. */
export function conditionalOverlay(wb: Workbook, ws: Worksheet, calc: CalcEngine, styles: StyleResolver): OverlayFn | undefined {
  if (ws.conditionalFormatting.length === 0) return undefined;
  const sheet = ws.title;
  const ctx: Context = {
    ws,
    calc,
    palette: styles.palette,
    dxfs: dxfsOf(wb),
    valueAt: (row, col) => calc.cellValue(sheet, row, col),
  };
  const entries: Array<{ priority: number; rule: ConditionalFormattingRule; ranges: Range[] }> = [];
  for (const cf of ws.conditionalFormatting) {
    const ranges = sqrefRanges(cf);
    for (const rule of cf.rules) entries.push({ priority: rule.priority, rule, ranges });
  }
  entries.sort((a, b) => a.priority - b.priority);
  const compiled: CompiledRule[] = [];
  for (const e of entries) {
    const c = compileRule(ctx, e.rule, e.ranges);
    if (c) compiled.push(c);
  }
  if (compiled.length === 0) return undefined;

  const buckets = new Map<number, CompiledRule[]>();
  const rulesForBucket = (b: number): CompiledRule[] => {
    let list = buckets.get(b);
    if (!list) {
      const band: Range = { r1: b * BUCKET_ROWS + 1, r2: (b + 1) * BUCKET_ROWS, c1: 1, c2: MAX_COL };
      list = compiled.filter((c) => c.ranges.some((r) => rangesIntersect(r, band)));
      buckets.set(b, list);
    }
    return list;
  };

  const memo = new Map<number, CellOverlay | null>();
  return (row, col) => {
    const rules = rulesForBucket(Math.floor((row - 1) / BUCKET_ROWS));
    if (rules.length === 0) return undefined;
    const key = row * (MAX_COL + 1) + col;
    const hit = memo.get(key);
    if (hit !== undefined) return hit ?? undefined;
    let out: MutableOverlay | null = null;
    let value: CalcScalar | undefined;
    for (const rule of rules) {
      if (!rule.ranges.some((r) => inRange(r, row, col))) continue;
      value ??= ctx.valueAt(row, col);
      const res = rule.evaluate(row, col, value);
      if (!res) continue;
      out ??= {};
      mergeInto(out, res);
      if (rule.stopIfTrue) break;
    }
    if (memo.size > MEMO_LIMIT) memo.clear();
    memo.set(key, out);
    return out ?? undefined;
  };
}

// ---- editing ------------------------------------------------------------------

/** One rule with the cells it applies to, the unit the Rules Manager edits. */
export interface RuleEntry {
  readonly sqref: string;
  readonly rule: ConditionalFormattingRule;
}

/** Every rule on the sheet, highest priority (lowest number) first. */
export function listRules(ws: Worksheet): RuleEntry[] {
  const out: RuleEntry[] = [];
  for (const cf of ws.conditionalFormatting) {
    const sqref = sqrefRanges(cf).map((r) => rangeAddress(r)).join(' ');
    for (const rule of cf.rules) out.push({ sqref, rule });
  }
  return out.sort((a, b) => a.rule.priority - b.rule.priority);
}

/** Parse an "Applies to" text like `=$A$1:$B$5,$D:$D` into ranges; undefined when malformed. */
export function parseAppliesTo(text: string): Range[] | undefined {
  const pieces = text
    .replace(/^=/, '')
    .split(/[\s,]+/)
    .filter((p) => p.length > 0);
  if (pieces.length === 0) return undefined;
  const out: Range[] = [];
  for (const piece of pieces) {
    const parsed = parseRangeAddress(piece);
    if (!parsed) return undefined;
    out.push(parsed.range);
  }
  return out;
}

function blockFor(ranges: readonly Range[], rules: ConditionalFormattingRule[]): ConditionalFormatting {
  return makeConditionalFormatting({ sqref: { ranges: ranges.map(toBoundaries) }, rules });
}

/** A rule as the dialogs describe it; `priority` is assigned on insert. */
export type NewRule = Omit<Parameters<typeof makeCfRule>[0], 'priority' | 'dxfId'>;

/**
 * Add `rule` over `ranges` as the sheet's highest-priority rule (Excel puts a
 * new rule on top), registering `format` as its DXF. One undo step.
 */
export function addRule(doc: SpreadsheetEditor, ranges: readonly Range[], rule: NewRule, format?: CfFormat): void {
  const ws = doc.ws;
  doc.transact('Conditional Formatting', (tx) => {
    tx.sheet(ws, 'conditionalFormatting');
    for (const cf of ws.conditionalFormatting) for (const r of cf.rules) r.priority += 1;
    const dxfId = format ? registerCfFormat(doc.wb, format) : undefined;
    const made = makeCfRule({ ...rule, priority: 1, ...(dxfId !== undefined ? { dxfId } : {}) });
    ws.conditionalFormatting.push(blockFor(ranges, [made]));
  });
}

/** `outer` minus `hole` as up to four disjoint rectangles. */
export function subtractRange(outer: Range, hole: Range): Range[] {
  if (!rangesIntersect(outer, hole)) return [outer];
  const out: Range[] = [];
  if (hole.r1 > outer.r1) out.push({ r1: outer.r1, r2: hole.r1 - 1, c1: outer.c1, c2: outer.c2 });
  if (hole.r2 < outer.r2) out.push({ r1: hole.r2 + 1, r2: outer.r2, c1: outer.c1, c2: outer.c2 });
  const r1 = Math.max(outer.r1, hole.r1);
  const r2 = Math.min(outer.r2, hole.r2);
  if (hole.c1 > outer.c1) out.push({ r1, r2, c1: outer.c1, c2: hole.c1 - 1 });
  if (hole.c2 < outer.c2) out.push({ r1, r2, c1: hole.c2 + 1, c2: outer.c2 });
  return out;
}

/** Clear Rules ▸ from Selected Cells: cut `ranges` out of every rule's sqref, dropping rules left empty. */
export function clearRulesIn(doc: SpreadsheetEditor, ranges: readonly Range[]): void {
  const ws = doc.ws;
  if (!ws.conditionalFormatting.some((cf) => sqrefRanges(cf).some((r) => ranges.some((h) => rangesIntersect(r, h))))) return;
  doc.transact('Clear Rules', (tx) => {
    tx.sheet(ws, 'conditionalFormatting');
    const next: ConditionalFormatting[] = [];
    for (const cf of ws.conditionalFormatting) {
      let parts = sqrefRanges(cf);
      for (const hole of ranges) parts = parts.flatMap((p) => subtractRange(p, hole));
      if (parts.length > 0) next.push({ ...cf, sqref: { ranges: parts.map(toBoundaries) } });
    }
    ws.conditionalFormatting = next;
  });
}

export function clearSheetRules(doc: SpreadsheetEditor): void {
  const ws = doc.ws;
  if (ws.conditionalFormatting.length === 0) return;
  doc.transact('Clear Rules', (tx) => {
    tx.sheet(ws, 'conditionalFormatting');
    ws.conditionalFormatting = [];
  });
}

/** A Rules Manager row: the rule plus its edited "Applies to" and optional new format. */
export interface DraftRule {
  readonly ranges: readonly Range[];
  readonly rule: ConditionalFormattingRule;
  readonly format?: CfFormat;
}

/**
 * Replace every rule on the sheet with `drafts`, in order: priorities are
 * renumbered from 1 and rules sharing an identical sqref share one block, as
 * Excel writes them. One undo step.
 */
export function replaceRules(doc: SpreadsheetEditor, drafts: readonly DraftRule[]): void {
  const ws = doc.ws;
  doc.transact('Manage Rules', (tx) => {
    tx.sheet(ws, 'conditionalFormatting');
    const blocks = new Map<string, ConditionalFormatting>();
    const order: ConditionalFormatting[] = [];
    drafts.forEach((d, i) => {
      const dxfId = d.format ? registerCfFormat(doc.wb, d.format) : d.rule.dxfId;
      const rule = makeCfRule({ ...d.rule, priority: i + 1, ...(dxfId !== undefined ? { dxfId } : {}) });
      const key = d.ranges.map((r) => rangeAddress(r)).join(' ');
      const block = blocks.get(key);
      if (block) block.rules.push(rule);
      else {
        const made = blockFor(d.ranges, [rule]);
        blocks.set(key, made);
        order.push(made);
      }
    });
    ws.conditionalFormatting = order;
  });
}

// ---- rule forms (New / Edit Formatting Rule) -----------------------------------

export type RuleFormKind = 'values' | 'contains' | 'topBottom' | 'average' | 'unique' | 'formula';
export type ValuesStyle = 'scale2' | 'scale3' | 'dataBar' | 'iconSet';
export type ContainsKind = 'cellValue' | 'text' | 'date' | 'blanks' | 'noBlanks' | 'errors' | 'noErrors';
export type AverageChoice = 'above' | 'below' | 'equalAbove' | 'equalBelow' | 'std1Above' | 'std1Below' | 'std2Above' | 'std2Below' | 'std3Above' | 'std3Below';

export interface CfvoForm {
  type: CfvoType;
  val: string;
  /** RRGGBB, no '#'. */
  color: string;
}

/** Flat, bindable state behind the rule editor; only the fields of the chosen kind matter. */
export interface RuleForm {
  kind: RuleFormKind;
  style: ValuesStyle;
  min: CfvoForm;
  mid: CfvoForm;
  max: CfvoForm;
  barColor: string;
  barOnly: boolean;
  iconSet: string;
  reverse: boolean;
  iconOnly: boolean;
  contains: ContainsKind;
  operator: CellIsOperator;
  value1: string;
  value2: string;
  textOp: TextOperator;
  text: string;
  period: TimePeriod;
  bottom: boolean;
  rank: number;
  percent: boolean;
  average: AverageChoice;
  unique: boolean;
  formula: string;
}

export function defaultRuleForm(): RuleForm {
  return {
    kind: 'values',
    style: 'scale2',
    min: { type: 'min', val: '', color: 'F8696B' },
    mid: { type: 'percentile', val: '50', color: 'FFEB84' },
    max: { type: 'max', val: '', color: '63BE7B' },
    barColor: '638EC6',
    barOnly: false,
    iconSet: '3TrafficLights1',
    reverse: false,
    iconOnly: false,
    contains: 'cellValue',
    operator: 'between',
    value1: '',
    value2: '',
    textOp: 'containsText',
    text: '',
    period: 'yesterday',
    bottom: false,
    rank: 10,
    percent: false,
    average: 'above',
    unique: false,
    formula: '',
  };
}

const CELL_IS_OPERATORS: ReadonlySet<string> = new Set<CellIsOperator>(['lessThan', 'lessThanOrEqual', 'equal', 'notEqual', 'greaterThanOrEqual', 'greaterThan', 'between', 'notBetween']);

function isCellIsOperator(op: string | undefined): op is CellIsOperator {
  return op !== undefined && CELL_IS_OPERATORS.has(op);
}

/** What a user typed as a comparison value, as formula text: `=A1` → `A1`, `5` → `5`, `abc` → `"abc"`. */
export function operandFormula(input: string): string {
  const s = input.trim();
  if (s.startsWith('=')) return toStorageFormula(s.slice(1));
  if (s !== '' && Number.isFinite(Number(s))) return s;
  return `"${s.replaceAll('"', '""')}"`;
}

/** Inverse of {@link operandFormula} for showing a stored operand in the editor. */
export function operandText(formula: string): string {
  const m = /^"((?:[^"]|"")*)"$/.exec(formula);
  if (m?.[1] !== undefined) return m[1].replaceAll('""', '"');
  if (formula !== '' && Number.isFinite(Number(formula))) return formula;
  return `=${fromStorageFormula(formula)}`;
}

const quote = (text: string): string => `"${text.replaceAll('"', '""')}"`;

/**
 * The formulas Excel writes alongside text / blank / error / date rules. The
 * rule type alone decides how the editor evaluates them, but Excel reads the
 * formula, so a saved file needs it. `a1` is the top-left cell of the range.
 */
function companionFormula(rule: { type: string; text?: string; timePeriod?: TimePeriod }, a1: string): string | undefined {
  const t = quote(rule.text ?? '');
  switch (rule.type) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH(${t},${a1})))`;
    case 'notContainsText':
      return `ISERROR(SEARCH(${t},${a1}))`;
    case 'beginsWith':
      return `LEFT(${a1},LEN(${t}))=${t}`;
    case 'endsWith':
      return `RIGHT(${a1},LEN(${t}))=${t}`;
    case 'containsBlanks':
      return `LEN(TRIM(${a1}))=0`;
    case 'notContainsBlanks':
      return `LEN(TRIM(${a1}))>0`;
    case 'containsErrors':
      return `ISERROR(${a1})`;
    case 'notContainsErrors':
      return `NOT(ISERROR(${a1}))`;
    case 'timePeriod':
      switch (rule.timePeriod) {
        case 'today':
          return `FLOOR(${a1},1)=TODAY()`;
        case 'yesterday':
          return `FLOOR(${a1},1)=TODAY()-1`;
        case 'tomorrow':
          return `FLOOR(${a1},1)=TODAY()+1`;
        case 'last7Days':
          return `AND(TODAY()-FLOOR(${a1},1)<=6,FLOOR(${a1},1)<=TODAY())`;
        case 'lastWeek':
          return `AND(TODAY()-ROUNDDOWN(${a1},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${a1},0)<(WEEKDAY(TODAY())+7))`;
        case 'thisWeek':
          return `AND(TODAY()-ROUNDDOWN(${a1},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${a1},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
        case 'nextWeek':
          return `AND(ROUNDDOWN(${a1},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${a1},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
        case 'lastMonth':
          return `AND(MONTH(${a1})=MONTH(EDATE(TODAY(),0-1)),YEAR(${a1})=YEAR(EDATE(TODAY(),0-1)))`;
        case 'thisMonth':
          return `AND(MONTH(${a1})=MONTH(TODAY()),YEAR(${a1})=YEAR(TODAY()))`;
        case 'nextMonth':
          return `AND(MONTH(${a1})=MONTH(EDATE(TODAY(),0+1)),YEAR(${a1})=YEAR(EDATE(TODAY(),0+1)))`;
        case undefined:
          return undefined;
      }
  }
  return undefined;
}

/** A text / blank / error / date rule with the companion formula Excel expects. */
export function withCompanionFormula(rule: NewRule, a1: string): NewRule {
  const f = companionFormula(rule, a1);
  return f === undefined ? rule : { ...rule, formulas: [f] };
}

const cfvoOf = (c: CfvoForm): CfvoSpec => (c.type === 'min' || c.type === 'max' ? { type: c.type } : { type: c.type, val: c.val.trim().replace(/^=/, '') || '0' });

const TEXT_RULE_TYPE: Readonly<Record<TextOperator, NewRule['type']>> = {
  containsText: 'containsText',
  notContains: 'notContainsText',
  beginsWith: 'beginsWith',
  endsWith: 'endsWith',
};

export const AVERAGE_CHOICES: readonly AverageChoice[] = ['above', 'below', 'equalAbove', 'equalBelow', 'std1Above', 'std1Below', 'std2Above', 'std2Below', 'std3Above', 'std3Below'];

const AVERAGE_FIELDS: Readonly<Record<AverageChoice, Pick<NewRule, 'aboveAverage' | 'equalAverage' | 'stdDev'>>> = {
  above: {},
  below: { aboveAverage: false },
  equalAbove: { equalAverage: true },
  equalBelow: { aboveAverage: false, equalAverage: true },
  std1Above: { stdDev: 1 },
  std1Below: { aboveAverage: false, stdDev: 1 },
  std2Above: { stdDev: 2 },
  std2Below: { aboveAverage: false, stdDev: 2 },
  std3Above: { stdDev: 3 },
  std3Below: { aboveAverage: false, stdDev: 3 },
};

/** Whether rules of this form carry a DXF (visual rules draw their own formatting). */
export function formUsesFormat(form: RuleForm): boolean {
  return form.kind !== 'values';
}

/** The library rule for an editor form; `a1` is the top-left cell of the range it applies to. */
export function ruleFromForm(form: RuleForm, a1: string): NewRule {
  switch (form.kind) {
    case 'values': {
      const argbOf = (hex: string): Color => ({ rgb: `FF${hex.toUpperCase()}` });
      let visual: VisualRule;
      if (form.style === 'scale2' || form.style === 'scale3') {
        const stops = form.style === 'scale3' ? [form.min, form.mid, form.max] : [form.min, form.max];
        visual = { kind: 'colorScale', cfvos: stops.map(cfvoOf), colors: stops.map((s) => argbOf(s.color)) };
      } else if (form.style === 'dataBar') {
        visual = { kind: 'dataBar', cfvos: [cfvoOf(form.min), cfvoOf(form.max)], color: argbOf(form.barColor), showValue: !form.barOnly, minLength: 10, maxLength: 90, extended: false };
      } else {
        visual = { kind: 'iconSet', iconSet: form.iconSet, cfvos: defaultIconCfvos(iconCount(form.iconSet)), reverse: form.reverse, showValue: !form.iconOnly };
      }
      const type = visual.kind;
      return { type, formulas: [], innerXml: visualRuleXml(visual) };
    }
    case 'contains':
      switch (form.contains) {
        case 'cellValue': {
          const two = form.operator === 'between' || form.operator === 'notBetween';
          return { type: 'cellIs', operator: form.operator, formulas: two ? [operandFormula(form.value1), operandFormula(form.value2)] : [operandFormula(form.value1)] };
        }
        case 'text':
          return withCompanionFormula({ type: TEXT_RULE_TYPE[form.textOp], operator: form.textOp, text: form.text, formulas: [] }, a1);
        case 'date':
          return withCompanionFormula({ type: 'timePeriod', timePeriod: form.period, formulas: [] }, a1);
        case 'blanks':
          return withCompanionFormula({ type: 'containsBlanks', formulas: [] }, a1);
        case 'noBlanks':
          return withCompanionFormula({ type: 'notContainsBlanks', formulas: [] }, a1);
        case 'errors':
          return withCompanionFormula({ type: 'containsErrors', formulas: [] }, a1);
        case 'noErrors':
          return withCompanionFormula({ type: 'notContainsErrors', formulas: [] }, a1);
      }
      break;
    case 'topBottom':
      return { type: 'top10', rank: Math.max(1, Math.round(form.rank)), formulas: [], ...(form.bottom ? { bottom: true } : {}), ...(form.percent ? { percent: true } : {}) };
    case 'average':
      return { type: 'aboveAverage', formulas: [], ...AVERAGE_FIELDS[form.average] };
    case 'unique':
      return { type: form.unique ? 'uniqueValues' : 'duplicateValues', formulas: [] };
    case 'formula':
      return { type: 'expression', formulas: [toStorageFormula(form.formula.trim().replace(/^=/, ''))] };
  }
  return { type: 'expression', formulas: ['FALSE'] };
}

const hexOf = (c: Color | undefined, palette: ThemePalette): string => resolveColor(c, palette, '#000000').slice(1);

/** Load an existing rule into the editor form. */
export function formFromRule(rule: ConditionalFormattingRule, palette: ThemePalette): RuleForm {
  const form = defaultRuleForm();
  const visual = parseVisualRule(rule);
  const cfvoForm = (c: CfvoSpec | undefined, color: Color | undefined, fallback: CfvoForm): CfvoForm =>
    c ? { type: c.type, val: c.val ?? '', color: color ? hexOf(color, palette) : fallback.color } : fallback;
  switch (rule.type) {
    case 'colorScale':
      form.kind = 'values';
      if (visual?.kind === 'colorScale') {
        const three = visual.cfvos.length === 3;
        form.style = three ? 'scale3' : 'scale2';
        form.min = cfvoForm(visual.cfvos[0], visual.colors[0], form.min);
        if (three) form.mid = cfvoForm(visual.cfvos[1], visual.colors[1], form.mid);
        form.max = cfvoForm(visual.cfvos[three ? 2 : 1], visual.colors[three ? 2 : 1], form.max);
      }
      break;
    case 'dataBar':
      form.kind = 'values';
      form.style = 'dataBar';
      if (visual?.kind === 'dataBar') {
        form.min = cfvoForm(visual.cfvos[0], undefined, form.min);
        form.max = cfvoForm(visual.cfvos[1], undefined, form.max);
        form.barColor = hexOf(visual.color, palette);
        form.barOnly = !visual.showValue;
      }
      break;
    case 'iconSet':
      form.kind = 'values';
      form.style = 'iconSet';
      if (visual?.kind === 'iconSet') {
        form.iconSet = visual.iconSet;
        form.reverse = visual.reverse;
        form.iconOnly = !visual.showValue;
      }
      break;
    case 'cellIs':
      form.kind = 'contains';
      form.contains = 'cellValue';
      form.operator = isCellIsOperator(rule.operator) ? rule.operator : 'equal';
      form.value1 = operandText(rule.formulas[0] ?? '');
      form.value2 = operandText(rule.formulas[1] ?? '');
      break;
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      form.kind = 'contains';
      form.contains = 'text';
      form.textOp = rule.type === 'notContainsText' ? 'notContains' : rule.type;
      form.text = rule.text ?? '';
      break;
    case 'timePeriod':
      form.kind = 'contains';
      form.contains = 'date';
      form.period = rule.timePeriod ?? 'today';
      break;
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      form.kind = 'contains';
      form.contains = ({ containsBlanks: 'blanks', notContainsBlanks: 'noBlanks', containsErrors: 'errors', notContainsErrors: 'noErrors' } as const)[rule.type];
      break;
    case 'top10':
      form.kind = 'topBottom';
      form.bottom = rule.bottom === true;
      form.percent = rule.percent === true;
      form.rank = rule.rank ?? 10;
      break;
    case 'aboveAverage': {
      form.kind = 'average';
      const above = rule.aboveAverage !== false;
      const sd = rule.stdDev ?? 0;
      const match = AVERAGE_CHOICES.find((k) => {
        const f = AVERAGE_FIELDS[k];
        return (f.aboveAverage ?? true) === above && (f.equalAverage ?? false) === (rule.equalAverage === true && sd === 0) && (f.stdDev ?? 0) === sd;
      });
      form.average = match ?? (above ? 'above' : 'below');
      break;
    }
    case 'duplicateValues':
    case 'uniqueValues':
      form.kind = 'unique';
      form.unique = rule.type === 'uniqueValues';
      break;
    case 'expression':
      form.kind = 'formula';
      form.formula = `=${fromStorageFormula(rule.formulas[0] ?? '')}`;
      break;
  }
  return form;
}
