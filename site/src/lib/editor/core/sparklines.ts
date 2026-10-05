// Insert ▸ Sparklines and the Sparkline tab. A sparkline group owns the type,
// colours and Show flags of every sparkline in it, so the tab edits the group
// of the active cell's sparkline. Every edit is one undo step on the sheet's
// `sparklineGroups`.

import type { Color } from '@office-kit/xlsx/styles';
import { makeSparklineGroup, type Sparkline, type SparklineGroup, type SparklineType, type Worksheet } from '@office-kit/xlsx/worksheet';
import type { CalcEngine } from '../calc/index.ts';
import { inRange, parseRangeAddress, quoteSheetName, rangeAddress, rangeOf, type CellPos, type Range } from './address.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';

export interface SparklineCell {
  readonly groupIndex: number;
  readonly group: SparklineGroup;
  readonly sparkline: Sparkline;
  readonly row: number;
  readonly col: number;
}

/** Every sparkline on `ws`, keyed `row:col` by the cell it draws in. */
export function sparklineIndex(ws: Worksheet): Map<string, SparklineCell> {
  const out = new Map<string, SparklineCell>();
  (ws.sparklineGroups ?? []).forEach((group, groupIndex) => {
    for (const sparkline of group.sparklines) {
      const at = parseRangeAddress(sparkline.location);
      if (!at) continue;
      out.set(`${at.range.r1}:${at.range.c1}`, { groupIndex, group, sparkline, row: at.range.r1, col: at.range.c1 });
    }
  });
  return out;
}

/**
 * The plotted values, in range order. Blank and text cells are null: Excel
 * leaves a gap for them (or zero / a connecting span, per the group setting).
 */
export function sparklineValues(calc: CalcEngine, defaultSheet: string, formula: string): Array<number | null> {
  const ref = parseRangeAddress(formula);
  if (!ref) return [];
  const sheet = ref.sheet ?? defaultSheet;
  const { r1, c1, r2, c2 } = ref.range;
  const out: Array<number | null> = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const v = calc.cellValue(sheet, r, c);
      out.push(typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : null);
    }
  }
  return out;
}

export type SparklineError = 'sparklineBadData' | 'sparklineBadLocation' | 'sparklineSizeMismatch';

/**
 * Pair a data range with a one-row or one-column location range the way
 * Excel's Create Sparklines dialog does: each location cell takes the data
 * row (or column) lined up with it.
 */
export function pairSparklines(dataText: string, locationText: string, sheet: string): Sparkline[] | SparklineError {
  const data = parseRangeAddress(dataText);
  if (!data) return 'sparklineBadData';
  const loc = parseRangeAddress(locationText);
  if (!loc || (loc.sheet !== undefined && loc.sheet !== sheet)) return 'sparklineBadLocation';
  const L = loc.range;
  const D = data.range;
  if (L.r1 !== L.r2 && L.c1 !== L.c2) return 'sparklineBadLocation';
  const prefix = `${quoteSheetName(data.sheet ?? sheet)}!`;
  const cells: CellPos[] = [];
  for (let r = L.r1; r <= L.r2; r++) for (let c = L.c1; c <= L.c2; c++) cells.push({ row: r, col: c });
  const rows = D.r2 - D.r1 + 1;
  const cols = D.c2 - D.c1 + 1;
  // A single location cell takes the whole (one-dimensional) data range.
  if (cells.length === 1 && (rows === 1 || cols === 1)) {
    const at = cells[0];
    return at ? [{ formula: prefix + rangeAddress(D), location: rangeAddress(rangeOf(at)) }] : 'sparklineBadLocation';
  }
  const byRows = L.c1 === L.c2 ? rows === cells.length : cols !== cells.length && rows === cells.length;
  if (!byRows && cols !== cells.length) return 'sparklineSizeMismatch';
  return cells.map((at, i) => {
    const slice: Range = byRows ? { r1: D.r1 + i, r2: D.r1 + i, c1: D.c1, c2: D.c2 } : { r1: D.r1, r2: D.r2, c1: D.c1 + i, c2: D.c1 + i };
    return { formula: prefix + rangeAddress(slice), location: rangeAddress(rangeOf(at)) };
  });
}

const locationIn = (s: Sparkline, range: Range): boolean => {
  const at = parseRangeAddress(s.location);
  return at !== undefined && inRange(range, at.range.r1, at.range.c1);
};

/** Drop the sparklines drawn inside `range`, and any group left empty. */
function removeIn(ws: Worksheet, range: Range): void {
  const groups = ws.sparklineGroups ?? [];
  for (const g of groups) g.sparklines = g.sparklines.filter((s) => !locationIn(s, range));
  ws.sparklineGroups = groups.filter((g) => g.sparklines.length > 0);
}

/** Create Sparklines: one new group, replacing sparklines already in those cells. */
export function createSparklines(doc: SpreadsheetEditor, type: SparklineType, dataText: string, locationText: string): SparklineError | undefined {
  const ws = doc.ws;
  const sparklines = pairSparklines(dataText, locationText, ws.title);
  if (typeof sparklines === 'string') return sparklines;
  const loc = parseRangeAddress(locationText);
  doc.transact('Insert Sparklines', (tx) => {
    tx.sheet(ws, 'sparklineGroups');
    if (loc) removeIn(ws, loc.range);
    (ws.sparklineGroups ??= []).push(makeSparklineGroup({ ...(type === 'line' ? {} : { type }), sparklines }));
  });
  return undefined;
}

export function editSparklineGroup(doc: SpreadsheetEditor, groupIndex: number, label: string, fn: (g: SparklineGroup) => void): void {
  const ws = doc.ws;
  if (!ws.sparklineGroups?.[groupIndex]) return;
  doc.transact(label, (tx) => {
    tx.sheet(ws, 'sparklineGroups');
    const g = ws.sparklineGroups?.[groupIndex];
    if (g) fn(g);
  });
}

export function setSparklineType(doc: SpreadsheetEditor, groupIndex: number, type: SparklineType): void {
  editSparklineGroup(doc, groupIndex, 'Sparkline Type', (g) => {
    if (type === 'line') delete g.type;
    else g.type = type;
  });
}

/** The Show checkboxes of the Sparkline tab. */
export const SPARKLINE_FLAGS = ['high', 'low', 'negative', 'first', 'last', 'markers'] as const;
export type SparklineFlag = (typeof SPARKLINE_FLAGS)[number];

export function setSparklineFlag(doc: SpreadsheetEditor, groupIndex: number, flag: SparklineFlag, on: boolean): void {
  editSparklineGroup(doc, groupIndex, 'Sparkline Show', (g) => {
    g[flag] = on;
  });
}

/** Sparkline Color (the series) and Marker Color ▸ Negative / High / … slots. */
export type SparklineColorSlot = 'colorSeries' | 'colorNegative' | 'colorMarkers' | 'colorHigh' | 'colorLow' | 'colorFirst' | 'colorLast';

export function setSparklineColor(doc: SpreadsheetEditor, groupIndex: number, slot: SparklineColorSlot, hex: string): void {
  const color: Color = { rgb: `FF${hex.replace('#', '').toUpperCase()}` };
  editSparklineGroup(doc, groupIndex, 'Sparkline Color', (g) => {
    g[slot] = color;
  });
}

/** Clear: remove the sparklines in the selected cells. */
export function clearSparklines(doc: SpreadsheetEditor, ranges: readonly Range[]): void {
  const ws = doc.ws;
  if (!(ws.sparklineGroups ?? []).some((g) => g.sparklines.some((s) => ranges.some((r) => locationIn(s, r))))) return;
  doc.transact('Clear Sparklines', (tx) => {
    tx.sheet(ws, 'sparklineGroups');
    for (const r of ranges) removeIn(ws, r);
  });
}
