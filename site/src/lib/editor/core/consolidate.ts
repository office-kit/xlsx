// Data ▸ Consolidate: combine several source ranges into one summary table,
// by position or by matching top-row / left-column labels. Excel writes the
// results as values (unless linking), and keeps the dialog's settings on the
// sheet as <dataConsolidate>, which is how the dialog remembers its sources.

import type { CellValue } from '@office-kit/xlsx/cell';
import type { DataConsolidate, DataConsolidateFunction } from '@office-kit/xlsx/worksheet';
import type { CalcScalar } from '../calc/index.ts';
import { parseRangeAddress, quoteSheetName, rangeAddress, type Range } from './address.ts';
import { getCellAt, setValueAt } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import type { MessageKey } from '../i18n/i18n.svelte.ts';

/** The function list in Excel's order. */
export const CONSOLIDATE_FUNCTIONS: ReadonlyArray<{ readonly fn: DataConsolidateFunction; readonly label: MessageKey }> = [
  { fn: 'sum', label: 'dtFnSum' },
  { fn: 'count', label: 'dtFnCount' },
  { fn: 'average', label: 'dtFnAverage' },
  { fn: 'max', label: 'dtFnMax' },
  { fn: 'min', label: 'dtFnMin' },
  { fn: 'product', label: 'dtFnProduct' },
  { fn: 'countNums', label: 'dtFnCountNums' },
  { fn: 'stdDev', label: 'dtFnStdDev' },
  { fn: 'stdDevp', label: 'dtFnStdDevp' },
  { fn: 'var', label: 'dtFnVar' },
  { fn: 'varp', label: 'dtFnVarp' },
];

/** A source range's values, row-major. */
export type Grid = ReadonlyArray<ReadonlyArray<CalcScalar>>;

export function aggregate(fn: DataConsolidateFunction, values: readonly CalcScalar[]): number | undefined {
  const nums = values.filter((v): v is number => typeof v === 'number');
  const n = nums.length;
  const sum = nums.reduce((a, b) => a + b, 0);
  const sq = (mean: number) => nums.reduce((a, b) => a + (b - mean) ** 2, 0);
  switch (fn) {
    case 'sum':
      return n > 0 ? sum : undefined;
    case 'count':
      return values.filter((v) => v !== null && v !== '').length;
    case 'countNums':
      return n;
    case 'average':
      return n > 0 ? sum / n : undefined;
    case 'max':
      return n > 0 ? Math.max(...nums) : undefined;
    case 'min':
      return n > 0 ? Math.min(...nums) : undefined;
    case 'product':
      return n > 0 ? nums.reduce((a, b) => a * b, 1) : undefined;
    case 'stdDev':
      return n > 1 ? Math.sqrt(sq(sum / n) / (n - 1)) : undefined;
    case 'stdDevp':
      return n > 0 ? Math.sqrt(sq(sum / n) / n) : undefined;
    case 'var':
      return n > 1 ? sq(sum / n) / (n - 1) : undefined;
    case 'varp':
      return n > 0 ? sq(sum / n) / n : undefined;
  }
}

const labelOf = (v: CalcScalar): string => (v === null ? '' : String(v));

/**
 * Consolidate `sources`. With labels, rows / columns are matched by their
 * label (case-insensitive, first-seen order) and the labels head the result;
 * without, cells are matched by position. Empty result cells are null.
 */
export function consolidateGrids(sources: readonly Grid[], fn: DataConsolidateFunction, topLabels: boolean, leftLabels: boolean): CalcScalar[][] {
  const rowKeys: string[] = [];
  const colKeys: string[] = [];
  const rowIndex = new Map<string, number>();
  const colIndex = new Map<string, number>();
  const slot = (keys: string[], index: Map<string, number>, label: string) => {
    const k = label.toLowerCase();
    let i = index.get(k);
    if (i === undefined) {
      i = keys.length;
      keys.push(label);
      index.set(k, i);
    }
    return i;
  };
  const buckets = new Map<string, CalcScalar[]>();
  let rows = 0;
  let cols = 0;
  for (const grid of sources) {
    const header = topLabels ? (grid[0] ?? []) : [];
    for (let r = topLabels ? 1 : 0; r < grid.length; r++) {
      const line = grid[r] ?? [];
      const ri = leftLabels ? slot(rowKeys, rowIndex, labelOf(line[0] ?? null)) : r - (topLabels ? 1 : 0);
      for (let c = leftLabels ? 1 : 0; c < line.length; c++) {
        const ci = topLabels ? slot(colKeys, colIndex, labelOf(header[c] ?? null)) : c - (leftLabels ? 1 : 0);
        const key = `${ri}:${ci}`;
        let list = buckets.get(key);
        if (!list) {
          list = [];
          buckets.set(key, list);
        }
        list.push(line[c] ?? null);
        rows = Math.max(rows, ri + 1);
        cols = Math.max(cols, ci + 1);
      }
    }
  }
  const out: CalcScalar[][] = [];
  if (topLabels) out.push([...(leftLabels ? [null] : []), ...colKeys]);
  for (let r = 0; r < rows; r++) {
    const line: CalcScalar[] = leftLabels ? [rowKeys[r] ?? null] : [];
    for (let c = 0; c < cols; c++) line.push(aggregate(fn, buckets.get(`${r}:${c}`) ?? []) ?? null);
    out.push(line);
  }
  return out;
}

export interface ConsolidateSource {
  readonly sheet: string;
  readonly range: Range;
}

/** Parse "Sheet1!$A$1:$C$5" (sheet optional: the active one). */
export function parseSource(ctl: EditorController, text: string): ConsolidateSource | undefined {
  const parsed = parseRangeAddress(text.trim().replace(/^=/, ''));
  if (!parsed) return undefined;
  const title = parsed.sheet ?? ctl.doc.ws.title;
  const sheet = ctl.doc.calc.sheetName(title);
  return sheet ? { sheet, range: parsed.range } : undefined;
}

export function sourceText(src: ConsolidateSource): string {
  return `${quoteSheetName(src.sheet)}!${rangeAddress(src.range, true)}`;
}

function gridOf(ctl: EditorController, src: ConsolidateSource): Grid {
  const out: CalcScalar[][] = [];
  for (let r = src.range.r1; r <= src.range.r2; r++) {
    const line: CalcScalar[] = [];
    for (let c = src.range.c1; c <= src.range.c2; c++) line.push(ctl.doc.calc.cellValue(src.sheet, r, c));
    out.push(line);
  }
  return out;
}

function toCellValue(v: CalcScalar): CellValue {
  if (v === null || typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  return { kind: 'error', code: v.code };
}

/** Run Consolidate into the active cell's position, as one undo step. */
export function runConsolidate(ctl: EditorController, sources: readonly ConsolidateSource[], settings: Required<Pick<DataConsolidate, 'function' | 'topLabels' | 'leftLabels'>>): Range {
  const doc = ctl.doc;
  const ws = doc.ws;
  const result = consolidateGrids(
    sources.map((s) => gridOf(ctl, s)),
    settings.function,
    settings.topLabels,
    settings.leftLabels,
  );
  const { row, col } = doc.selection.active;
  const width = result.reduce((m, line) => Math.max(m, line.length), 0);
  const target: Range = { r1: row, c1: col, r2: row + Math.max(result.length, 1) - 1, c2: col + Math.max(width, 1) - 1 };
  doc.transact('Consolidate', (tx) => {
    tx.cells(ws, target);
    tx.sheet(ws, 'dataConsolidate');
    result.forEach((line, i) =>
      line.forEach((v, j) => {
        if (v !== null) setValueAt(ws, row + i, col + j, toCellValue(v), ctl.defaultStyleAt(row + i, col + j));
        else if (getCellAt(ws, row + i, col + j)) setValueAt(ws, row + i, col + j, null);
      }),
    );
    ws.dataConsolidate = {
      function: settings.function,
      topLabels: settings.topLabels,
      leftLabels: settings.leftLabels,
      dataRefs: sources.map((s) => ({ ref: rangeAddress(s.range), sheet: s.sheet })),
    };
  });
  return target;
}
