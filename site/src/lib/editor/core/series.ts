// Home ▸ Fill ▸ Series: fill each row (or column) of the selection from its
// first cell with a linear, growth or date series, or with AutoFill's own
// pattern, stopping at an optional stop value like Excel's Series dialog.

import type { CellValue } from '@office-kit/xlsx/cell';
import { makeCell } from '@office-kit/xlsx/cell';
import { isDateFormat } from '@office-kit/xlsx/styles';
import { translateFormula } from '../calc/index.ts';
import type { Range } from './address.ts';
import { MAX_COL, MAX_ROW } from './address.ts';
import { getCellAt, isBlank } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import { extendSeries, type SeriesSeed } from './fill.ts';
import { dateToSerial, serialToDate } from './input.ts';

export type SeriesType = 'linear' | 'growth' | 'date' | 'autofill';
export type DateUnit = 'day' | 'weekday' | 'month' | 'year';

export interface SeriesOptions {
  readonly type: SeriesType;
  readonly step: number;
  /** Values beyond this (in the step's direction) are not written. */
  readonly stop?: number;
  readonly dateUnit: DateUnit;
}

function roundFloat(v: number): number {
  return Number(v.toPrecision(15));
}

function addMonths(serial: number, months: number, date1904: boolean): number {
  const d = serialToDate(serial, date1904);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + months;
  const targetYear = y + Math.floor(m / 12);
  const targetMonth = ((m % 12) + 12) % 12;
  // Excel clamps to the month's last day (Jan 31 + 1 month = Feb 28/29).
  const lastDay = new Date(Date.UTC(targetYear, targetMonth + 1, 0)).getUTCDate();
  const day = Math.min(d.getUTCDate(), lastDay);
  return dateToSerial(targetYear, targetMonth + 1, day, date1904) + (serial - Math.floor(serial));
}

function isWeekend(serial: number, date1904: boolean): boolean {
  const day = serialToDate(Math.floor(serial), date1904).getUTCDay();
  return day === 0 || day === 6;
}

function addWeekdays(serial: number, days: number, date1904: boolean): number {
  let s = serial;
  const dir = days >= 0 ? 1 : -1;
  for (let n = Math.abs(days); n > 0; ) {
    s += dir;
    if (!isWeekend(s, date1904)) n--;
  }
  return s;
}

/**
 * The values following `start` in a numeric series, at most `count` of them,
 * cut off at `stop`. AutoFill is not numeric and is handled by the caller.
 */
export function seriesValues(start: number, count: number, opts: SeriesOptions, date1904 = false): number[] {
  const out: number[] = [];
  // The stop bounds the series in whichever direction it runs (a growth step
  // below 1 shrinks toward the stop rather than growing past it).
  const beyond = (v: number): boolean => opts.stop !== undefined && (v >= start ? v > opts.stop : v < opts.stop);
  for (let i = 1; i <= count; i++) {
    let v: number;
    if (opts.type === 'growth') v = roundFloat(start * opts.step ** i);
    else if (opts.type === 'date') {
      switch (opts.dateUnit) {
        case 'day':
          v = start + opts.step * i;
          break;
        case 'weekday':
          v = addWeekdays(start, Math.round(opts.step) * i, date1904);
          break;
        case 'month':
          v = addMonths(start, Math.round(opts.step) * i, date1904);
          break;
        case 'year':
          v = addMonths(start, Math.round(opts.step) * 12 * i, date1904);
          break;
      }
    } else v = roundFloat(start + opts.step * i);
    if (beyond(v)) break;
    out.push(v);
  }
  return out;
}

/** Most cells a single-cell selection fills toward a stop value. */
const OPEN_ENDED_LIMIT = 100_000;

/**
 * Fill `range` line by line: `byRows` makes each row a series running right,
 * otherwise each column runs down. A one-cell-long line with a stop value
 * extends until the stop is reached.
 */
export function fillSeries(ctl: EditorController, range: Range, byRows: boolean, opts: SeriesOptions): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const date1904 = doc.wb.date1904;
  const lineLength = byRows ? range.c2 - range.c1 + 1 : range.r2 - range.r1 + 1;
  const limit = byRows ? MAX_COL - range.c1 : MAX_ROW - range.r1;
  const openEnded = lineLength === 1 && opts.stop !== undefined && opts.type !== 'autofill';
  const count = openEnded ? Math.min(limit, OPEN_ENDED_LIMIT) : lineLength - 1;
  const lines = byRows ? range.r2 - range.r1 + 1 : range.c2 - range.c1 + 1;
  const isDate = (styleId: number) => isDateFormat(doc.styles.get(styleId).numFmt);

  const plan: Array<{ row: number; col: number; value: CellValue; styleId: number }> = [];
  for (let k = 0; k < lines; k++) {
    const row0 = byRows ? range.r1 + k : range.r1;
    const col0 = byRows ? range.c1 : range.c1 + k;
    const at = (i: number) => (byRows ? { row: row0, col: col0 + i } : { row: row0 + i, col: col0 });
    const first = getCellAt(ws, row0, col0);
    if (!first || isBlank(first)) continue;
    let produced: SeriesSeed[];
    let offset: number;
    if (opts.type === 'autofill') {
      // The leading filled cells are the pattern, as when dragging the fill handle.
      const seeds: SeriesSeed[] = [];
      for (let i = 0; i <= count; i++) {
        const c = getCellAt(ws, at(i).row, at(i).col);
        if (!c || isBlank(c)) break;
        seeds.push({ value: c.value, styleId: c.styleId });
      }
      offset = seeds.length;
      produced = extendSeries(seeds, Math.max(0, count + 1 - seeds.length), 1, byRows ? 'col' : 'row', { translate: translateFormula, isDate, date1904: doc.wb.date1904 });
    } else {
      if (typeof first.value !== 'number') continue;
      offset = 1;
      produced = seriesValues(first.value, count, opts, date1904).map((value) => ({ value, styleId: first.styleId }));
    }
    for (const [i, seed] of produced.entries()) plan.push({ ...at(offset + i), value: seed.value, styleId: seed.styleId });
  }
  if (plan.length === 0) return;
  // An open-ended series can run past the selection.
  let far = byRows ? range.c2 : range.r2;
  for (const p of plan) far = Math.max(far, byRows ? p.col : p.row);
  const touched: Range = byRows ? { ...range, c2: far } : { ...range, r2: far };
  doc.transact('Series', (tx) => {
    tx.cells(ws, touched);
    for (const { row, col, value, styleId } of plan) {
      let rowMap = ws.rows.get(row);
      if (!rowMap) {
        rowMap = new Map();
        ws.rows.set(row, rowMap);
      }
      rowMap.set(col, makeCell(row, col, value, styleId));
    }
  });
  ctl.selectRange(touched, doc.selection.active);
}
