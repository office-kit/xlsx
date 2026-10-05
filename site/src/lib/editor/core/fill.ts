// AutoFill: what dragging the fill handle produces. Excel looks at the source
// block one line at a time (a column when filling down, a row when filling
// across) and continues each line's pattern: linear number trends, dates,
// "Item 1, Item 2", weekday and month names, or plain repetition. Formulas are
// re-anchored by their offset from the source cell.

import type { CellValue } from '@office-kit/xlsx/cell';

export interface SeriesSeed {
  readonly value: CellValue;
  readonly styleId: number;
}

const WEEKDAYS_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const WEEKDAYS_LONG = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const MONTHS_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTHS_LONG = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const WEEKDAYS_JA = ['日', '月', '火', '水', '木', '金', '土'];
const WEEKDAYS_JA_LONG = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];
const MONTHS_JA = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];

/** Excel's built-in custom lists (Preferences ▸ Custom Lists). */
export const CUSTOM_LISTS: ReadonlyArray<readonly string[]> = [WEEKDAYS_SHORT, WEEKDAYS_LONG, MONTHS_SHORT, MONTHS_LONG, WEEKDAYS_JA, WEEKDAYS_JA_LONG, MONTHS_JA, QUARTERS];

function listMatch(text: string): { list: readonly string[]; index: number } | undefined {
  for (const list of CUSTOM_LISTS) {
    const index = list.findIndex((item) => item.toLowerCase() === text.toLowerCase());
    if (index >= 0) return { list, index };
  }
  return undefined;
}

function matchCase(template: string, word: string): string {
  if (template === template.toUpperCase()) return word.toUpperCase();
  if (template === template.toLowerCase()) return word.toLowerCase();
  return word;
}

/** Text ending in an integer ("Item 7", "Q3-2024" is not). */
const TEXT_NUMBER = /^(.*?)(\d+)$/;

function isFormula(v: CellValue): v is Extract<CellValue, { kind: 'formula' }> {
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula';
}

/**
 * How seeds extend: `auto` follows Excel's guess (a lone number copies, two
 * numbers or a list continue), `copy` repeats, `series` always counts on.
 */
export type SeriesMode = 'auto' | 'copy' | 'series';

export interface FillContext {
  readonly translate: (formula: string, dRow: number, dCol: number) => string;
  /** True when the seed's number format is a date (dates step by day, then by the seed's interval). */
  readonly isDate: (styleId: number) => boolean;
}

/**
 * Extend one line of seeds by `count` cells. `step` is +1 when filling down /
 * right and -1 when filling up / left (series then run backwards). `offsets`
 * give each produced cell's distance from the seed it copies, for formula
 * re-anchoring along the fill axis.
 */
export function extendSeries(
  seeds: readonly SeriesSeed[],
  count: number,
  direction: 1 | -1,
  axis: 'row' | 'col',
  ctx: FillContext,
  mode: SeriesMode = 'auto',
): SeriesSeed[] {
  const n = seeds.length;
  const out: SeriesSeed[] = [];
  if (n === 0) return out;
  const values = seeds.map((s) => s.value);

  const nums = values.filter((v): v is number => typeof v === 'number');
  if (mode !== 'copy' && nums.length === n) {
    // A single number copies (Excel), unless it's a date, which counts up by a day.
    if (n === 1 && mode === 'auto' && !ctx.isDate(seeds[0]?.styleId ?? 0)) return repeat(seeds, count, direction, axis, ctx);
    const { slope, intercept } = n === 1 ? { slope: 1, intercept: nums[0] ?? 0 } : linearFit(nums);
    for (let i = 1; i <= count; i++) {
      const x = direction === 1 ? n - 1 + i : -i;
      const seed = seeds[direction === 1 ? (i - 1) % n : n - 1 - ((i - 1) % n)] ?? seeds[0];
      out.push({ value: roundFloat(intercept + slope * x), styleId: seed?.styleId ?? 0 });
    }
    return out;
  }

  // Lists and "text + number" continue per seed position.
  const texts = values.filter((v): v is string => typeof v === 'string');
  if (mode !== 'copy' && texts.length === n) {
    const lists = texts.map(listMatch);
    if (lists.every((m) => m !== undefined) && lists.every((m) => m?.list === lists[0]?.list)) {
      const list = lists[0]?.list ?? [];
      const idx = lists.map((m) => m?.index ?? 0);
      const stride = n > 1 ? ((idx[n - 1] ?? 0) - (idx[0] ?? 0)) / (n - 1) : 1;
      const step = Number.isInteger(stride) && stride !== 0 ? stride : 1;
      for (let i = 1; i <= count; i++) {
        const base = direction === 1 ? (idx[n - 1] ?? 0) + step * i : (idx[0] ?? 0) - step * i;
        const word = list[((base % list.length) + list.length) % list.length] ?? '';
        const seed = seeds[direction === 1 ? (i - 1) % n : n - 1 - ((i - 1) % n)];
        out.push({ value: matchCase(texts[0] ?? '', word), styleId: seed?.styleId ?? 0 });
      }
      return out;
    }
    const parts = texts.map((t) => TEXT_NUMBER.exec(t));
    if (parts.every((p) => p !== null) && parts.every((p) => p?.[1] === parts[0]?.[1])) {
      const prefix = parts[0]?.[1] ?? '';
      const suffixes = parts.map((p) => Number(p?.[2] ?? 0));
      const width = parts[0]?.[2]?.length ?? 1;
      const pad = parts[0]?.[2]?.startsWith('0') ? width : 0;
      const step = n > 1 ? ((suffixes[n - 1] ?? 0) - (suffixes[0] ?? 0)) / (n - 1) : 1;
      for (let i = 1; i <= count; i++) {
        const v = Math.abs(Math.round(direction === 1 ? (suffixes[n - 1] ?? 0) + step * i : (suffixes[0] ?? 0) - step * i));
        const seed = seeds[direction === 1 ? (i - 1) % n : n - 1 - ((i - 1) % n)];
        out.push({ value: `${prefix}${String(v).padStart(pad, '0')}`, styleId: seed?.styleId ?? 0 });
      }
      return out;
    }
  }
  return repeat(seeds, count, direction, axis, ctx);
}

function repeat(seeds: readonly SeriesSeed[], count: number, direction: 1 | -1, axis: 'row' | 'col', ctx: FillContext): SeriesSeed[] {
  const n = seeds.length;
  const out: SeriesSeed[] = [];
  for (let i = 1; i <= count; i++) {
    const k = direction === 1 ? (i - 1) % n : n - 1 - ((i - 1) % n);
    const seed = seeds[k];
    if (!seed) continue;
    // Distance from the copied seed to the produced cell along the fill axis.
    const distance = direction === 1 ? n - k + i - 1 : -(k + i);
    let value = seed.value;
    if (isFormula(value) && value.formula) {
      const formula = ctx.translate(value.formula, axis === 'row' ? distance : 0, axis === 'col' ? distance : 0);
      value = { kind: 'formula', t: 'normal', formula };
    }
    out.push({ value, styleId: seed.styleId });
  }
  return out;
}

function linearFit(ys: readonly number[]): { slope: number; intercept: number } {
  const n = ys.length;
  let sx = 0;
  let sy = 0;
  let sxx = 0;
  let sxy = 0;
  ys.forEach((y, x) => {
    sx += x;
    sy += y;
    sxx += x * x;
    sxy += x * y;
  });
  const slope = (n * sxy - sx * sy) / (n * sxx - sx * sx);
  return { slope, intercept: (sy - slope * sx) / n };
}

function roundFloat(v: number): number {
  return Number(v.toPrecision(15));
}
