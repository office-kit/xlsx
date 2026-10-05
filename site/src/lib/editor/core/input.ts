// Turns the text a user typed into a cell into a cell value, the way Excel's
// data entry does: numbers with grouping, percentages, currency, fractions,
// dates and times (stored as serials plus an implied number format),
// booleans, error literals, formulas, and `'` to force text.

import type { CellValue, ExcelErrorCode } from '@office-kit/xlsx/cell';
import { makeFormula } from '@office-kit/xlsx/cell';
import { ERROR_CODES } from '@office-kit/xlsx/utils';
import { hmsFromSerial, ymdFromSerial } from '../calc/dates.ts';
import { fromStorageFormula, toStorageFormula } from '../calc/index.ts';

export type DateOrder = 'mdy' | 'ymd' | 'dmy';

export interface ParsedInput {
  readonly value: CellValue;
  /** Format Excel applies when the cell is still General (e.g. typing `12%`). */
  readonly impliedFormat?: string;
}

const MS_PER_DAY = 86_400_000;
const EPOCH_1900_OFFSET = 25_569; // serial of 1970-01-01 in the 1900 system
const EPOCH_1904_SHIFT = 1_462;

export function dateToSerial(year: number, month: number, day: number, date1904 = false): number {
  const serial = Date.UTC(year, month - 1, day) / MS_PER_DAY + EPOCH_1900_OFFSET;
  if (date1904) return serial - EPOCH_1904_SHIFT;
  // The 1900 system counts a 29 February 1900 (Lotus's leap-year bug, serial
  // 60), so 1 January 1900 is 1 and only dates from 1 March on line up with
  // the calendar.
  if (year === 1900 && month <= 2) return month === 2 && day === 29 ? LEAP_BUG_SERIAL : serial - 1;
  return serial;
}

const LEAP_BUG_SERIAL = 60;

export function serialToDate(serial: number, date1904 = false): Date {
  const s = date1904 ? serial + EPOCH_1904_SHIFT : serial;
  return new Date(Math.round((s - EPOCH_1900_OFFSET) * MS_PER_DAY));
}

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

function monthFromName(name: string): number | undefined {
  const i = MONTHS.indexOf(name.slice(0, 3).toLowerCase());
  return i < 0 ? undefined : i + 1;
}

function validDate(y: number, m: number, d: number): boolean {
  if (m < 1 || m > 12 || d < 1 || y < 1900 || y > 9999) return false;
  if (y === 1900 && m === 2 && d === 29) return true; // see dateToSerial
  const dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
}

/** Two-digit years follow Excel's 1930 pivot. */
function fullYear(y: number, text: string): number {
  if (text.length > 2) return y;
  return y < 30 ? 2000 + y : 1900 + y;
}

function parseDate(text: string, order: DateOrder, today: Date): { y: number; m: number; d: number; fmt: string } | undefined {
  const jp = /^(\d{1,4})年(\d{1,2})月(\d{1,2})日$/.exec(text);
  if (jp?.[1] && jp[2] && jp[3]) {
    const y = fullYear(Number(jp[1]), jp[1]);
    const m = Number(jp[2]);
    const d = Number(jp[3]);
    return validDate(y, m, d) ? { y, m, d, fmt: 'yyyy"年"m"月"d"日"' } : undefined;
  }
  const parts = text.split(/[/\-.]/);
  if (parts.length === 3 && parts.every((p) => /^\d+$/.test(p))) {
    const [a = '', b = '', c = ''] = parts;
    let y: number;
    let m: number;
    let d: number;
    if (a.length === 4 || order === 'ymd') {
      y = fullYear(Number(a), a);
      m = Number(b);
      d = Number(c);
    } else if (order === 'dmy') {
      d = Number(a);
      m = Number(b);
      y = fullYear(Number(c), c);
    } else {
      m = Number(a);
      d = Number(b);
      y = fullYear(Number(c), c);
    }
    // A typed date takes the locale's short date (numFmt 14, m/d/yy in Excel for Mac en-US), the
    // same format =DATE() gets; an ISO date (2024-01-31) is no exception.
    return validDate(y, m, d) ? { y, m, d, fmt: order === 'ymd' ? 'yyyy/m/d' : 'm/d/yy' } : undefined;
  }
  if (parts.length === 2 && parts.every((p) => /^\d{1,2}$/.test(p)) && /[/-]/.test(text)) {
    const [a = '', b = ''] = parts;
    const y = today.getFullYear();
    const [m, d] = order === 'dmy' ? [Number(b), Number(a)] : [Number(a), Number(b)];
    return validDate(y, m, d) ? { y, m, d, fmt: order === 'ymd' ? 'm"月"d"日"' : 'd-mmm' } : undefined;
  }
  // "1/2024", "Jan 2024", "January-24": the first of that month, shown as mmm-yy.
  const my = /^(?:(\d{1,2})[/-](\d{4})|([A-Za-z]{3,9})[\s-](\d{4}))$/.exec(text);
  if (my) {
    const m = my[1] ? Number(my[1]) : monthFromName(my[3] ?? '');
    const y = Number(my[2] ?? my[4]);
    return m && validDate(y, m, 1) ? { y, m, d: 1, fmt: 'mmm-yy' } : undefined;
  }
  // "5-Jan", "5 Jan 2024", "Jan 5", "Jan 5, 2024"
  const dm = /^(\d{1,2})[\s-]([A-Za-z]{3,9})(?:[\s-](\d{2,4}))?$/.exec(text);
  if (dm?.[1] && dm[2]) {
    const m = monthFromName(dm[2]);
    const y = dm[3] ? fullYear(Number(dm[3]), dm[3]) : today.getFullYear();
    const d = Number(dm[1]);
    return m && validDate(y, m, d) ? { y, m, d, fmt: dm[3] ? 'd-mmm-yy' : 'd-mmm' } : undefined;
  }
  const md = /^([A-Za-z]{3,9})\s+(\d{1,2})(?:,?\s+(\d{2,4}))?$/.exec(text);
  if (md?.[1] && md[2]) {
    const m = monthFromName(md[1]);
    const y = md[3] ? fullYear(Number(md[3]), md[3]) : today.getFullYear();
    const d = Number(md[2]);
    return m && validDate(y, m, d) ? { y, m, d, fmt: md[3] ? 'd-mmm-yy' : 'd-mmm' } : undefined;
  }
  return undefined;
}

function parseTime(text: string): { fraction: number; fmt: string } | undefined {
  const m = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([AaPp][Mm])?$/.exec(text);
  if (!m?.[1] || !m[2]) return undefined;
  let h = Number(m[1]);
  const min = Number(m[2]);
  const sec = m[3] ? Number(m[3]) : 0;
  const ampm = m[4]?.toLowerCase();
  if (min > 59 || sec >= 60) return undefined;
  if (ampm) {
    if (h < 1 || h > 12) return undefined;
    if (ampm === 'pm' && h !== 12) h += 12;
    if (ampm === 'am' && h === 12) h = 0;
  } else if (h > 23) {
    // Excel accepts "25:30" as an elapsed duration.
    return { fraction: (h * 3600 + min * 60 + sec) / 86400, fmt: m[3] ? '[h]:mm:ss' : '[h]:mm' };
  }
  const fmt = ampm ? (m[3] ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : m[3] ? 'h:mm:ss' : 'h:mm';
  return { fraction: (h * 3600 + min * 60 + sec) / 86400, fmt };
}

const CURRENCY_SYMBOLS = ['$', '¥', '￥', '€', '£'];

function parseNumber(raw: string): ParsedInput | undefined {
  let text = raw.trim();
  let negative = false;
  if (/^\(.*\)$/.test(text)) {
    negative = true;
    text = text.slice(1, -1).trim();
  }
  let currency: string | undefined;
  for (const sym of CURRENCY_SYMBOLS) {
    if (text.startsWith(sym) || text.startsWith(`-${sym}`)) {
      currency = sym;
      text = text.replace(sym, '');
      break;
    }
  }
  let percent = false;
  if (text.endsWith('%')) {
    percent = true;
    text = text.slice(0, -1).trim();
  }
  const grouped = text.includes(',');
  if (grouped && !/^[+-]?\d{1,3}(,\d{3})*(\.\d*)?$/.test(text)) return undefined;
  const plain = text.replaceAll(',', '');
  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(plain)) return undefined;
  let n = Number(truncateDigits(plain));
  if (!Number.isFinite(n)) return undefined;
  if (negative) n = -n;
  if (percent) {
    const decimals = plain.split('.')[1]?.length ?? 0;
    return { value: n / 100, impliedFormat: decimals > 0 ? `0.${'0'.repeat(decimals)}%` : '0%' };
  }
  if (currency) {
    const decimals = plain.includes('.') ? '.00' : '';
    const sym = currency === '￥' ? '¥' : currency;
    return { value: n, impliedFormat: `"${sym}"#,##0${decimals};-"${sym}"#,##0${decimals}` };
  }
  if (/[eE]/.test(plain)) return { value: n, impliedFormat: '0.00E+00' };
  if (grouped) return { value: n, impliedFormat: plain.includes('.') ? '#,##0.00' : '#,##0' };
  return { value: n };
}

const SIGNIFICANT_DIGITS = 15;

/** Excel keeps 15 significant digits of typed input and zeroes the rest (no rounding). */
function truncateDigits(plain: string): string {
  const m = /^([+-]?)(\d*)(?:\.(\d*))?([eE][+-]?\d+)?$/.exec(plain);
  if (!m) return plain;
  const [, sign = '', int = '', frac = '', exp = ''] = m;
  const digits = int + frac;
  const first = digits.search(/[1-9]/);
  if (first < 0 || digits.length - first <= SIGNIFICANT_DIGITS) return plain;
  const kept = digits.slice(0, first + SIGNIFICANT_DIGITS).padEnd(digits.length, '0');
  return `${sign}${kept.slice(0, int.length)}.${kept.slice(int.length)}${exp}`;
}

function parseFraction(text: string): ParsedInput | undefined {
  const m = /^(-)?(?:(\d+)\s+)?(\d+)\/(\d+)$/.exec(text);
  if (!m?.[3] || !m[4] || m[2] === undefined) return undefined;
  const den = Number(m[4]);
  if (den === 0) return undefined;
  const v = Number(m[2]) + Number(m[3]) / den;
  return { value: m[1] ? -v : v, impliedFormat: den < 10 ? '# ?/?' : '# ??/??' };
}

export interface ParseOptions {
  readonly dateOrder?: DateOrder;
  readonly date1904?: boolean;
  readonly today?: Date;
  /** The workbook's sheet names, so a typed reference takes the real name's case. */
  readonly sheetTitles?: readonly string[];
  /** The table holding the cell, which unqualified structured references point at. */
  readonly table?: string | undefined;
}

// A General cell takes the format of what its formula's leading function
// returns: =TODAY() reads as a date, =NOW() as a date and time (Excel 16).
const LEADING_FUNCTION_FORMATS: ReadonlyMap<string, string> = new Map([
  ['DATE', 'm/d/yy'],
  ['TODAY', 'm/d/yy'],
  ['DATEVALUE', 'm/d/yy'],
  ['NOW', 'm/d/yy h:mm'],
  ['TIME', 'h:mm AM/PM'],
  ['TIMEVALUE', 'h:mm AM/PM'],
]);

function formulaInput(body: string, opts: ParseOptions): ParsedInput {
  const fn = /^[+-]*([A-Za-z.]+)\(/.exec(body)?.[1]?.toUpperCase();
  const impliedFormat = fn === undefined ? undefined : LEADING_FUNCTION_FORMATS.get(fn);
  const value = makeFormula(toStorageFormula(body, opts.sheetTitles, opts.table));
  return impliedFormat === undefined ? { value } : { value, impliedFormat };
}

/** U+FF01–U+FF5E are the full-width forms of ASCII `!`–`~`. */
const FULLWIDTH_ASCII_OFFSET = 0xfee0;
const FULLWIDTH_RE = /[\uFF01-\uFF5E\u3000]/;

/** `＝ＳＵＭ（Ａ１）` → `=SUM(A1)`; text inside `"…"` keeps its spelling. */
function halfWidth(text: string): string {
  let out = '';
  let quoted = false;
  for (const ch of text) {
    const code = ch.charCodeAt(0);
    const half = code >= 0xff01 && code <= 0xff5e ? String.fromCharCode(code - FULLWIDTH_ASCII_OFFSET) : code === 0x3000 ? ' ' : ch;
    if (half === '"') quoted = !quoted;
    out += quoted && half !== '"' ? ch : half;
  }
  return out;
}

/**
 * Interpret typed text. Formula text keeps the user's spelling (after the `=`)
 * apart from the `_xlfn.` / `_xlpm.` prefixes Excel needs in a file; the
 * formula engine validates it separately so a syntax error can be reported
 * before the value is committed.
 */
export function parseInput(input: string, opts: ParseOptions = {}): ParsedInput {
  const parsed = parseTyped(input, opts);
  if (typeof parsed.value !== 'string' || !FULLWIDTH_RE.test(input)) return parsed;
  // Excel reads full-width digits, signs and formulas (an IME's default) as
  // their ASCII forms; text that is still text keeps its full-width spelling.
  const folded = parseTyped(halfWidth(input), opts);
  return typeof folded.value === 'string' ? parsed : folded;
}

function parseTyped(input: string, opts: ParseOptions): ParsedInput {
  if (input === '') return { value: null };
  if (input.startsWith("'")) return { value: input.slice(1) };
  if (input.startsWith('=') && input.length > 1) return formulaInput(input.slice(1), opts);
  // Excel turns "+A1", "-A1*2" and "+1+2" into formulas, but keeps "+5" / "-5"
  // numbers and "- item" text.
  if (
    (input.startsWith('+') || input.startsWith('-')) &&
    input.length > 1 &&
    parseNumber(input) === undefined &&
    parseFraction(input.trim()) === undefined &&
    /^[+-][A-Za-z0-9($.+-]/.test(input)
  ) {
    return formulaInput(input, opts);
  }
  const text = input.trim();
  const upper = text.toUpperCase();
  if (upper === 'TRUE') return { value: true };
  if (upper === 'FALSE') return { value: false };
  if (ERROR_CODES.has(upper)) return { value: { kind: 'error', code: upper as ExcelErrorCode } };

  const num = parseNumber(text);
  if (num) return num;
  const frac = parseFraction(text);
  if (frac) return frac;

  const today = opts.today ?? new Date();
  const order = opts.dateOrder ?? 'mdy';
  const time = parseTime(text);
  if (time) return { value: time.fraction, impliedFormat: time.fmt };
  const date = parseDate(text, order, today);
  if (date) return { value: dateToSerial(date.y, date.m, date.d, opts.date1904), impliedFormat: date.fmt };
  // "2024/1/5 13:30", "1/31/2024 1:30 PM": any space may separate the date from the time.
  for (let space = text.indexOf(' '); space > 0; space = text.indexOf(' ', space + 1)) {
    const d = parseDate(text.slice(0, space).trim(), order, today);
    const t = d && parseTime(text.slice(space + 1).trim());
    if (d && t) {
      return {
        value: dateToSerial(d.y, d.m, d.d, opts.date1904) + t.fraction,
        impliedFormat: `${order === 'ymd' ? 'yyyy/m/d' : 'm/d/yy'} h:mm`,
      };
    }
  }
  return { value: input };
}

/** Largest serial Excel shows as a date (12/31/9999); anything past it, or below 0, shows as a number. */
const MAX_DATE_SERIAL = 2958465;
const SECONDS_PER_DAY = 86_400;

/**
 * The text the cell editor starts with when the user presses F2 or
 * double-clicks: the formula with its `=`, or the value in an editable form.
 * Like Excel's formula bar, it never drops precision the cell's format hides:
 * a date shows its four-digit year and, when it has one, its time to the
 * second, and a percentage shows every digit — so committing it unchanged
 * keeps the value.
 */
export function editTextFor(value: CellValue, formatted: string, isDateFormat: boolean, opts: { dateOrder: DateOrder; date1904: boolean; table?: string | undefined }): string {
  if (value === null) return '';
  if (typeof value === 'object' && !(value instanceof Date)) {
    if (value.kind === 'formula') return `=${fromStorageFormula(value.formula, opts.table)}`;
    if (value.kind === 'error') return value.code;
    if (value.kind === 'rich-text') return value.runs.map((r) => r.text).join('');
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') {
    if (isDateFormat && value >= 0 && value < MAX_DATE_SERIAL + 1) return dateEditText(value, opts);
    if (formatted.endsWith('%')) return `${Number((value * 100).toPrecision(15))}%`;
    return String(Number(value.toPrecision(15)));
  }
  return String(value);
}

function dateEditText(serial: number, { dateOrder, date1904 }: { dateOrder: DateOrder; date1904: boolean }): string {
  // Round to the second first, so 23:59:59.7 carries into the next day.
  const rounded = Math.round(serial * SECONDS_PER_DAY) / SECONDS_PER_DAY;
  const day = Math.floor(rounded);
  let time = '';
  if (rounded !== day) {
    const { h, mi, s } = hmsFromSerial(rounded);
    const mmss = `${String(mi).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
    time = dateOrder === 'mdy' ? `${h % 12 || 12}:${mmss} ${h < 12 ? 'AM' : 'PM'}` : `${h}:${mmss}`;
  }
  // A time with no date part shows on its own.
  if (day === 0 && time) return time;
  const { y, m, d } = ymdFromSerial(day, date1904);
  const date = dateOrder === 'ymd' ? `${y}/${m}/${d}` : dateOrder === 'dmy' ? `${d}/${m}/${y}` : `${m}/${d}/${y}`;
  return time ? `${date} ${time}` : date;
}
