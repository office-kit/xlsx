// Excel's implicit conversions between numbers, text and booleans, and its
// ordering of mixed-type values.

import { parseDateTimeText } from './dates.ts';
import { type CalcError, type CalcScalar, ERRORS, isError } from './types.ts';

/** Excel keeps 15 significant digits; comparisons and text conversion happen at that precision. */
const PRECISION = 15;

const MAX_FULL_TEXT = 20;

export function round15(n: number): number {
  if (n === 0 || !Number.isFinite(n)) return n;
  return Number(n.toPrecision(PRECISION));
}

/**
 * A number as Excel turns it into text (`=1/3&""`, TEXTJOIN, string
 * comparison): at most 15 significant digits, no grouping, written out in
 * full unless that takes more than 20 characters. Measured against Excel:
 * 1E19 and 1E-15 print in full, 1E20, 1E-19 and 2/3*1E-5 as `6.66666666666667E-06`.
 */
export function numberToText(n: number): string {
  if (n === 0) return '0';
  const sign = n < 0 ? '-' : '';
  const [mantissa = '', e = '0'] = Math.abs(n).toExponential(PRECISION - 1).split('e');
  const digits = mantissa.replace('.', '').replace(/0+$/, '');
  const exp = Number(e);
  let full: string;
  if (exp < 0) full = `0.${'0'.repeat(-exp - 1)}${digits}`;
  else if (digits.length <= exp + 1) full = digits + '0'.repeat(exp + 1 - digits.length);
  else full = `${digits.slice(0, exp + 1)}.${digits.slice(exp + 1)}`;
  if (full.length <= MAX_FULL_TEXT) return sign + full;
  const point = digits.length > 1 ? `${digits[0]}.${digits.slice(1)}` : digits;
  return `${sign}${point}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
}

const NUMERIC_TEXT = /^([+-]?)\s*(\$?)\s*([+-]?)((?:\d{1,3}(?:,\d{3})+|\d*)(?:\.\d*)?)(?:[eE]([+-]?\d+))?\s*(%?)$/;

/**
 * Read text the way Excel coerces it to a number: `" 1,234.5 "`, `"$12"`,
 * `"50%"`, `"(12)"`, `"1e3"`, and dates / times such as `"2024-01-15"` or
 * `"12:30"`. Undefined when the text is not numeric.
 */
const MIXED_FRACTION = /^(-)?(\d+)\s+(\d+)\/(\d+)$/;

export function parseNumericText(input: string, date1904 = false): number | undefined {
  let text = input.trim();
  if (text === '') return undefined;
  let negate = false;
  if (text.startsWith('(') && text.endsWith(')')) {
    text = text.slice(1, -1).trim();
    negate = true;
  }
  const m = NUMERIC_TEXT.exec(text);
  if (m !== null && /\d/.test(m[4] ?? '') && !(m[1] !== '' && m[3] !== '')) {
    let value = Number((m[4] ?? '').replace(/,/g, ''));
    if (m[5] !== undefined) value *= 10 ** Number(m[5]);
    if (m[6] === '%') value /= 100;
    if (m[1] === '-' || m[3] === '-') value = -value;
    if (negate) value = -value;
    return Number.isFinite(value) ? value : undefined;
  }
  if (negate) return undefined;
  // "1 1/2" is a mixed number; a bare "1/2" stays a date (January 2), as in Excel.
  const frac = MIXED_FRACTION.exec(text);
  if (frac && Number(frac[4]) !== 0) {
    const v = Number(frac[2]) + Number(frac[3]) / Number(frac[4]);
    return frac[1] ? -v : v;
  }
  const dt = parseDateTimeText(text, date1904, new Date().getFullYear());
  return dt === undefined ? undefined : dt.date + dt.time;
}

export function toNumber(v: CalcScalar, date1904 = false): number | CalcError {
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') return parseNumericText(v, date1904) ?? ERRORS.VALUE;
  return v;
}

export function toText(v: CalcScalar): string | CalcError {
  if (typeof v === 'string') return v;
  if (v === null) return '';
  if (typeof v === 'number') return numberToText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return v;
}

export function toBoolean(v: CalcScalar): boolean | CalcError {
  if (typeof v === 'boolean') return v;
  if (v === null) return false;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') {
    const upper = v.toUpperCase();
    if (upper === 'TRUE') return true;
    if (upper === 'FALSE') return false;
    return ERRORS.VALUE;
  }
  return v;
}

// Excel orders numbers < text < booleans; within text, case is ignored.
const typeRank = (v: number | string | boolean): number => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);

/**
 * Three-way comparison of two non-error scalars with Excel's operator rules:
 * a blank takes the type of the other side, mixed types order by type, text
 * compares case-insensitively, numbers at 15 significant digits.
 */
export function compareScalars(a: Exclude<CalcScalar, CalcError>, b: Exclude<CalcScalar, CalcError>): number {
  if (a === null && b === null) return 0;
  const left = a ?? blankLike(b);
  const right = b ?? blankLike(a);
  const ra = typeRank(left);
  const rb = typeRank(right);
  if (ra !== rb) return ra - rb;
  if (typeof left === 'number' && typeof right === 'number') {
    const x = round15(left);
    const y = round15(right);
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof left === 'string' && typeof right === 'string') return compareText(left, right);
  return Number(left) - Number(right);
}

const blankLike = (other: CalcScalar): number | string | boolean =>
  typeof other === 'string' ? '' : typeof other === 'boolean' ? false : 0;

export function compareText(a: string, b: string): number {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

/** Equality as the `=` operator sees it. */
export function scalarsEqual(a: Exclude<CalcScalar, CalcError>, b: Exclude<CalcScalar, CalcError>): boolean {
  return compareScalars(a, b) === 0;
}

export function isNumericScalar(v: CalcScalar): v is number {
  return typeof v === 'number';
}

/** First error among `values`, if any — Excel reports the leftmost one. */
export function firstError(values: readonly CalcScalar[]): CalcError | undefined {
  for (const v of values) if (isError(v)) return v;
  return undefined;
}

/** Within this relative distance, two magnitudes that cancel count as equal (2^-48, as LibreOffice's approxAdd). */
const CANCEL_EPSILON = 2 ** -48;

/** `a + b`, but 0 when `b` cancels `a` down to floating-point noise, as Excel's last addition does. */
export function approxAdd(a: number, b: number): number {
  if (Math.sign(a) === -Math.sign(b) && Math.abs(a + b) < Math.abs(a) * CANCEL_EPSILON) return 0;
  return a + b;
}
