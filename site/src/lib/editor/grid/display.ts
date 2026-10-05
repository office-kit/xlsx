// What a cell shows: formatted text, its colour override from the number
// format (`[Red]`), and the alignment Excel uses for General alignment.

import type { Cell, CellValue } from '@office-kit/xlsx/cell';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import type { Workbook } from '@office-kit/xlsx/workbook';

export type ValueKind = 'number' | 'text' | 'bool' | 'error' | 'empty';

export interface CellDisplay {
  readonly text: string;
  readonly kind: ValueKind;
  /** Colour from a `[Red]` / `[Color10]` format section. */
  readonly color?: string;
  /** Raw number, for General-format width fitting. */
  readonly number?: number;
}

function effectiveValue(value: CellValue): CellValue | number | string | boolean {
  if (value !== null && typeof value === 'object' && !(value instanceof Date) && value.kind === 'formula') {
    if (value.cachedValueType === 'error') return { kind: 'error', code: String(value.cachedValue) as `#${string}` };
    return value.cachedValue ?? null;
  }
  return value;
}

export function valueKind(value: CellValue): ValueKind {
  const v = effectiveValue(value);
  if (v === null || v === '') return 'empty';
  if (typeof v === 'number' || v instanceof Date) return 'number';
  if (typeof v === 'boolean') return 'bool';
  if (typeof v === 'string') return 'text';
  if (v.kind === 'error') return 'error';
  if (v.kind === 'duration') return 'number';
  return 'text';
}

const NAMED_COLORS: Readonly<Record<string, string>> = {
  black: '#000000',
  white: '#FFFFFF',
  red: '#FF0000',
  green: '#00FF00',
  blue: '#0000FF',
  yellow: '#FFFF00',
  magenta: '#FF00FF',
  cyan: '#00FFFF',
};

// The 56-colour legacy palette `[ColorN]` indexes (1-based).
const PALETTE_56 = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000',
  '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99',
  '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i] ?? '';
    if (ch === '"') quoted = !quoted;
    if (ch === '\\' && !quoted) {
      cur += ch + (code[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === ';' && !quoted) {
      out.push(cur);
      cur = '';
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}

const colorCache = new Map<string, ReadonlyArray<string | undefined>>();

function sectionColors(code: string): ReadonlyArray<string | undefined> {
  let colors = colorCache.get(code);
  if (!colors) {
    colors = splitSections(code).map((section) => {
      const m = /\[(black|white|red|green|blue|yellow|magenta|cyan|color\s*(\d{1,2}))\]/i.exec(section);
      if (!m?.[1]) return undefined;
      if (m[2]) {
        const hex = PALETTE_56[Number(m[2]) - 1];
        return hex ? `#${hex}` : undefined;
      }
      return NAMED_COLORS[m[1].toLowerCase()];
    });
    colorCache.set(code, colors);
  }
  return colors;
}

function formatColor(code: string, value: number | string): string | undefined {
  if (!code.includes('[')) return undefined;
  const colors = sectionColors(code);
  if (typeof value === 'string') return colors[3];
  if (colors.length === 1) return colors[0];
  if (value < 0) return colors[1];
  if (value === 0 && colors.length > 2) return colors[2];
  return colors[0];
}

export function displayCell(wb: Workbook, cell: Cell, numFmt: string): CellDisplay {
  const kind = valueKind(cell.value);
  if (kind === 'empty') return { text: '', kind };
  const v = effectiveValue(cell.value);
  // The library's General keeps all 15 stored digits; a cell shows Excel's 11.
  const text = typeof v === 'number' && numFmt === 'General' ? generalText(v) : getCellDisplayText(wb, cell);
  if (typeof v === 'number') {
    const color = formatColor(numFmt, v);
    return color ? { text, kind, number: v, color } : { text, kind, number: v };
  }
  if (typeof v === 'string') {
    const color = formatColor(numFmt, v);
    return color ? { text, kind, color } : { text, kind };
  }
  return { text, kind };
}

/** `1.42857e-5` → `1.42857E-05`: Excel's exponent has at least two digits, and no trailing zeros. */
function excelExponential(n: number, digits: number): string {
  const [mantissa = '', exp = ''] = n.toExponential(digits).split('e');
  const m = mantissa.includes('.') ? mantissa.replace(/0+$/, '').replace(/\.$/, '') : mantissa;
  const sign = exp.startsWith('-') ? '-' : '+';
  return `${m}E${sign}${exp.replace(/^[+-]/, '').padStart(2, '0')}`;
}

const GENERAL_MAX_CHARS = 11;
const GENERAL_MIN_SIGNIFICANT = 6;

const significantDigits = (s: string): number => s.replace(/^[-0.]+/, '').replace(/\./, '').length;

/**
 * What General shows for `n` in a wide enough column: at most 11 characters
 * besides the sign (`3.141592654`, `0.333333333`), and scientific with six
 * significant digits once the fixed form would need 12 integer digits
 * (`1.23457E+11`) or keep fewer than six digits of a tiny value
 * (`1.42857E-05`). Measured in Excel 16 for Mac.
 */
export function generalText(n: number): string {
  if (n === 0 || !Number.isFinite(n)) return String(n);
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  const stored = Number(abs.toPrecision(15));
  const storedDigits = (abs.toExponential(14).split('e')[0] ?? '').replace('.', '').replace(/0+$/, '').length;
  const intDigits = Math.max(1, Math.floor(Math.log10(abs)) + 1);
  if (intDigits <= GENERAL_MAX_CHARS) {
    const decimals = Math.max(0, GENERAL_MAX_CHARS - 1 - intDigits);
    // Trim by hand: String(1e-7) would already be exponential.
    const fixed = abs.toFixed(decimals).replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '');
    const fits = !fixed.includes('e') && fixed.replace('.', '').length <= GENERAL_MAX_CHARS;
    if (fits && (Number(fixed) === stored || significantDigits(fixed) >= Math.min(GENERAL_MIN_SIGNIFICANT, storedDigits))) {
      return sign + fixed;
    }
  }
  return sign + excelExponential(abs, GENERAL_MIN_SIGNIFICANT - 1);
}

/**
 * Shorten a General-formatted number to fit `maxWidth`, the way Excel drops
 * decimals and then falls back to scientific notation before giving up with
 * `#`s. `measure` returns the pixel width of a string in the cell's font.
 */
export function fitGeneralNumber(n: number, maxWidth: number, measure: (s: string) => number): string {
  const full = generalText(n);
  if (measure(full) <= maxWidth) return full;
  const abs = Math.abs(n);
  if (abs >= 1e-4 && abs < 1e11) {
    const intDigits = Math.max(1, Math.floor(Math.log10(abs)) + 1);
    for (let decimals = Math.max(0, 10 - intDigits); decimals >= 0; decimals--) {
      const s = String(Number(n.toFixed(decimals)));
      if (measure(s) <= maxWidth) return s;
    }
  }
  for (let digits = 5; digits >= 0; digits--) {
    const s = excelExponential(n, digits);
    if (measure(s) <= maxWidth) return s;
  }
  return hashes(maxWidth, measure);
}

export function hashes(maxWidth: number, measure: (s: string) => number): string {
  const w = Math.max(1, measure('#'));
  return '#'.repeat(Math.max(1, Math.floor(maxWidth / w)));
}
