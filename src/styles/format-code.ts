// Interpreter for Excel number-format codes: a `formatCode` plus a value in,
// the text Excel prints in the cell out.
//
// ECMA-376 types `numFmt@formatCode` as a plain string (§18.8.30) and leaves
// the grammar to the application, so Excel's own rules are the specification
// here. A code is up to four `;`-separated sections chosen by the value's sign
// (positive, negative, zero, text), and each section mixes literals with
// placeholders: `0` `#` `?` for digits, `y` `m` `d` `h` `s` for date parts, `@`
// for text.
//
// `parseFormatCode` returns `undefined` for a code outside the subset
// `getCellDisplayText` documents, so callers degrade to a plain coercion rather
// than print a guess.

import { excelToDate, type ExcelEpoch } from '../utils/datetime.js';
import { FORMAT_COLOR_NAMES } from './numbers.js';

/** Excel carries 15 significant decimal digits and never shows more. */
const EXCEL_SIGNIFICANT_DIGITS = 15;
/** Positive, negative, zero, text. */
const MAX_SECTIONS = 4;
const GROUP_SEPARATOR = ',';
const GROUP_SIZE = 3;
const SCALE_DIVISOR = 1000;
const PERCENT_MULTIPLIER = 100;
/** `mmm` and wider render a month or weekday name, never a minute. */
const NAME_TOKEN_WIDTH = 3;
const MONTH_FULL_WIDTH = 4;
/** Excel pads a `General` exponent to two digits, so 1e-7 shows as `1E-07`. */
const GENERAL_EXPONENT_DIGITS = 2;
const FULL_YEAR_DIGITS = 4;
/** Excel stores time to the millisecond, so `.0000` gains nothing over `.000`. */
const MAX_SUBSECOND_DIGITS = 3;
const SECONDS_PER_DAY = 86_400;
const SECONDS_PER_HOUR = 3_600;
const SECONDS_PER_MINUTE = 60;
const MINUTES_PER_HOUR = 60;
const HOURS_PER_DAY = 24;
const HOURS_PER_HALF_DAY = 12;
const ZERO_CHAR_CODE = '0'.charCodeAt(0);
const FIVE_CHAR_CODE = '5'.charCodeAt(0);
const NINE_CHAR_CODE = '9'.charCodeAt(0);
/** Terms and closeness at which the fraction search stops chasing a double. */
const FRACTION_MAX_TERMS = 64;
const FRACTION_EPSILON = 1e-12;

const MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
] as const;

const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'] as const;
const JA_WEEKDAY_NAMES = ['日', '月', '火', '水', '木', '金', '土'] as const;
const JA_WEEKDAY_SUFFIX = '曜日';
const JA_MONTH_SUFFIX = '月';
const JA_AM = '午前';
const JA_PM = '午後';

/** A Japanese era: the Gregorian day it starts and the three ways `g` / `gg` / `ggg` spell it. */
interface JapaneseEra {
  readonly start: readonly [year: number, month: number, day: number];
  readonly names: readonly [latin: string, short: string, long: string];
}

// Newest first. Meiji is the era of every day the 1900 date system reaches before Taisho.
const JAPANESE_ERAS: readonly JapaneseEra[] = [
  { start: [2019, 5, 1], names: ['R', '令', '令和'] },
  { start: [1989, 1, 8], names: ['H', '平', '平成'] },
  { start: [1926, 12, 25], names: ['S', '昭', '昭和'] },
  { start: [1912, 7, 30], names: ['T', '大', '大正'] },
  { start: [1868, 1, 1], names: ['M', '明', '明治'] },
];

/** `[$-411]`, `[$-ja-JP]`, `[$¥-411]`: the locale part of the id names Japanese. */
const JAPANESE_LOCALE_RE = /-(?:[0-9a-f]*0?411|ja(?:-jp)?)(?:-|$)/i;

const COLOR_NAMES: ReadonlySet<string> = new Set(FORMAT_COLOR_NAMES.toLowerCase().split('|'));
const INDEXED_COLOR_RE = /^color\s*\d+$/i;
/**
 * The body of an elapsed-time bracket: one letter repeated, as in `[h]` or
 * `[mm]`. `isTimedeltaFormat` in `./numbers.js` answers the different question
 * of whether a whole code is a duration, so it cannot stand in here.
 */
const ELAPSED_RE = /^([hms])\1*$/i;

// ---- tokens ----------------------------------------------------------------

type DigitPlaceholder = '0' | '#' | '?';

type DatePart = 'year' | 'month' | 'day' | 'hour' | 'minute' | 'second' | 'monthOrMinute' | 'era' | 'eraYear';
type ElapsedPart = 'hour' | 'minute' | 'second';

interface DatePartToken {
  readonly kind: 'datePart';
  readonly part: DatePart;
  readonly width: number;
}

interface ElapsedToken {
  readonly kind: 'elapsed';
  readonly part: ElapsedPart;
  readonly width: number;
}

type Token =
  | { readonly kind: 'literal'; readonly text: string }
  | { readonly kind: 'general' }
  | { readonly kind: 'digit'; readonly placeholder: DigitPlaceholder }
  | { readonly kind: 'point' }
  | { readonly kind: 'comma' }
  | { readonly kind: 'percent' }
  | { readonly kind: 'exponent'; readonly explicitSign: boolean }
  | { readonly kind: 'slash' }
  | { readonly kind: 'fixedDenominator'; readonly text: string }
  | { readonly kind: 'textPlaceholder' }
  /** Excel prints the meridiem in the case the code carries, so `am/pm` stays lowercase. */
  | { readonly kind: 'meridiem'; readonly am: string; readonly pm: string }
  | { readonly kind: 'subsecond'; readonly digits: number }
  | DatePartToken
  | ElapsedToken;

// ---- sections --------------------------------------------------------------

interface NumberFormatSection {
  readonly kind: 'number';
  readonly tokens: readonly Token[];
  readonly grouped: boolean;
  /** Trailing commas; each one divides the value by a thousand. */
  readonly scale: number;
  /** `%` occurrences; each one multiplies the value by a hundred. */
  readonly percents: number;
  /** Token index where the integer region ends (exclusive). */
  readonly intEnd: number;
  readonly intPlaceholders: readonly DigitPlaceholder[];
  readonly fracPlaceholders: readonly DigitPlaceholder[];
  /** Token index of the `E+` / `E-` marker, or -1 outside scientific codes. */
  readonly exponentIndex: number;
  readonly exponentDigits: number;
  readonly exponentExplicitSign: boolean;
}

interface DateFormatSection {
  readonly kind: 'date';
  readonly tokens: readonly Token[];
  readonly elapsed: boolean;
  /** A `[$-411]`-style locale: month, weekday and AM/PM names print in Japanese. */
  readonly japanese: boolean;
  /** A year, month or day part: the section names a day, not just a time within one. */
  readonly calendar: boolean;
  readonly hasMeridiem: boolean;
  /** Widest fractional-second group in the section; the serial is rounded once at this precision. */
  readonly subsecondDigits: number;
}

interface FractionFormatSection {
  readonly kind: 'fraction';
  readonly tokens: readonly Token[];
  readonly slashIndex: number;
  /** Token range holding the whole-number run, or -1 when the code has none. */
  readonly wholeStart: number;
  readonly wholeEnd: number;
  readonly numeratorStart: number;
  readonly denominatorEnd: number;
  /** Placeholder count in the denominator run; caps the denominator at 10^n - 1. */
  readonly denominatorDigits: number;
  /** Non-zero when the code spells the denominator out, as in `# ?/16`. */
  readonly fixedDenominator: number;
}

interface TextFormatSection {
  readonly kind: 'text';
  readonly tokens: readonly Token[];
}

type FormatSection =
  | { readonly kind: 'general' }
  | { readonly kind: 'blank' }
  | { readonly kind: 'literal'; readonly text: string }
  | NumberFormatSection
  | DateFormatSection
  | FractionFormatSection
  | TextFormatSection;

/** A `[>=100]`-style comparison that picks a section by value instead of by sign. */
interface SectionCondition {
  readonly op: '<' | '<=' | '>' | '>=' | '=' | '<>';
  readonly operand: number;
}

export interface ParsedFormat {
  readonly sections: readonly FormatSection[];
  /** Per section; present only when the code has at least one comparison. */
  readonly conditions?: readonly (SectionCondition | undefined)[];
}

// ---- decimal arithmetic ----------------------------------------------------

interface DecimalParts {
  readonly int: string;
  readonly frac: string;
}

/**
 * Plain (never exponential) decimal digits of a finite non-negative number at
 * Excel's precision. Rounding has to go through this string: `(1.005).toFixed(2)`
 * is `"1.00"` because the double nearest 1.005 sits just below it, while Excel
 * rounds the 15-digit decimal it shows and prints `1.01`.
 */
const toPlainDecimal = (value: number): string => {
  const text = value.toPrecision(EXCEL_SIGNIFICANT_DIGITS);
  const parsed = /^(\d+)(?:\.(\d+))?e([+-]\d+)$/.exec(text);
  if (parsed === null) return text;
  const intPart = parsed[1] ?? '';
  const digits = intPart + (parsed[2] ?? '');
  const point = intPart.length + Number(parsed[3] ?? '0');
  if (point <= 0) return `0.${'0'.repeat(-point)}${digits}`;
  if (point >= digits.length) return digits + '0'.repeat(point - digits.length);
  return `${digits.slice(0, point)}.${digits.slice(point)}`;
};

/** Add one to the last digit of a digit string, growing it on overflow. */
const carryOne = (digits: string): string => {
  let i = digits.length - 1;
  let tail = '';
  while (i >= 0 && digits.charCodeAt(i) === NINE_CHAR_CODE) {
    tail = `0${tail}`;
    i--;
  }
  if (i < 0) return `1${tail}`;
  return digits.slice(0, i) + String(digits.charCodeAt(i) - ZERO_CHAR_CODE + 1) + tail;
};

/** Round `int`.`frac` to `decimals` places, half away from zero. */
const roundParts = (int: string, frac: string, decimals: number): DecimalParts => {
  if (frac.length <= decimals) return { int, frac: frac.padEnd(decimals, '0') };
  const kept = frac.slice(0, decimals);
  if (frac.charCodeAt(decimals) < FIVE_CHAR_CODE) return { int, frac: kept };
  const bumped = carryOne(int + kept);
  return { int: bumped.slice(0, bumped.length - decimals), frac: bumped.slice(bumped.length - decimals) };
};

const roundDecimal = (value: number, decimals: number): DecimalParts => {
  const text = toPlainDecimal(value);
  const dot = text.indexOf('.');
  if (dot === -1) return roundParts(text, '', decimals);
  return roundParts(text.slice(0, dot), text.slice(dot + 1), decimals);
};

/** Significant digits of a non-negative number, most significant first. */
const significantDigits = (value: number): string => toPlainDecimal(value).replace('.', '').replace(/^0+/, '');

/** Power of ten of the leading significant digit; 0 for a zero value. */
const decimalOrder = (value: number): number => {
  const text = toPlainDecimal(value);
  const dot = text.indexOf('.');
  const int = (dot === -1 ? text : text.slice(0, dot)).replace(/^0+/, '');
  if (int.length > 0) return int.length - 1;
  const frac = dot === -1 ? '' : text.slice(dot + 1);
  const leading = frac.search(/[1-9]/);
  return leading === -1 ? 0 : -(leading + 1);
};

// ---- tokenizer -------------------------------------------------------------

const pushLiteral = (tokens: Token[], text: string): void => {
  const last = tokens[tokens.length - 1];
  if (last?.kind === 'literal') tokens[tokens.length - 1] = { kind: 'literal', text: last.text + text };
  else tokens.push({ kind: 'literal', text });
};

const datePartOf = (letter: string): DatePart => {
  if (letter === 'y') return 'year';
  if (letter === 'd') return 'day';
  if (letter === 'h') return 'hour';
  if (letter === 's') return 'second';
  if (letter === 'g') return 'era';
  if (letter === 'e') return 'eraYear';
  return 'monthOrMinute';
};

/** What the brackets of a section said beyond the tokens they produced. */
interface SectionContext {
  japanese: boolean;
}

/** `[Red]`, `[$€-407]`, `[hh]`, `[>100]`: everything Excel puts in brackets. */
const tokenizeBracket = (body: string, tokens: Token[], context: SectionContext): boolean => {
  const elapsed = ELAPSED_RE.exec(body);
  if (elapsed !== null) {
    const letter = (elapsed[1] ?? 'h').toLowerCase();
    const part: ElapsedPart = letter === 'h' ? 'hour' : letter === 'm' ? 'minute' : 'second';
    tokens.push({ kind: 'elapsed', part, width: body.length });
    return true;
  }
  if (body.startsWith('$')) {
    // `[$<symbol>-<locale>]`: the symbol prints, and a Japanese locale switches
    // month, weekday and AM/PM names to Japanese. Other locales keep English.
    const symbol = body.slice(1).split('-')[0] ?? '';
    if (symbol.length > 0) pushLiteral(tokens, symbol);
    if (JAPANESE_LOCALE_RE.test(body)) context.japanese = true;
    return true;
  }
  const lower = body.toLowerCase();
  if (COLOR_NAMES.has(lower) || INDEXED_COLOR_RE.test(body)) return true;
  // Comparison sections (`[>=100]"big"`) and calendar modifiers (`[DBNum1]`)
  // change the value that prints, so a code carrying one is out of scope.
  return false;
};

const MERIDIEM_LONG = 'AM/PM';
const MERIDIEM_SHORT = 'A/P';

const meridiemToken = (raw: string): Token => {
  const slash = raw.indexOf('/');
  return { kind: 'meridiem', am: raw.slice(0, slash), pm: raw.slice(slash + 1) };
};

const tokenizeSection = (src: string, context: SectionContext): Token[] | undefined => {
  const tokens: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src.charAt(i);
    if (ch === '"') {
      const end = src.indexOf('"', i + 1);
      if (end === -1) return undefined;
      pushLiteral(tokens, src.slice(i + 1, end));
      i = end + 1;
      continue;
    }
    if (ch === '\\') {
      pushLiteral(tokens, src.charAt(i + 1));
      i += 2;
      continue;
    }
    if (ch === '_') {
      // `_x` reserves the width of `x`; the cell shows a space there.
      pushLiteral(tokens, ' ');
      i += 2;
      continue;
    }
    if (ch === '*') {
      // `*x` repeats `x` until the column is full, so how much it prints is not
      // a property of the cell. Nothing is emitted for it.
      i += 2;
      continue;
    }
    if (ch === '[') {
      const end = src.indexOf(']', i + 1);
      if (end === -1) return undefined;
      if (!tokenizeBracket(src.slice(i + 1, end), tokens, context)) return undefined;
      i = end + 1;
      continue;
    }
    if (ch === '0' || ch === '#' || ch === '?') {
      tokens.push({ kind: 'digit', placeholder: ch });
      i++;
      continue;
    }
    if (ch === '.') {
      tokens.push({ kind: 'point' });
      i++;
      continue;
    }
    if (ch === ',') {
      tokens.push({ kind: 'comma' });
      i++;
      continue;
    }
    if (ch === '%') {
      tokens.push({ kind: 'percent' });
      i++;
      continue;
    }
    if (ch === '/') {
      tokens.push({ kind: 'slash' });
      i++;
      // A denominator beginning with 1–9 is a fixed integer, including zeros.
      const fixed = /^[1-9][0-9]*/.exec(src.slice(i));
      if (fixed !== null) {
        tokens.push({ kind: 'fixedDenominator', text: fixed[0] });
        i += fixed[0].length;
      }
      continue;
    }
    if (ch === '@') {
      tokens.push({ kind: 'textPlaceholder' });
      i++;
      continue;
    }
    const long = src.slice(i, i + MERIDIEM_LONG.length);
    if (long.toUpperCase() === MERIDIEM_LONG) {
      tokens.push(meridiemToken(long));
      i += MERIDIEM_LONG.length;
      continue;
    }
    const short = src.slice(i, i + MERIDIEM_SHORT.length);
    if (short.toUpperCase() === MERIDIEM_SHORT) {
      tokens.push(meridiemToken(short));
      i += MERIDIEM_SHORT.length;
      continue;
    }
    if (src.slice(i, i + 'General'.length).toLowerCase() === 'general') {
      tokens.push({ kind: 'general' });
      i += 'General'.length;
      continue;
    }
    const lower = ch.toLowerCase();
    if (lower === 'e') {
      const sign = src.charAt(i + 1);
      if (sign === '+' || sign === '-') {
        tokens.push({ kind: 'exponent', explicitSign: sign === '+' });
        i += 2;
        continue;
      }
    }
    // `aaa` / `aaaa` are the weekday in the section's language, like `ddd` / `dddd`.
    if (lower === 'a' && src.slice(i, i + NAME_TOKEN_WIDTH).toLowerCase() === 'aaa') {
      let width = NAME_TOKEN_WIDTH;
      while (src.charAt(i + width).toLowerCase() === 'a') width++;
      tokens.push({ kind: 'datePart', part: 'day', width });
      i += width;
      continue;
    }
    // `g` is the Japanese era and a bare `e` the year within it; Excel draws
    // both from the Japanese calendar whatever the section's locale.
    if (lower === 'y' || lower === 'm' || lower === 'd' || lower === 'h' || lower === 's' || lower === 'g' || lower === 'e') {
      let width = 1;
      while (src.charAt(i + width).toLowerCase() === lower) width++;
      tokens.push({ kind: 'datePart', part: datePartOf(lower), width });
      i += width;
      continue;
    }
    // `b` (Buddhist calendar) shifts the calendar system.
    if (lower === 'b') return undefined;
    pushLiteral(tokens, ch);
    i++;
  }
  return tokens;
};

// ---- section parsing -------------------------------------------------------

const splitSections = (code: string): string[] | undefined => {
  const out: string[] = [];
  let start = 0;
  let i = 0;
  while (i < code.length) {
    const ch = code.charAt(i);
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      if (end === -1) return undefined;
      i = end + 1;
      continue;
    }
    if (ch === '[') {
      const end = code.indexOf(']', i + 1);
      if (end === -1) return undefined;
      i = end + 1;
      continue;
    }
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
      continue;
    }
    if (ch === ';') {
      out.push(code.slice(start, i));
      start = i + 1;
    }
    i++;
  }
  out.push(code.slice(start));
  return out;
};

const literalFor = (token: Token): string => {
  switch (token.kind) {
    case 'literal':
    case 'fixedDenominator':
      return token.text;
    case 'percent':
      return '%';
    case 'slash':
      return '/';
    case 'point':
      return '.';
    case 'comma':
      return GROUP_SEPARATOR;
    default:
      return '';
  }
};

const onlyCommasBetween = (tokens: readonly Token[], from: number, to: number): boolean => {
  for (let i = from + 1; i < to; i++) {
    if (tokens[i]?.kind !== 'comma') return false;
  }
  return true;
};

interface CommaResolution {
  readonly tokens: Token[];
  readonly grouped: boolean;
  readonly scale: number;
}

/**
 * A comma means three different things depending on where it sits: thousands
 * grouping between digit placeholders, division by a thousand after the last
 * one, and a plain comma anywhere else.
 */
const resolveCommas = (tokens: readonly Token[]): CommaResolution => {
  const isDigit = tokens.map((t) => t.kind === 'digit');
  const firstDigit = isDigit.indexOf(true);
  const lastDigit = isDigit.lastIndexOf(true);
  const out: Token[] = [];
  let grouped = false;
  let scale = 0;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token === undefined) continue;
    if (token.kind !== 'comma') {
      out.push(token);
      continue;
    }
    if (firstDigit !== -1 && i > firstDigit && i < lastDigit) {
      grouped = true;
      continue;
    }
    if (lastDigit !== -1 && i > lastDigit && onlyCommasBetween(tokens, lastDigit, i)) {
      scale++;
      continue;
    }
    pushLiteral(out, GROUP_SEPARATOR);
  }
  return { tokens: out, grouped, scale };
};

const placeholdersIn = (tokens: readonly Token[], from: number, to: number): DigitPlaceholder[] => {
  const out: DigitPlaceholder[] = [];
  for (let i = from; i < to; i++) {
    const token = tokens[i];
    if (token?.kind === 'digit') out.push(token.placeholder);
  }
  return out;
};

const resolveNumberSection = (resolved: CommaResolution): NumberFormatSection | undefined => {
  const { tokens } = resolved;
  // A second decimal point or exponent splices two numeric layouts into one
  // section (`0.00" ("0.00")"`), leaving no single number to lay out.
  if (tokens.filter((t) => t.kind === 'point').length > 1) return undefined;
  if (tokens.filter((t) => t.kind === 'exponent').length > 1) return undefined;
  const pointIndex = tokens.findIndex((t) => t.kind === 'point');
  const exponentIndex = tokens.findIndex((t) => t.kind === 'exponent');
  const intEnd = pointIndex !== -1 ? pointIndex : exponentIndex !== -1 ? exponentIndex : tokens.length;
  const fracEnd = exponentIndex !== -1 ? exponentIndex : tokens.length;
  const exponentToken = exponentIndex === -1 ? undefined : tokens[exponentIndex];
  return {
    kind: 'number',
    tokens,
    grouped: resolved.grouped,
    scale: resolved.scale,
    percents: tokens.filter((t) => t.kind === 'percent').length,
    intEnd,
    intPlaceholders: placeholdersIn(tokens, 0, intEnd),
    fracPlaceholders: pointIndex === -1 ? [] : placeholdersIn(tokens, pointIndex + 1, fracEnd),
    exponentIndex,
    exponentDigits: exponentIndex === -1 ? 0 : placeholdersIn(tokens, exponentIndex + 1, tokens.length).length,
    exponentExplicitSign: exponentToken?.kind === 'exponent' ? exponentToken.explicitSign : false,
  };
};

const digitRunStart = (tokens: readonly Token[], end: number): number => {
  let start = end;
  while (start > 0 && tokens[start - 1]?.kind === 'digit') start--;
  return start;
};

/** `# ??/??`: a slash with digit placeholders on both sides is a fraction. */
const resolveFractionSection = (tokens: readonly Token[]): FractionFormatSection | undefined => {
  const slashIndex = tokens.findIndex((t) => t.kind === 'slash');
  if (slashIndex < 1 || tokens[slashIndex - 1]?.kind !== 'digit') return undefined;

  const numeratorStart = digitRunStart(tokens, slashIndex);
  const after = tokens[slashIndex + 1];
  let denominatorEnd: number;
  let fixedDenominator = 0;
  if (after?.kind === 'digit') {
    denominatorEnd = slashIndex + 1;
    while (tokens[denominatorEnd]?.kind === 'digit') denominatorEnd++;
  } else if (after?.kind === 'fixedDenominator') {
    denominatorEnd = slashIndex + 2;
    fixedDenominator = Number(after.text);
    if (!Number.isSafeInteger(fixedDenominator) || fixedDenominator === 0) return undefined;
  } else {
    return undefined;
  }

  let wholeEnd = -1;
  let wholeStart = -1;
  for (let i = numeratorStart - 1; i >= 0; i--) {
    if (tokens[i]?.kind !== 'digit') continue;
    wholeEnd = i + 1;
    wholeStart = digitRunStart(tokens, wholeEnd);
    break;
  }

  return {
    kind: 'fraction',
    tokens,
    slashIndex,
    wholeStart,
    wholeEnd,
    numeratorStart,
    denominatorEnd,
    denominatorDigits: fixedDenominator > 0 ? 0 : placeholdersIn(tokens, slashIndex + 1, denominatorEnd).length,
    fixedDenominator,
  };
};

const nextDateToken = (tokens: readonly Token[], index: number, step: number): DatePartToken | ElapsedToken | undefined => {
  for (let i = index + step; i >= 0 && i < tokens.length; i += step) {
    const token = tokens[i];
    if (token === undefined) return undefined;
    if (token.kind === 'datePart' || token.kind === 'elapsed') return token;
  }
  return undefined;
};

const resolveDateSection = (tokens: readonly Token[], japanese: boolean): DateFormatSection | undefined => {
  // `.0` right after a seconds token is fractional seconds, not a decimal point.
  const withSubseconds: Token[] = [];
  let subsecondDigits = 0;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token === undefined) continue;
    if (token.kind !== 'point') {
      withSubseconds.push(token);
      continue;
    }
    let digits = 0;
    for (let j = i + 1; j < tokens.length; j++) {
      const next = tokens[j];
      if (next?.kind !== 'digit' || next.placeholder !== '0') break;
      digits++;
    }
    const previous = withSubseconds[withSubseconds.length - 1];
    const afterSeconds = previous?.kind === 'datePart' || previous?.kind === 'elapsed' ? previous.part === 'second' : false;
    if (digits === 0 || !afterSeconds) {
      pushLiteral(withSubseconds, '.');
      continue;
    }
    withSubseconds.push({ kind: 'subsecond', digits });
    subsecondDigits = Math.max(subsecondDigits, digits);
    i += digits;
  }

  // Excel reads `m` as a minute next to an hour or a second, and as a month
  // everywhere else. `mmm` and wider are always month or weekday names.
  const resolved: Token[] = withSubseconds.map((token, index) => {
    if (token.kind !== 'datePart' || token.part !== 'monthOrMinute') return token;
    if (token.width >= NAME_TOKEN_WIDTH) return { kind: 'datePart', part: 'month', width: token.width };
    const previous = nextDateToken(withSubseconds, index, -1);
    const next = nextDateToken(withSubseconds, index, 1);
    const minute = previous?.part === 'hour' || next?.part === 'second';
    return { kind: 'datePart', part: minute ? 'minute' : 'month', width: token.width };
  });

  const elapsed = resolved.some((t) => t.kind === 'elapsed');
  const calendarPart = resolved.some(
    (t) =>
      t.kind === 'datePart' &&
      (t.part === 'year' || t.part === 'month' || t.part === 'day' || t.part === 'era' || t.part === 'eraYear'),
  );
  // A leftover digit placeholder means the code mixes a numeric layout into a
  // date layout, and an elapsed duration has no calendar date to print.
  if (resolved.some((t) => t.kind === 'digit' || t.kind === 'exponent')) return undefined;
  if (elapsed && calendarPart) return undefined;

  return {
    kind: 'date',
    tokens: resolved,
    elapsed,
    japanese,
    calendar: calendarPart,
    hasMeridiem: resolved.some((t) => t.kind === 'meridiem'),
    subsecondDigits,
  };
};

const parseSection = (src: string): FormatSection | undefined => {
  const context: SectionContext = { japanese: false };
  const tokens = tokenizeSection(src, context);
  if (tokens === undefined) return undefined;
  if (tokens.length === 0) return { kind: 'blank' };
  if (tokens.some((t) => t.kind === 'general')) {
    // `General` is a whole format, not a placeholder that combines with others.
    return tokens.length === 1 ? { kind: 'general' } : undefined;
  }
  const hasText = tokens.some((t) => t.kind === 'textPlaceholder');
  const hasDate = tokens.some((t) => t.kind === 'datePart' || t.kind === 'elapsed' || t.kind === 'meridiem');
  if (hasText && hasDate) return undefined;
  if (hasText) return { kind: 'text', tokens };
  if (hasDate) return resolveDateSection(tokens, context.japanese);
  if (tokens.every((token) => token.kind === 'literal')) {
    return { kind: 'literal', text: tokens.map(literalFor).join('') };
  }
  const resolved = resolveCommas(tokens);
  return resolveFractionSection(resolved.tokens) ?? resolveNumberSection(resolved);
};

/**
 * Distinct format codes in a workbook number in the tens while cells number in
 * the millions, so the parse is memoised. Both bounds exist because the codes
 * arrive from an untrusted `styles.xml`: a full map drops its oldest entry
 * rather than every entry, so a workbook carrying more codes than fit loses one
 * cached parse per new code instead of all of them, and a code longer than the
 * 255 characters Excel itself allows is parsed on each call rather than stored.
 */
const PARSE_CACHE_LIMIT = 256;
const MAX_CACHED_CODE_LENGTH = 255;
const parseCache = new Map<string, { readonly format: ParsedFormat | undefined }>();

const CONDITION_RE = /^\[(<=|>=|<>|<|>|=)\s*(-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\]/;
const LEADING_BRACKET_RE = /^\[[^\]]*\]/;

/** Pull a comparison out of the brackets leading a section (`[Red][<0]0`). */
const takeCondition = (src: string): { src: string; condition?: SectionCondition } => {
  let rest = src;
  let kept = '';
  while (rest.startsWith('[')) {
    const cond = CONDITION_RE.exec(rest);
    if (cond) {
      const op = cond[1] as SectionCondition['op'];
      return { src: kept + rest.slice(cond[0].length), condition: { op, operand: Number(cond[2]) } };
    }
    const bracket = LEADING_BRACKET_RE.exec(rest);
    if (!bracket) break;
    kept += bracket[0];
    rest = rest.slice(bracket[0].length);
  }
  return { src };
};

const buildFormat = (code: string): ParsedFormat | undefined => {
  const raw = splitSections(code);
  if (raw === undefined || raw.length > MAX_SECTIONS) return undefined;
  const sections: FormatSection[] = [];
  const conditions: (SectionCondition | undefined)[] = [];
  for (const piece of raw) {
    const { src, condition } = takeCondition(piece);
    const section = parseSection(src);
    if (section === undefined) return undefined;
    sections.push(section);
    conditions.push(condition);
  }
  return conditions.some((c) => c !== undefined) ? { sections, conditions } : { sections };
};

/** Parse a `numFmt` format code, or `undefined` when it is out of scope. */
export function parseFormatCode(code: string): ParsedFormat | undefined {
  const cached = parseCache.get(code);
  if (cached !== undefined) return cached.format;
  const format = buildFormat(code);
  if (code.length > MAX_CACHED_CODE_LENGTH) return format;
  if (parseCache.size >= PARSE_CACHE_LIMIT) {
    const oldest = parseCache.keys().next().value;
    if (oldest !== undefined) parseCache.delete(oldest);
  }
  parseCache.set(code, { format });
  return format;
}

/**
 * True when the format reads as a calendar date: a date section carrying a
 * year, month or day, rather than a time of day (`h:mm`), an elapsed span
 * (`[h]:mm`) or a numeric layout. The section selected for the value decides.
 */
export function hasCalendarDate(format: ParsedFormat, value: number): boolean {
  const first = pickSection(format, value)?.section;
  return first?.kind === 'date' && first.calendar;
}

// ---- rendering: numeric layouts -------------------------------------------

const padFor = (placeholder: DigitPlaceholder): string => {
  if (placeholder === '0') return '0';
  return placeholder === '?' ? ' ' : '';
};

const withGrouping = (digits: string, grouped: boolean): string => {
  if (!grouped || digits.length <= GROUP_SIZE) return digits;
  const groups: string[] = [];
  for (let end = digits.length; end > 0; end -= GROUP_SIZE) {
    groups.unshift(digits.slice(Math.max(0, end - GROUP_SIZE), end));
  }
  return groups.join(GROUP_SEPARATOR);
};

/**
 * Render `[from, to)` right to left, feeding `digits` into the placeholders from
 * the low-order end. The leftmost placeholder absorbs whatever is left over,
 * which is how `#,##0` still prints all seven digits of 1,234,567.
 */
const renderDigitsRightToLeft = (tokens: readonly Token[], from: number, to: number, digits: string): string => {
  const leftmost = placeholderIndex(tokens, from, to);
  const pieces: string[] = [];
  let pos = digits.length;
  for (let i = to - 1; i >= from; i--) {
    const token = tokens[i];
    if (token === undefined) continue;
    if (token.kind !== 'digit') {
      pieces.push(literalFor(token));
      continue;
    }
    if (i === leftmost) {
      pieces.push(pos > 0 ? digits.slice(0, pos) : padFor(token.placeholder));
      pos = 0;
      continue;
    }
    if (pos === 0) {
      pieces.push(padFor(token.placeholder));
      continue;
    }
    let take = pos - 1;
    // A grouping separator belongs to the digit on its right.
    while (take > 0 && digits.charAt(take - 1) === GROUP_SEPARATOR) take--;
    pieces.push(digits.slice(take, pos));
    pos = take;
  }
  pieces.reverse();
  return pieces.join('');
};

const placeholderIndex = (tokens: readonly Token[], from: number, to: number): number => {
  for (let i = from; i < to; i++) {
    if (tokens[i]?.kind === 'digit') return i;
  }
  return -1;
};

/** `#` drops trailing zeros, `?` blanks them, `0` keeps them. */
const renderFracDigits = (placeholders: readonly DigitPlaceholder[], digits: string): string => {
  const pieces: string[] = [];
  let trailing = true;
  for (let i = placeholders.length - 1; i >= 0; i--) {
    const placeholder = placeholders[i];
    if (placeholder === undefined) continue;
    const digit = digits.charAt(i);
    if (trailing && digit === '0' && placeholder !== '0') {
      pieces.push(placeholder === '?' ? ' ' : '');
      continue;
    }
    trailing = false;
    pieces.push(digit);
  }
  pieces.reverse();
  return pieces.join('');
};

const assembleNumber = (section: NumberFormatSection, parts: DecimalParts, exponentText: string): string => {
  const minIntDigits = section.intPlaceholders.filter((p) => p === '0').length;
  const intDigits = withGrouping(parts.int.replace(/^0+/, '').padStart(minIntDigits, '0'), section.grouped);
  const head = renderDigitsRightToLeft(section.tokens, 0, section.intEnd, intDigits);
  // `.00` has no integer placeholder to feed the digits into, and Excel prints
  // them anyway: immediately left of the point, with nothing for a zero.
  const intText = section.intPlaceholders.length === 0 ? head + intDigits : head;
  const fracText = renderFracDigits(section.fracPlaceholders, parts.frac);

  const tail: string[] = [];
  let fracEmitted = false;
  for (let i = section.intEnd; i < section.tokens.length; i++) {
    const token = section.tokens[i];
    if (token === undefined) continue;
    if (token.kind === 'point') {
      tail.push('.');
      continue;
    }
    if (token.kind === 'exponent') {
      tail.push(exponentText);
      continue;
    }
    if (token.kind === 'digit') {
      // Exponent digits already went into `exponentText`.
      if (section.exponentIndex !== -1 && i > section.exponentIndex) continue;
      if (!fracEmitted) {
        tail.push(fracText);
        fracEmitted = true;
      }
      continue;
    }
    tail.push(literalFor(token));
  }
  return intText + tail.join('');
};

interface RenderedSection {
  readonly text: string;
  /** True when every digit the section printed is a zero. */
  readonly zero: boolean;
}

const partsAreZero = (parts: DecimalParts): boolean => !/[1-9]/.test(parts.int + parts.frac);

const renderScientificSection = (section: NumberFormatSection, magnitude: number): RenderedSection => {
  const decimals = section.fracPlaceholders.length;
  // `##0.0E+0` asks for up to three integer digits, so the exponent moves in
  // steps of three: engineering notation.
  const step = Math.max(1, section.intPlaceholders.length);
  let digits = significantDigits(magnitude);
  let order = decimalOrder(magnitude);
  for (;;) {
    const exponent = magnitude === 0 ? 0 : Math.floor(order / step) * step;
    const intWidth = order - exponent + 1;
    const parts = roundParts(digits.slice(0, intWidth).padEnd(intWidth, '0') || '0', digits.slice(intWidth), decimals);
    if (parts.int.length > intWidth) {
      // Rounding pushed the mantissa to the next power of ten (9.99 became
      // 10.0 at one integer digit); step the exponent instead of widening.
      order += 1;
      digits = '1';
      continue;
    }
    const sign = exponent < 0 ? '-' : section.exponentExplicitSign ? '+' : '';
    const exponentText = `E${sign}${String(Math.abs(exponent)).padStart(section.exponentDigits, '0')}`;
    return { text: assembleNumber(section, parts, exponentText), zero: partsAreZero(parts) };
  }
};

const renderNumberSection = (section: NumberFormatSection, magnitude: number): RenderedSection | undefined => {
  const scaled = (magnitude * PERCENT_MULTIPLIER ** section.percents) / SCALE_DIVISOR ** section.scale;
  if (!Number.isFinite(scaled)) return undefined;
  if (section.exponentIndex !== -1) return renderScientificSection(section, scaled);
  const parts = roundDecimal(scaled, section.fracPlaceholders.length);
  return { text: assembleNumber(section, parts, ''), zero: partsAreZero(parts) };
};

/**
 * Closest `p/q` with `q <= maxDenominator`. Continued-fraction convergents alone
 * are not always the closest (0.7 gives 2/3 before 5/7), so the last step also
 * weighs the best semiconvergent.
 */
const bestFraction = (value: number, maxDenominator: number): { numerator: number; denominator: number } => {
  let numeratorPrev = 0;
  let numerator = 1;
  let denominatorPrev = 1;
  let denominator = 0;
  let best = { numerator: Math.round(value), denominator: 1 };
  let x = value;
  for (let term = 0; term < FRACTION_MAX_TERMS; term++) {
    const a = Math.floor(x);
    const nextNumerator = a * numerator + numeratorPrev;
    const nextDenominator = a * denominator + denominatorPrev;
    if (nextDenominator > maxDenominator) {
      const steps = Math.floor((maxDenominator - denominatorPrev) / denominator);
      if (steps > 0) {
        const candidate = { numerator: steps * numerator + numeratorPrev, denominator: steps * denominator + denominatorPrev };
        if (Math.abs(value - candidate.numerator / candidate.denominator) < Math.abs(value - best.numerator / best.denominator)) {
          best = candidate;
        }
      }
      break;
    }
    numeratorPrev = numerator;
    numerator = nextNumerator;
    denominatorPrev = denominator;
    denominator = nextDenominator;
    best = { numerator, denominator };
    const remainder = x - a;
    if (remainder < FRACTION_EPSILON) break;
    x = 1 / remainder;
  }
  return best;
};

const blankFor = (placeholder: DigitPlaceholder): string => (placeholder === '#' ? '' : ' ');

const renderDigitsLeftToRight = (tokens: readonly Token[], from: number, to: number, digits: string): string => {
  const pieces: string[] = [];
  let pos = 0;
  const lastPlaceholder = lastPlaceholderIndex(tokens, from, to);
  for (let i = from; i < to; i++) {
    const token = tokens[i];
    if (token === undefined) continue;
    if (token.kind !== 'digit') {
      pieces.push(literalFor(token));
      continue;
    }
    if (i === lastPlaceholder) {
      pieces.push(pos < digits.length ? digits.slice(pos) : padFor(token.placeholder));
      pos = digits.length;
      continue;
    }
    if (pos < digits.length) {
      pieces.push(digits.charAt(pos));
      pos++;
      continue;
    }
    pieces.push(padFor(token.placeholder));
  }
  return pieces.join('');
};

const lastPlaceholderIndex = (tokens: readonly Token[], from: number, to: number): number => {
  for (let i = to - 1; i >= from; i--) {
    if (tokens[i]?.kind === 'digit') return i;
  }
  return -1;
};

const blankRange = (tokens: readonly Token[], from: number, to: number): string => {
  const pieces: string[] = [];
  for (let i = from; i < to; i++) {
    const token = tokens[i];
    if (token === undefined) continue;
    pieces.push(token.kind === 'digit' ? blankFor(token.placeholder) : token.kind === 'slash' ? ' ' : literalFor(token));
  }
  return pieces.join('');
};

/** The digits a code spells the denominator out as: `16` in `# ?/16`. */
const fixedDenominatorText = (section: FractionFormatSection): string => {
  const token = section.tokens[section.slashIndex + 1];
  return token?.kind === 'fixedDenominator' ? token.text : '';
};

/**
 * `?` and `#` pad the denominator on the right, which is what lines the slash up
 * down a column. A `0` cannot: a pad zero on the right would multiply the
 * denominator by ten, so a run carrying one fills from the right instead.
 */
const denominatorText = (section: FractionFormatSection, denominator: number): string => {
  const from = section.slashIndex + 1;
  const digits = String(denominator);
  return placeholdersIn(section.tokens, from, section.denominatorEnd).includes('0')
    ? renderDigitsRightToLeft(section.tokens, from, section.denominatorEnd, digits)
    : renderDigitsLeftToRight(section.tokens, from, section.denominatorEnd, digits);
};

const renderFractionSection = (section: FractionFormatSection, magnitude: number): RenderedSection => {
  const hasWhole = section.wholeStart !== -1;
  const maxDenominator =
    section.fixedDenominator > 0 ? section.fixedDenominator : 10 ** section.denominatorDigits - 1;
  let whole = hasWhole ? Math.floor(magnitude) : 0;
  const remainder = magnitude - whole;
  const approximation =
    section.fixedDenominator > 0
      ? { numerator: Math.round(remainder * section.fixedDenominator), denominator: section.fixedDenominator }
      : bestFraction(remainder, Math.max(1, maxDenominator));
  const denominator = approximation.denominator;
  let numerator = approximation.numerator;
  if (hasWhole && numerator === denominator && numerator !== 0) {
    whole += 1;
    numerator = 0;
  }

  const pieces: string[] = [];
  if (hasWhole) {
    pieces.push(renderDigitsRightToLeft(section.tokens, 0, section.wholeStart, ''));
    const wholeDigits = whole === 0 ? '' : String(whole);
    pieces.push(renderDigitsRightToLeft(section.tokens, section.wholeStart, section.wholeEnd, wholeDigits));
    pieces.push(blankRange(section.tokens, section.wholeEnd, section.numeratorStart));
  } else {
    pieces.push(renderDigitsRightToLeft(section.tokens, 0, section.numeratorStart, ''));
  }

  if (numerator === 0 && hasWhole) {
    // An exact whole number leaves the fraction blank, the way Excel does. A
    // spelled-out denominator goes with it: it is a literal rather than a run of
    // placeholders, so blanking only the placeholders leaves a stray `16` behind.
    pieces.push(blankRange(section.tokens, section.numeratorStart, section.slashIndex + 1));
    pieces.push(
      section.fixedDenominator > 0
        ? ' '.repeat(fixedDenominatorText(section).length)
        : blankRange(section.tokens, section.slashIndex + 1, section.denominatorEnd),
    );
  } else {
    pieces.push(renderDigitsRightToLeft(section.tokens, section.numeratorStart, section.slashIndex, String(numerator)));
    pieces.push('/');
    pieces.push(section.fixedDenominator > 0 ? fixedDenominatorText(section) : denominatorText(section, denominator));
  }
  pieces.push(renderDigitsRightToLeft(section.tokens, section.denominatorEnd, section.tokens.length, ''));
  return { text: pieces.join(''), zero: whole === 0 && numerator === 0 };
};

/**
 * Excel's `General`: the value at the 15 significant digits Excel stores, which
 * is what turns `1.1 + 2.2` into `3.3` instead of `3.3000000000000003`. How
 * many of those digits Excel shows, and where it gives up and switches to
 * scientific notation, also depends on the column width, which is not part of
 * the cell data, so every stored digit is kept and the notation switches where
 * JavaScript's own shortest round-trip form does.
 */
const renderGeneral = (magnitude: number): string => {
  const plain = String(Number(magnitude.toPrecision(EXCEL_SIGNIFICANT_DIGITS)));
  const exponential = /^(\d(?:\.\d+)?)e([+-])(\d+)$/.exec(plain);
  if (exponential === null) return plain;
  return `${exponential[1] ?? ''}E${exponential[2] ?? ''}${(exponential[3] ?? '').padStart(GENERAL_EXPONENT_DIGITS, '0')}`;
};

// ---- rendering: date layouts ----------------------------------------------

interface DateFields {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly weekday: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
  readonly totalSeconds: number;
  readonly subsecond: string;
}

const PHANTOM_YEAR = 1900;
/** Serial → the fields Excel prints for it (weekday 0 = Sunday, as WEEKDAY() - 1 reports). */
const PHANTOM_DAYS: ReadonlyMap<number, { month: number; day: number; weekday: number }> = new Map([
  [0, { month: 1, day: 0, weekday: 6 }],
  [60, { month: 2, day: 29, weekday: 3 }],
]);

const twoDigits = (value: number, width: number): string => (width === 1 ? String(value) : String(value).padStart(2, '0'));

const japaneseEraOf = (fields: DateFields): { era: JapaneseEra; year: number } => {
  const key = fields.year * 10_000 + fields.month * 100 + fields.day;
  const era =
    JAPANESE_ERAS.find(({ start: [y, m, d] }) => key >= y * 10_000 + m * 100 + d) ??
    (JAPANESE_ERAS[JAPANESE_ERAS.length - 1] as JapaneseEra);
  return { era, year: fields.year - era.start[0] + 1 };
};

const renderDatePart = (token: DatePartToken, fields: DateFields, section: DateFormatSection): string => {
  switch (token.part) {
    case 'era': {
      const { names } = japaneseEraOf(fields).era;
      return names[Math.min(token.width, names.length) - 1] ?? '';
    }
    case 'eraYear':
      return twoDigits(japaneseEraOf(fields).year, token.width);
    case 'year':
      return token.width <= 2
        ? String(fields.year % 100).padStart(2, '0')
        : String(fields.year).padStart(FULL_YEAR_DIGITS, '0');
    case 'month': {
      if (token.width < NAME_TOKEN_WIDTH) return twoDigits(fields.month, token.width);
      if (section.japanese) return token.width > MONTH_FULL_WIDTH ? String(fields.month) : `${fields.month}${JA_MONTH_SUFFIX}`;
      const name = MONTH_NAMES[fields.month - 1] ?? '';
      if (token.width === NAME_TOKEN_WIDTH) return name.slice(0, NAME_TOKEN_WIDTH);
      return token.width === MONTH_FULL_WIDTH ? name : name.slice(0, 1);
    }
    case 'day': {
      if (token.width < NAME_TOKEN_WIDTH) return twoDigits(fields.day, token.width);
      if (section.japanese) {
        const short = JA_WEEKDAY_NAMES[fields.weekday] ?? '';
        return token.width === NAME_TOKEN_WIDTH ? short : short + JA_WEEKDAY_SUFFIX;
      }
      const name = WEEKDAY_NAMES[fields.weekday] ?? '';
      return token.width === NAME_TOKEN_WIDTH ? name.slice(0, NAME_TOKEN_WIDTH) : name;
    }
    case 'hour': {
      const hour = section.hasMeridiem ? fields.hour % HOURS_PER_HALF_DAY || HOURS_PER_HALF_DAY : fields.hour;
      return twoDigits(hour, token.width);
    }
    case 'minute':
      return twoDigits(fields.minute, token.width);
    case 'second':
      return twoDigits(fields.second, token.width);
    default:
      return '';
  }
};

const renderElapsed = (token: ElapsedToken, fields: DateFields): string => {
  const total =
    token.part === 'hour'
      ? Math.floor(fields.totalSeconds / SECONDS_PER_HOUR)
      : token.part === 'minute'
        ? Math.floor(fields.totalSeconds / SECONDS_PER_MINUTE)
        : fields.totalSeconds;
  return String(total).padStart(token.width, '0');
};

/**
 * A section can carry two fractional-second groups of different widths
 * (`[ss].000" "ss.0`). The serial is rounded once, at the widest, so a narrower
 * group rounds the digits that rounding already produced.
 */
const subsecondText = (subsecond: string, digits: number): string => {
  if (digits >= subsecond.length) return subsecond.padEnd(digits, '0');
  const kept = subsecond.slice(0, digits);
  if (subsecond.charCodeAt(digits) < FIVE_CHAR_CODE) return kept;
  const bumped = carryOne(kept);
  // Carrying out of the leading digit would have to bump the second as well,
  // which the wider group in the same section does not show.
  return bumped.length > digits ? '9'.repeat(digits) : bumped;
};

const renderDateTokens = (section: DateFormatSection, fields: DateFields): string => {
  const pieces: string[] = [];
  for (const token of section.tokens) {
    switch (token.kind) {
      case 'datePart':
        pieces.push(renderDatePart(token, fields, section));
        break;
      case 'elapsed':
        pieces.push(renderElapsed(token, fields));
        break;
      case 'meridiem':
        if (section.japanese && token.am.length > 1) pieces.push(fields.hour < HOURS_PER_HALF_DAY ? JA_AM : JA_PM);
        else pieces.push(fields.hour < HOURS_PER_HALF_DAY ? token.am : token.pm);
        break;
      case 'subsecond':
        pieces.push(`.${subsecondText(fields.subsecond, token.digits)}`);
        break;
      default:
        pieces.push(literalFor(token));
    }
  }
  return pieces.join('');
};

const renderDateSection = (section: DateFormatSection, serial: number, epoch: ExcelEpoch): string | undefined => {
  const subDigits = Math.min(section.subsecondDigits, MAX_SUBSECOND_DIGITS);
  const subUnits = 10 ** subDigits;
  const unitsPerDay = SECONDS_PER_DAY * subUnits;
  const units = Math.round(serial * unitsPerDay);
  if (!Number.isSafeInteger(units)) return undefined;
  const totalSeconds = Math.floor(units / subUnits);
  const subsecond = String(units % subUnits).padStart(subDigits, '0');

  if (section.elapsed) {
    return renderDateTokens(section, {
      year: 0,
      month: 1,
      day: 1,
      weekday: 0,
      hour: Math.floor(totalSeconds / SECONDS_PER_HOUR) % HOURS_PER_DAY,
      minute: Math.floor(totalSeconds / SECONDS_PER_MINUTE) % MINUTES_PER_HOUR,
      second: totalSeconds % SECONDS_PER_MINUTE,
      totalSeconds,
      subsecond,
    });
  }

  const date = excelToDate(units / unitsPerDay, { epoch });
  // A serial can be a safe integer of sub-second units and still land outside
  // the range a `Date` covers, which would print NaN into every field.
  if (Number.isNaN(date.getTime())) return undefined;
  // Excel's 1900 system shows two days no calendar has: serial 0 as 1/0/1900
  // and serial 60 as the phantom 2/29/1900. `excelToDate` folds both onto real
  // days, so the displayed fields are put back here.
  const wholeDay = Math.floor(units / unitsPerDay);
  const phantom = epoch === 'mac' ? undefined : PHANTOM_DAYS.get(wholeDay);
  return renderDateTokens(section, {
    year: phantom ? PHANTOM_YEAR : date.getUTCFullYear(),
    month: phantom ? phantom.month : date.getUTCMonth() + 1,
    day: phantom ? phantom.day : date.getUTCDate(),
    weekday: phantom ? phantom.weekday : date.getUTCDay(),
    hour: date.getUTCHours(),
    minute: date.getUTCMinutes(),
    second: date.getUTCSeconds(),
    totalSeconds,
    subsecond,
  });
};

// ---- rendering: entry points ----------------------------------------------

interface PickedSection {
  readonly section: FormatSection;
  /** True when the caller has to supply the minus sign itself. */
  readonly signed: boolean;
}

/**
 * Excel picks the section by sign: one section covers everything, two split
 * at zero, three or four add a dedicated zero branch. Only a lone section has
 * to be handed the sign, because the others spell their own.
 */
const pickSection = (format: ParsedFormat, value: number): PickedSection | undefined => {
  if (format.conditions) return pickConditional(format.sections, format.conditions, value);
  const { sections } = format;
  const first = sections[0];
  if (first === undefined) return undefined;
  if (sections.length === 1) return { section: first, signed: value < 0 };
  if (value < 0) {
    const negative = sections[1];
    return negative === undefined ? undefined : { section: negative, signed: false };
  }
  if (value === 0 && sections.length > 2) {
    const zero = sections[2];
    return zero === undefined ? undefined : { section: zero, signed: false };
  }
  return { section: first, signed: false };
};

const holds = (c: SectionCondition, value: number): boolean => {
  switch (c.op) {
    case '<':
      return value < c.operand;
    case '<=':
      return value <= c.operand;
    case '>':
      return value > c.operand;
    case '>=':
      return value >= c.operand;
    case '=':
      return value === c.operand;
    case '<>':
      return value !== c.operand;
  }
};

/** A comparison that only negative values can meet: such a section spells its own sign. */
const negativeOnly = (c: SectionCondition): boolean =>
  (c.op === '<' && c.operand <= 0) || (c.op === '<=' && c.operand < 0) || (c.op === '=' && c.operand < 0);

/**
 * Comparison sections, as Excel for Mac reads them: the first two sections are
 * tried in order (a second section without its own comparison means `[<0]`
 * when a third follows, and "everything else" when it is the last), then the
 * third catches the rest. A value no section takes has no reading (`####`).
 */
const pickConditional = (
  sections: readonly FormatSection[],
  conditions: readonly (SectionCondition | undefined)[],
  value: number,
): PickedSection | undefined => {
  const numeric = sections.length === MAX_SECTIONS ? sections.slice(0, -1) : sections;
  for (let i = 0; i < numeric.length; i++) {
    const section = numeric[i];
    if (section === undefined) continue;
    const isLast = i === numeric.length - 1;
    const condition = conditions[i] ?? (isLast ? undefined : i === 0 ? { op: '>=', operand: 0 } : { op: '<', operand: 0 });
    if (condition === undefined || holds(condition, value)) {
      return { section, signed: value < 0 && !(condition !== undefined && negativeOnly(condition)) };
    }
    // Only the first two sections carry comparisons; a third is the catch-all.
    if (i >= 1 && isLast) return undefined;
  }
  return undefined;
};

const renderSection = (section: FormatSection, magnitude: number, epoch: ExcelEpoch): RenderedSection | undefined => {
  switch (section.kind) {
    case 'literal':
      return { text: section.text, zero: true };
    case 'blank':
      return { text: '', zero: true };
    // A numeric value in a text-formatted cell falls back to General.
    case 'general':
    case 'text':
      return { text: renderGeneral(magnitude), zero: magnitude === 0 };
    case 'number':
      return renderNumberSection(section, magnitude);
    case 'fraction':
      return renderFractionSection(section, magnitude);
    case 'date': {
      const text = renderDateSection(section, magnitude, epoch);
      return text === undefined ? undefined : { text, zero: false };
    }
  }
};

/**
 * Render a numeric cell value (a plain number, or a date / duration already
 * converted to its Excel serial). `undefined` means the value has no reading
 * under this format and the caller should degrade.
 */
export function renderNumericValue(format: ParsedFormat, value: number, epoch: ExcelEpoch): string | undefined {
  if (!Number.isFinite(value)) return undefined;
  const picked = pickSection(format, value);
  if (picked === undefined) return undefined;
  // A negative serial is not a calendar date. Excel fills the cell with `#`
  // characters, and how many depends on the column width, so there is nothing
  // to print. An elapsed span is different: it has a sign, and `picked.signed`
  // puts it in front.
  if (picked.section.kind === 'date' && !picked.section.elapsed && value < 0) return undefined;
  const rendered = renderSection(picked.section, Math.abs(value), epoch);
  if (rendered === undefined) return undefined;
  if (!picked.signed || rendered.zero) return rendered.text;
  return `-${rendered.text}`;
}

/**
 * Render a string cell value. Only the fourth section of a code applies to
 * text, plus a lone section that is itself a text layout (`@" pcs"`); under any
 * other code Excel prints the string unchanged.
 */
export function renderTextValue(format: ParsedFormat, text: string): string {
  const section =
    format.sections.length === MAX_SECTIONS
      ? format.sections[MAX_SECTIONS - 1]
      : format.sections.length === 1
        ? format.sections[0]
        : undefined;
  if (section === undefined) return text;
  if (section.kind === 'blank') return '';
  if (section.kind === 'literal' && format.sections.length === MAX_SECTIONS) return section.text;
  if (section.kind !== 'text') return text;
  return section.tokens.map((token) => (token.kind === 'textPlaceholder' ? text : literalFor(token))).join('');
}
