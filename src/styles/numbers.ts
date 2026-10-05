// Number-format value object + the OOXML built-in format catalogue.
// Mirrors openpyxl/openpyxl/styles/numbers.py.
//
// Excel keeps two parallel namespaces for number formats:
//   * IDs 0–163 are reserved for the built-in catalogue (sparsely
//     populated; only 36 of them are actually defined).
//   * IDs ≥ 164 are user-defined / locale-specific. The Stylesheet
//     pool allocates these on demand (phase 2 §3.4).

import { OpenXmlSchemaError } from '../utils/exceptions.js';

/** Canonical OOXML built-in number formats — verbatim from openpyxl. */
export const BUILTIN_FORMATS: Readonly<Record<number, string>> = Object.freeze({
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0_);(#,##0)',
  38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
});

/** First numFmtId Excel reserves for user-defined formats. */
export const BUILTIN_FORMATS_MAX_SIZE = 164;

const REVERSE: ReadonlyMap<string, number> = new Map(Object.entries(BUILTIN_FORMATS).map(([k, v]) => [v, Number(k)]));

// ---- well-known format codes (named exports for ergonomics) ---------------
//
// Inlined rather than indexed off BUILTIN_FORMATS to keep TS' tuple
// lookup from widening to `string | undefined`.

export const FORMAT_GENERAL = 'General';
export const FORMAT_TEXT = '@';
export const FORMAT_NUMBER = '0';
export const FORMAT_NUMBER_00 = '0.00';
export const FORMAT_PERCENTAGE = '0%';
export const FORMAT_PERCENTAGE_00 = '0.00%';
export const FORMAT_DATE_DATETIME = 'yyyy-mm-dd h:mm:ss';
export const FORMAT_DATE_TIMEDELTA = '[hh]:mm:ss';
export const FORMAT_DATE_YYYYMMDD2 = 'yyyy-mm-dd';

// ---- helpers ---------------------------------------------------------------

/** Look up the format code for a given numFmtId, or `undefined` if unknown. */
export function builtinFormatCode(id: number): string | undefined {
  return Object.hasOwn(BUILTIN_FORMATS, id) ? BUILTIN_FORMATS[id] : undefined;
}

/** Look up the numFmtId for a given format code, or `undefined` if not built-in. */
export function builtinFormatId(code: string): number | undefined {
  return REVERSE.get(code);
}

/** True iff `code` is one of the OOXML built-in format strings. */
export function isBuiltinFormat(code: string): boolean {
  return REVERSE.has(code);
}

// ---- date / timedelta heuristics ------------------------------------------
//
// Mirror openpyxl's regex strategy verbatim. The two patterns work together
// to decide whether a format implies a date / time / duration interpretation.

/**
 * The colour names Excel accepts in a bracket group (`[Red]`), as the regex
 * alternation `STRIP_RE` needs them. The format-code parser needs the same
 * names to tell a colour apart from a comparison or a calendar modifier, so
 * they live here once.
 */
export const FORMAT_COLOR_NAMES = 'BLACK|BLUE|CYAN|GREEN|MAGENTA|RED|WHITE|YELLOW';

const COLORS_GROUP = `\\[(${FORMAT_COLOR_NAMES})\\]`;
const LITERAL_GROUP = '"[^"]*"';
const LOCALE_GROUP = '\\[(?!hh?\\]|mm?\\]|ss?\\])[^\\]]*\\]';

const STRIP_RE = new RegExp(`${COLORS_GROUP}|${LITERAL_GROUP}|${LOCALE_GROUP}`, 'g');
// `g` is the Japanese era and `aaa` a weekday name; the `g` of `General` is not.
const DATE_TOKEN_RE = /(?<![_\\])(?:[dmhysDMHYS]|g(?!eneral)|[aA]{3})/;
const TIMEDELTA_RE = /\[hh?\](:mm(:ss(\.0*)?)?)?|\[mm?\](:ss(\.0*)?)?|\[ss?\](\.0*)?/i;

/**
 * Heuristic: does the format string imply a date / time interpretation?
 * Looks at only the first format section (positive-value branch); strips
 * literals, colour codes and locale modifiers before scanning for date
 * tokens (d/m/h/y/s, case-insensitive) that aren't escaped.
 */
export function isDateFormat(code: string | undefined | null): boolean {
  if (code == null) return false;
  const head = code.split(';')[0] ?? '';
  const stripped = head.replace(STRIP_RE, '');
  return DATE_TOKEN_RE.test(stripped);
}

/** Heuristic: does the format string indicate a duration ([h]:mm:ss etc.)? */
export function isTimedeltaFormat(code: string | undefined | null): boolean {
  if (code == null) return false;
  const head = code.split(';')[0] ?? '';
  return TIMEDELTA_RE.test(head);
}

/** Categorise a date format as 'date', 'time', 'datetime' or undefined. */
export function classifyDateFormat(code: string | undefined | null): 'date' | 'time' | 'datetime' | undefined {
  if (!isDateFormat(code)) return undefined;
  // Reach here only with a non-null code.
  const head = (code as string).split(';')[0] ?? '';
  // Same locale / literal stripping the date detector applies.
  const stripped = head.replace(STRIP_RE, '');
  // `aaa` is a weekday name; a lone `a` belongs to AM/PM.
  let date = /[aA]{3}/.test(stripped);
  let time = false;
  for (const ch of stripped) {
    if (ch === 'd' || ch === 'D' || ch === 'y' || ch === 'Y' || ch === 'g') date = true;
    else if (ch === 'h' || ch === 'H' || ch === 's' || ch === 'S') time = true;
    if (date && time) break;
  }
  if (date && time) return 'datetime';
  if (date) return 'date';
  return 'time';
}

// ---- NumberFormat value -----------------------------------------------------

export interface NumberFormat {
  /** Stylesheet-relative numFmtId. */
  readonly numFmtId: number;
  readonly formatCode: string;
}

export function makeNumberFormat(opts: { numFmtId: number; formatCode: string }): NumberFormat {
  if (!Number.isInteger(opts.numFmtId) || opts.numFmtId < 0) {
    throw new OpenXmlSchemaError(`NumberFormat numFmtId must be a non-negative integer; got ${opts.numFmtId}`);
  }
  if (typeof opts.formatCode !== 'string') {
    throw new OpenXmlSchemaError('NumberFormat formatCode must be a string');
  }
  return Object.freeze({ numFmtId: opts.numFmtId, formatCode: opts.formatCode });
}
