// Text functions. Excel's *B variants (LENB, MIDB…) count bytes only in
// double-byte locales; this engine follows the single-byte behaviour, where
// they equal their plain counterparts.

import { makeCell } from '@office-kit/xlsx/cell';
import { getCellDisplayText, setCellNumberFormat } from '@office-kit/xlsx/styles';
import { createWorkbook } from '@office-kit/xlsx/workbook';
import { parseNumericText, toText } from '../coerce.ts';
import type { FnContext, FunctionSpec } from '../function-spec.ts';
import { type CalcScalar, type CalcValue, ERRORS, isError, makeArray } from '../types.ts';
import { hasWildcard, wildcardRegex } from './criteria.ts';
import { roundTo } from './math.ts';
import { bool, forEachValue, num, nums, optBool, scalar, str, trunc } from './helpers.ts';

const C = 'Text' as const;
const MAX_TEXT = 32_767;

// TEXT / DOLLAR render through the library's number-format engine — the same
// one the grid displays cells with — via a private scratch workbook.
const scratch = createWorkbook();

/** Render a value through an Excel number format code. */
export function formatValue(value: number | string | boolean, code: string, date1904: boolean): string {
  scratch.date1904 = date1904;
  const cell = makeCell(1, 1, value);
  setCellNumberFormat(scratch, cell, code);
  return getCellDisplayText(scratch, cell);
}

const textFn = (name: string, description: string, syntax: string, fn: (s: string) => CalcValue): FunctionSpec => ({
  name,
  category: C,
  syntax,
  description,
  minArgs: 1,
  maxArgs: 1,
  impl: (args) => {
    const s = str(args[0]);
    return isError(s) ? s : fn(s);
  },
});

const leftRight = (name: string, fromLeft: boolean): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(text, [num_chars])`,
  description: fromLeft ? 'Returns the specified number of characters from the start of a text string.' : 'Returns the specified number of characters from the end of a text string.',
  minArgs: 1,
  maxArgs: 2,
  impl: (args, ctx) => {
    const s = str(args[0]);
    if (isError(s)) return s;
    const n = args[1] === undefined ? 1 : num(args[1], ctx);
    if (isError(n)) return n;
    const k = trunc(n);
    if (k < 0) return ERRORS.VALUE;
    return fromLeft ? s.slice(0, k) : k === 0 ? '' : s.slice(-k);
  },
});

const midFn = (name: string): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(text, start_num, num_chars)`,
  description: 'Returns a specific number of characters from a text string, starting at the position you specify.',
  minArgs: 3,
  maxArgs: 3,
  impl: (args, ctx) => {
    const s = str(args[0]);
    if (isError(s)) return s;
    const v = nums(args.slice(1), ctx, [undefined, undefined]);
    if (isError(v)) return v;
    const start = trunc(v[0] ?? 0);
    const len = trunc(v[1] ?? 0);
    if (start < 1 || len < 0) return ERRORS.VALUE;
    return s.slice(start - 1, start - 1 + len);
  },
});

const findFn = (name: string, insensitive: boolean): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(find_text, within_text, [start_num])`,
  description: insensitive
    ? 'Returns the position of one text string inside another, ignoring case and allowing wildcards.'
    : 'Returns the starting position of one text string within another (case-sensitive).',
  minArgs: 2,
  maxArgs: 3,
  impl: (args, ctx) => {
    const needle = str(args[0]);
    if (isError(needle)) return needle;
    const hay = str(args[1]);
    if (isError(hay)) return hay;
    const s = args[2] === undefined ? 1 : num(args[2], ctx);
    if (isError(s)) return s;
    const start = trunc(s);
    if (start < 1 || start > hay.length + 1) return ERRORS.VALUE;
    if (needle === '') return start;
    if (!insensitive) {
      const i = hay.indexOf(needle, start - 1);
      return i < 0 ? ERRORS.VALUE : i + 1;
    }
    if (hasWildcard(needle)) {
      // Unanchored search: try each start position against the anchored pattern's prefix.
      const re = new RegExp(wildcardRegex(needle).source.slice(1, -1), 'is');
      const m = re.exec(hay.slice(start - 1));
      return m === null ? ERRORS.VALUE : m.index + start;
    }
    const i = hay.toLowerCase().indexOf(needle.toLowerCase(), start - 1);
    return i < 0 ? ERRORS.VALUE : i + 1;
  },
});

// Windows-1252 code points 128–159; Excel's CHAR / CODE use the ANSI code page.
const CP1252 = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0x8d, 0x17d, 0x8f, 0x90, 0x2018,
  0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];

const charFromCode = (n: number): string => String.fromCharCode(n >= 128 && n < 160 ? (CP1252[n - 128] ?? n) : n);

const codeFromChar = (ch: number): number => {
  const idx = CP1252.indexOf(ch);
  if (idx >= 0) return 128 + idx;
  // Characters outside the code page come back as '?', as Excel reports them.
  return ch > 255 ? 63 : ch;
};

const groupThousands = (digits: string): string => digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

const fixed = (n: number, decimals: number, commas: boolean): string => {
  const r = roundTo(n, decimals, 'half');
  const body = Math.abs(r).toFixed(Math.max(0, decimals));
  const [int = '', frac] = body.split('.');
  const grouped = commas ? groupThousands(int) : int;
  return `${r < 0 ? '-' : ''}${grouped}${frac !== undefined ? `.${frac}` : ''}`;
};

/** Split `text` at every delimiter (any of several), honouring case-insensitive matching. */
const splitAll = (text: string, delimiters: readonly string[], insensitive: boolean): string[] => {
  const usable = delimiters.filter((d) => d !== '');
  if (usable.length === 0) return [text];
  const hay = insensitive ? text.toLowerCase() : text;
  const needles = insensitive ? usable.map((d) => d.toLowerCase()) : usable;
  const parts: string[] = [];
  let from = 0;
  for (let i = 0; i < text.length; ) {
    const hit = needles.find((d) => hay.startsWith(d, i));
    if (hit === undefined) {
      i++;
      continue;
    }
    parts.push(text.slice(from, i));
    i += hit.length;
    from = i;
  }
  parts.push(text.slice(from));
  return parts;
};

const delimiterList = (v: CalcValue | undefined, ctx: FnContext): string[] | typeof ERRORS.VALUE => {
  if (v === undefined || v === null) return [];
  const out: string[] = [];
  for (const x of ctx.toArray(v).data) {
    const t = toText(x);
    if (isError(t)) return ERRORS.VALUE;
    out.push(t);
  }
  return out;
};

/** TEXTBEFORE / TEXTAFTER. */
const beforeAfter = (after: boolean) => (args: CalcValue[], ctx: FnContext): CalcValue => {
  const text = str(args[0]);
  if (isError(text)) return text;
  const delims = delimiterList(args[1], ctx);
  if (isError(delims)) return delims;
  // An empty argument, TEXTAFTER(t, d,,,, "none"), takes the default here.
  const v = nums(args.slice(2, 5).map((a) => a ?? undefined), ctx, [1, 0, 0]);
  if (isError(v)) return v;
  const instance = trunc(v[0] ?? 1);
  const insensitive = v[1] === 1;
  const matchEnd = v[2] === 1;
  if (instance === 0 || Math.abs(instance) > text.length + 1) return ERRORS.VALUE;
  const notFound = args[5] === undefined ? ERRORS.NA : scalar(args[5]);
  const hay = insensitive ? text.toLowerCase() : text;
  const needles = (insensitive ? delims.map((d) => d.toLowerCase()) : delims);
  // Every delimiter occurrence as [start, end).
  const hits: Array<[number, number]> = [];
  for (let i = 0; i < text.length; ) {
    const d = needles.find((n) => n !== '' && hay.startsWith(n, i));
    if (d === undefined) {
      i++;
      continue;
    }
    hits.push([i, i + d.length]);
    i += d.length;
  }
  if (needles.some((n) => n === '')) {
    // An empty delimiter matches at the very start (or end, counting backwards).
    return instance > 0 ? (after ? text : '') : after ? '' : text;
  }
  if (matchEnd) hits.push([text.length, text.length]);
  const idx = instance > 0 ? instance - 1 : hits.length + instance;
  const hit = hits[idx];
  if (hit === undefined) return notFound;
  return after ? text.slice(hit[1]) : text.slice(0, hit[0]);
};

export const TEXT_FUNCTIONS: FunctionSpec[] = [
  leftRight('LEFT', true),
  leftRight('LEFTB', true),
  leftRight('RIGHT', false),
  leftRight('RIGHTB', false),
  midFn('MID'),
  midFn('MIDB'),
  textFn('LEN', 'Returns the number of characters in a text string.', 'LEN(text)', (s) => s.length),
  textFn('LENB', 'Returns the number of bytes used to represent the characters in a text string.', 'LENB(text)', (s) => s.length),
  findFn('FIND', false),
  findFn('FINDB', false),
  findFn('SEARCH', true),
  findFn('SEARCHB', true),
  {
    name: 'SUBSTITUTE',
    category: C,
    syntax: 'SUBSTITUTE(text, old_text, new_text, [instance_num])',
    description: 'Substitutes new text for old text in a text string.',
    minArgs: 3,
    maxArgs: 4,
    impl: (args, ctx) => {
      const [text, old, rep] = [str(args[0]), str(args[1]), str(args[2])];
      if (isError(text)) return text;
      if (isError(old)) return old;
      if (isError(rep)) return rep;
      if (old === '') return text;
      if (args[3] === undefined) return text.split(old).join(rep);
      const n = num(args[3], ctx);
      if (isError(n)) return n;
      const instance = trunc(n);
      if (instance < 1) return ERRORS.VALUE;
      let at = -1;
      for (let i = 0; i < instance; i++) {
        at = text.indexOf(old, at + 1);
        if (at < 0) return text;
      }
      return text.slice(0, at) + rep + text.slice(at + old.length);
    },
  },
  ...(['REPLACE', 'REPLACEB'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(old_text, start_num, num_chars, new_text)`,
      description: 'Replaces part of a text string with a different text string.',
      minArgs: 4,
      maxArgs: 4,
      impl: (args, ctx) => {
        const text = str(args[0]);
        if (isError(text)) return text;
        const v = nums(args.slice(1), ctx, [undefined, undefined]);
        if (isError(v)) return v;
        const rep = str(args[3]);
        if (isError(rep)) return rep;
        const start = trunc(v[0] ?? 0);
        const len = trunc(v[1] ?? 0);
        if (start < 1 || len < 0) return ERRORS.VALUE;
        return text.slice(0, start - 1) + rep + text.slice(start - 1 + len);
      },
    }),
  ),
  textFn('UPPER', 'Converts text to uppercase.', 'UPPER(text)', (s) => s.toUpperCase()),
  textFn('LOWER', 'Converts text to lowercase.', 'LOWER(text)', (s) => s.toLowerCase()),
  textFn('PROPER', 'Capitalizes the first letter in each word of a text value.', 'PROPER(text)', (s) =>
    s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, pre: string, ch: string) => pre + ch.toUpperCase()),
  ),
  textFn('TRIM', 'Removes extra spaces from text, leaving single spaces between words.', 'TRIM(text)', (s) => s.replace(/ +/g, ' ').replace(/^ | $/g, '')),
  // eslint-disable-next-line no-control-regex
  textFn('CLEAN', 'Removes all nonprintable characters from text.', 'CLEAN(text)', (s) => s.replace(/[\u0000-\u001f]/g, '')),
  {
    name: 'CONCAT',
    category: C,
    syntax: 'CONCAT(text1, [text2], ...)',
    description: 'Combines the text from multiple ranges and/or strings.',
    minArgs: 1,
    maxArgs: 254,
    args: ['ref'],
    impl: (args, ctx) => {
      let out = '';
      let error: CalcScalar | undefined;
      forEachValue(args, ctx, (v) => {
        const t = toText(v);
        if (isError(t)) error = t;
        else out += t;
        return error === undefined;
      });
      if (error !== undefined) return error;
      return out.length > MAX_TEXT ? ERRORS.VALUE : out;
    },
  },
  {
    name: 'CONCATENATE',
    category: C,
    syntax: 'CONCATENATE(text1, [text2], ...)',
    description: 'Joins several text items into one text item.',
    minArgs: 1,
    maxArgs: 255,
    impl: (args) => {
      let out = '';
      for (const a of args) {
        const t = str(a);
        if (isError(t)) return t;
        out += t;
      }
      return out.length > MAX_TEXT ? ERRORS.VALUE : out;
    },
  },
  {
    name: 'TEXTJOIN',
    category: C,
    syntax: 'TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)',
    description: 'Combines the text from multiple ranges and/or strings, with a delimiter between each value.',
    minArgs: 3,
    maxArgs: 252,
    args: ['value', 'scalar', 'ref'],
    impl: (args, ctx) => {
      const delims = delimiterList(args[0], ctx);
      if (isError(delims)) return delims;
      const ignore = bool(args[1]);
      if (isError(ignore)) return ignore;
      const parts: string[] = [];
      let error: CalcScalar | undefined;
      forEachValue(args.slice(2), ctx, (v) => {
        const t = toText(v);
        if (isError(t)) error = t;
        else if (!(ignore && t === '')) parts.push(t);
        return error === undefined;
      });
      if (error !== undefined) return error;
      let out = parts[0] ?? '';
      for (let i = 1; i < parts.length; i++) out += (delims.length === 0 ? '' : (delims[(i - 1) % delims.length] ?? '')) + (parts[i] ?? '');
      return out.length > MAX_TEXT ? ERRORS.VALUE : out;
    },
  },
  {
    name: 'TEXT',
    category: C,
    syntax: 'TEXT(value, format_text)',
    description: 'Formats a number and converts it to text.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = scalar(args[0]);
      const code = str(args[1]);
      if (isError(code)) return code;
      if (isError(v)) return v;
      let value: number | string | boolean = v ?? 0;
      if (typeof value === 'string') value = parseNumericText(value, ctx.host.date1904) ?? value;
      return formatValue(value, code, ctx.host.date1904);
    },
  },
  {
    name: 'VALUE',
    category: C,
    syntax: 'VALUE(text)',
    description: 'Converts a text string that represents a number to a number.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const v = scalar(args[0]);
      if (typeof v === 'number') return v;
      if (typeof v === 'boolean') return ERRORS.VALUE;
      if (v === null) return 0;
      if (isError(v)) return v;
      return parseNumericText(v, ctx.host.date1904) ?? ERRORS.VALUE;
    },
  },
  {
    name: 'NUMBERVALUE',
    category: C,
    syntax: 'NUMBERVALUE(text, [decimal_separator], [group_separator])',
    description: 'Converts text to a number, in a locale-independent way.',
    minArgs: 1,
    maxArgs: 3,
    impl: (args) => {
      const text = str(args[0]);
      if (isError(text)) return text;
      const dec = args[1] === undefined ? '.' : str(args[1]);
      if (isError(dec)) return dec;
      const grp = args[2] === undefined ? ',' : str(args[2]);
      if (isError(grp)) return grp;
      const d = dec.charAt(0);
      const g = grp.charAt(0);
      if (d === '' || d === g) return ERRORS.VALUE;
      let s = text.replace(/\s+/g, '');
      if (s === '') return 0;
      let pct = 0;
      while (s.endsWith('%')) {
        pct++;
        s = s.slice(0, -1);
      }
      const decAt = s.indexOf(d);
      if (decAt >= 0 && s.indexOf(d, decAt + 1) >= 0) return ERRORS.VALUE;
      const intPart = (decAt >= 0 ? s.slice(0, decAt) : s).split(g).join('');
      const fracPart = decAt >= 0 ? s.slice(decAt + 1) : '';
      if (fracPart.includes(g)) return ERRORS.VALUE;
      const n = Number(`${intPart === '' || intPart === '-' || intPart === '+' ? `${intPart}0` : intPart}.${fracPart || '0'}`);
      return Number.isFinite(n) ? n / 100 ** pct : ERRORS.VALUE;
    },
  },
  {
    name: 'REPT',
    category: C,
    syntax: 'REPT(text, number_times)',
    description: 'Repeats text a given number of times.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const s = str(args[0]);
      if (isError(s)) return s;
      const n = num(args[1], ctx);
      if (isError(n)) return n;
      const k = trunc(n);
      if (k < 0 || s.length * k > MAX_TEXT) return ERRORS.VALUE;
      return s.repeat(k);
    },
  },
  {
    name: 'EXACT',
    category: C,
    syntax: 'EXACT(text1, text2)',
    description: 'Checks whether two text strings are exactly the same (case-sensitive).',
    minArgs: 2,
    maxArgs: 2,
    impl: (args) => {
      const a = str(args[0]);
      if (isError(a)) return a;
      const b = str(args[1]);
      return isError(b) ? b : a === b;
    },
  },
  {
    name: 'CHAR',
    category: C,
    syntax: 'CHAR(number)',
    description: 'Returns the character specified by the code number.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const n = num(args[0], ctx);
      if (isError(n)) return n;
      const k = trunc(n);
      return k < 1 || k > 255 ? ERRORS.VALUE : charFromCode(k);
    },
  },
  textFn('CODE', 'Returns a numeric code for the first character in a text string.', 'CODE(text)', (s) =>
    s === '' ? ERRORS.VALUE : codeFromChar(s.charCodeAt(0)),
  ),
  {
    name: 'UNICHAR',
    category: C,
    syntax: 'UNICHAR(number)',
    description: 'Returns the Unicode character that is referenced by the given numeric value.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const n = num(args[0], ctx);
      if (isError(n)) return n;
      const k = trunc(n);
      return k < 1 || k > 0x10ffff ? ERRORS.VALUE : String.fromCodePoint(k);
    },
  },
  textFn('UNICODE', 'Returns the number (code point) that corresponds to the first character of the text.', 'UNICODE(text)', (s) =>
    s === '' ? ERRORS.VALUE : (s.codePointAt(0) ?? ERRORS.VALUE),
  ),
  {
    name: 'T',
    category: C,
    syntax: 'T(value)',
    description: 'Returns the text referred to by value, or empty text for anything else.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args) => {
      const v = scalar(args[0]);
      return typeof v === 'string' ? v : '';
    },
  },
  {
    name: 'DOLLAR',
    category: C,
    syntax: 'DOLLAR(number, [decimals])',
    description: 'Converts a number to text, using currency format.',
    minArgs: 1,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, 2]);
      if (isError(v)) return v;
      const [n = 0, d = 2] = v;
      const decimals = trunc(d);
      const body = fixed(Math.abs(n), decimals, true);
      return roundTo(n, decimals, 'half') < 0 ? `($${body})` : `$${body}`;
    },
  },
  {
    name: 'FIXED',
    category: C,
    syntax: 'FIXED(number, [decimals], [no_commas])',
    description: 'Formats a number as text with a fixed number of decimals.',
    minArgs: 1,
    maxArgs: 3,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, 2]);
      if (isError(v)) return v;
      const noCommas = optBool(args[2], false);
      if (isError(noCommas)) return noCommas;
      return fixed(v[0] ?? 0, trunc(v[1] ?? 2), !noCommas);
    },
  },
  {
    name: 'TEXTBEFORE',
    category: C,
    syntax: 'TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])',
    description: 'Returns text that occurs before a given character or string.',
    minArgs: 2,
    maxArgs: 6,
    args: ['scalar', 'value', 'scalar', 'scalar', 'scalar', 'scalar'],
    impl: beforeAfter(false),
  },
  {
    name: 'TEXTAFTER',
    category: C,
    syntax: 'TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])',
    description: 'Returns text that occurs after a given character or string.',
    minArgs: 2,
    maxArgs: 6,
    args: ['scalar', 'value', 'scalar', 'scalar', 'scalar', 'scalar'],
    impl: beforeAfter(true),
  },
  {
    name: 'TEXTSPLIT',
    category: C,
    syntax: 'TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])',
    description: 'Splits text strings by using column and row delimiters.',
    minArgs: 2,
    maxArgs: 6,
    args: ['scalar', 'value', 'value', 'scalar', 'scalar', 'scalar'],
    impl: (args, ctx) => {
      const text = str(args[0]);
      if (isError(text)) return text;
      const colDelims = delimiterList(args[1], ctx);
      if (isError(colDelims)) return colDelims;
      const rowDelims = delimiterList(args[2], ctx);
      if (isError(rowDelims)) return rowDelims;
      const ignore = optBool(args[3], false);
      if (isError(ignore)) return ignore;
      const mode = args[4] === undefined ? 0 : num(args[4], ctx);
      if (isError(mode)) return mode;
      const pad = args[5] === undefined ? ERRORS.NA : scalar(args[5]);
      const insensitive = mode === 1;
      const drop = (parts: string[]): string[] => (ignore ? parts.filter((p) => p !== '') : parts);
      const rows = drop(splitAll(text, rowDelims, insensitive)).map((r) => drop(splitAll(r, colDelims, insensitive)));
      if (rows.length === 0) return ERRORS.CALC;
      const cols = Math.max(1, ...rows.map((r) => r.length));
      const data: CalcScalar[] = [];
      for (const r of rows) for (let c = 0; c < cols; c++) data.push(c < r.length ? (r[c] ?? '') : pad);
      return makeArray(rows.length, cols, data);
    },
  },
  {
    name: 'VALUETOTEXT',
    category: C,
    syntax: 'VALUETOTEXT(value, [format])',
    description: 'Returns text from any specified value.',
    minArgs: 1,
    maxArgs: 2,
    acceptsErrors: true,
    impl: (args, ctx) => {
      const v = scalar(args[0]);
      const strict = args[1] === undefined ? 0 : num(args[1], ctx);
      if (isError(strict)) return strict;
      if (isError(v)) return v.code;
      const t = toText(v);
      if (isError(t)) return t;
      return strict === 1 && typeof v === 'string' ? `"${t}"` : t;
    },
  },
];
