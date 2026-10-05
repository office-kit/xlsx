// Formula tokenizer. Every token remembers its [start, end) span in the source
// text so the text utilities (reference highlighting, F4, translate, structural
// adjustment, sheet rename) can splice the original string and leave the rest
// of the user's formula — spacing, casing, line breaks — untouched.
//
// References are recognised here rather than in the parser because A1 syntax
// is not context free at the character level: `A1:B2`, `A:C`, `1:3` and
// `Sheet1:Sheet3!A1` all reuse `:`, and `LOG10(` is a function while `LOG10`
// alone is a cell. Deciding on whole lexemes keeps the parser simple.

import type { RefArea, SheetPrefix } from './ast.ts';
import { columnFromLetters } from './address.ts';
import { type CalcError, errorFromCode, MAX_COL, MAX_ROW } from './types.ts';

export class CalcParseError extends Error {
  override readonly name = 'CalcParseError';
  /** 0-based offset into the formula text where the problem was found. */
  readonly position: number;

  constructor(message: string, position: number) {
    super(message);
    this.position = position;
  }
}

interface TokenBase {
  readonly start: number;
  readonly end: number;
  /** Whitespace separated this token from the previous one (the intersection operator). */
  readonly spaceBefore: boolean;
}

export type Token = TokenBase &
  (
    | { readonly kind: 'number'; readonly value: number }
    | { readonly kind: 'string'; readonly value: string }
    | { readonly kind: 'bool'; readonly value: boolean }
    | { readonly kind: 'error'; readonly error: CalcError }
    /** `area` is undefined for a qualified `Sheet1!#REF!`. `prefixEnd` is where the `!` ends. */
    | {
        readonly kind: 'ref';
        readonly prefix: SheetPrefix | undefined;
        readonly prefixEnd: number;
        readonly area: RefArea | undefined;
      }
    | {
        readonly kind: 'name';
        readonly name: string;
        readonly prefix: SheetPrefix | undefined;
        readonly prefixEnd: number;
      }
    | { readonly kind: 'func'; readonly name: string }
    /** `body` is the text between the outer brackets. */
    | { readonly kind: 'struct'; readonly table: string | undefined; readonly body: string }
    | { readonly kind: 'op'; readonly op: string }
    | { readonly kind: 'lparen' | 'rparen' | 'lbrace' | 'rbrace' | 'comma' | 'semicolon' }
  );

const WORD = /[A-Za-z_\\ -￿][\w.\\? -￿]*/uy;
const WORD_CHAR = /[\w.\\? -￿]/u;
const NUMBER = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
const CELL = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/y;
const COL_RANGE = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y;
const ROW_RANGE = /(\$?)(\d{1,7}):(\$?)(\d{1,7})/y;
const SHEET_3D_TAIL = /:([\w. -￿]+)!/uy;
const EXTERNAL_PREFIX = /\[([^\]]+)\]([\w. -￿]+)(?::([\w. -￿]+))?!/uy;
const ERROR_LITERAL = /#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA)/iy;

const PUNCTUATION = new Map<string, 'lparen' | 'rparen' | 'lbrace' | 'rbrace' | 'comma' | 'semicolon'>([
  ['(', 'lparen'],
  [')', 'rparen'],
  ['{', 'lbrace'],
  ['}', 'rbrace'],
  [',', 'comma'],
  [';', 'semicolon'],
]);

const OPERATORS = ['<=', '>=', '<>', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', ':', '@'] as const;

const sticky = (re: RegExp, text: string, at: number): RegExpExecArray | null => {
  re.lastIndex = at;
  return re.exec(text);
};

// A reference must not run on into a longer identifier (`A1B`, `A1.x`) or a
// call / sheet / table suffix.
const endsCleanly = (text: string, at: number): boolean => {
  const ch = text[at];
  return ch === undefined || !(WORD_CHAR.test(ch) || ch === '(' || ch === '!' || ch === '[' || ch === '$');
};

interface RefMatch {
  readonly area: RefArea;
  readonly end: number;
}

const cellAt = (text: string, at: number): { row: number; col: number; rAbs: boolean; cAbs: boolean; end: number } | undefined => {
  const m = sticky(CELL, text, at);
  if (m === null) return undefined;
  const col = columnFromLetters(m[2] ?? '');
  const row = Number(m[4]);
  if (col > MAX_COL || row < 1 || row > MAX_ROW) return undefined;
  return { row, col, rAbs: m[3] === '$', cAbs: m[1] === '$', end: at + m[0].length };
};

/** Recognise a cell, cell range, column range or row range starting exactly at `at`. */
export function matchReference(text: string, at: number): RefMatch | undefined {
  const first = cellAt(text, at);
  if (first !== undefined) {
    if (text[first.end] === ':') {
      const second = cellAt(text, first.end + 1);
      if (second !== undefined && endsCleanly(text, second.end)) {
        return {
          end: second.end,
          area: {
            kind: 'area',
            r1: first.row,
            c1: first.col,
            r2: second.row,
            c2: second.col,
            r1Abs: first.rAbs,
            c1Abs: first.cAbs,
            r2Abs: second.rAbs,
            c2Abs: second.cAbs,
          },
        };
      }
    }
    if (endsCleanly(text, first.end)) {
      return {
        end: first.end,
        area: {
          kind: 'cell',
          r1: first.row,
          c1: first.col,
          r2: first.row,
          c2: first.col,
          r1Abs: first.rAbs,
          c1Abs: first.cAbs,
          r2Abs: first.rAbs,
          c2Abs: first.cAbs,
        },
      };
    }
  }
  const cols = sticky(COL_RANGE, text, at);
  if (cols !== null && endsCleanly(text, at + cols[0].length)) {
    const c1 = columnFromLetters(cols[2] ?? '');
    const c2 = columnFromLetters(cols[4] ?? '');
    if (c1 <= MAX_COL && c2 <= MAX_COL) {
      return {
        end: at + cols[0].length,
        area: { kind: 'cols', r1: 1, c1, r2: MAX_ROW, c2, r1Abs: false, c1Abs: cols[1] === '$', r2Abs: false, c2Abs: cols[3] === '$' },
      };
    }
  }
  const rows = sticky(ROW_RANGE, text, at);
  if (rows !== null && endsCleanly(text, at + rows[0].length)) {
    const r1 = Number(rows[2]);
    const r2 = Number(rows[4]);
    if (r1 >= 1 && r1 <= MAX_ROW && r2 >= 1 && r2 <= MAX_ROW) {
      return {
        end: at + rows[0].length,
        area: { kind: 'rows', r1, c1: 1, r2, c2: MAX_COL, r1Abs: rows[1] === '$', c1Abs: false, r2Abs: rows[3] === '$', c2Abs: false },
      };
    }
  }
  return undefined;
}

/** Read `[...]` starting at `at` (which holds `[`), honouring nesting and `'` escapes. Returns the index after `]`. */
const readBrackets = (text: string, at: number): number => {
  let depth = 0;
  for (let i = at; i < text.length; i++) {
    const ch = text[i];
    if (ch === "'") {
      i++;
    } else if (ch === '[') {
      depth++;
    } else if (ch === ']') {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  throw new CalcParseError('Unterminated structured reference', at);
};

/** Split the inside of a quoted sheet prefix into external / sheet / sheet2. */
const prefixFromQuoted = (inner: string): SheetPrefix => {
  let rest = inner;
  let external: string | undefined;
  const ext = /^\[([^\]]+)\]/.exec(rest);
  if (ext !== null) {
    external = ext[1];
    rest = rest.slice(ext[0].length);
  }
  // `:` cannot occur in a sheet name, so it only ever separates a 3-D range.
  const colon = rest.indexOf(':');
  const sheet = colon < 0 ? rest : rest.slice(0, colon);
  const sheet2 = colon < 0 ? undefined : rest.slice(colon + 1);
  return {
    sheet,
    ...(sheet2 !== undefined ? { sheet2 } : {}),
    ...(external !== undefined ? { external } : {}),
  };
};

/**
 * Excel keeps only the first 15 significant digits of a typed number and
 * drops the rest rather than rounding: 1234567890123456 is 1234567890123450.
 */
function numberLiteral(text: string): number {
  const [mantissa = '', exponent] = text.toLowerCase().split('e');
  let seen = 0;
  let kept = '';
  for (const ch of mantissa) {
    if (ch === '.') kept += ch;
    else if (seen >= 15) kept += '0';
    else {
      if (seen > 0 || ch !== '0') seen++;
      kept += ch;
    }
  }
  return Number(exponent === undefined ? kept : `${kept}e${exponent}`);
}

/**
 * Split formula text into tokens. With `lenient`, a syntax error ends the
 * stream instead of throwing, so half-typed formulas (`=SUM(A1:B2` or an open
 * string) still yield the references typed so far.
 */
export function tokenize(text: string, lenient = false): Token[] {
  const tokens: Token[] = [];
  try {
    scan(text, tokens);
  } catch (e) {
    if (lenient && e instanceof CalcParseError) return tokens;
    throw e;
  }
  return tokens;
}

/** The token after a sheet prefix: a reference, a qualified name, or `#REF!`. */
function afterPrefix(text: string, prefix: SheetPrefix, at: number, start: number, spaceBefore: boolean): Token {
  const ref = matchReference(text, at);
  if (ref !== undefined) return { kind: 'ref', prefix, prefixEnd: at, area: ref.area, start, end: ref.end, spaceBefore };
  const err = sticky(ERROR_LITERAL, text, at);
  if (err !== null && err[0].toUpperCase() === '#REF!') {
    return { kind: 'ref', prefix, prefixEnd: at, area: undefined, start, end: at + err[0].length, spaceBefore };
  }
  const word = sticky(WORD, text, at);
  if (word !== null) return { kind: 'name', name: word[0], prefix, prefixEnd: at, start, end: at + word[0].length, spaceBefore };
  throw new CalcParseError('Expected a reference after the sheet name', at);
}

const operatorAt = (text: string, at: number): string | undefined => OPERATORS.find((o) => text.startsWith(o, at));

function scan(text: string, tokens: Token[]): void {
  const pushPrefixed = (prefix: SheetPrefix, at: number, start: number, spaceBefore: boolean): number => {
    const token = afterPrefix(text, prefix, at, start, spaceBefore);
    tokens.push(token);
    return token.end;
  };
  let i = 0;
  const n = text.length;

  while (i < n) {
    let spaceBefore = false;
    while (i < n && /\s/.test(text[i] ?? '')) {
      i++;
      spaceBefore = true;
    }
    if (i >= n) break;
    const start = i;
    const ch = text[i] ?? '';
    const prev = tokens[tokens.length - 1];

    if (ch === '"') {
      let value = '';
      let j = i + 1;
      for (;;) {
        if (j >= n) throw new CalcParseError('Unterminated string', start);
        const c = text[j];
        if (c === '"') {
          if (text[j + 1] === '"') {
            value += '"';
            j += 2;
            continue;
          }
          break;
        }
        value += c;
        j++;
      }
      tokens.push({ kind: 'string', value, start, end: j + 1, spaceBefore });
      i = j + 1;
      continue;
    }

    if (ch === "'") {
      let j = i + 1;
      let inner = '';
      for (;;) {
        if (j >= n) throw new CalcParseError('Unterminated sheet name', start);
        const c = text[j];
        if (c === "'") {
          if (text[j + 1] === "'") {
            inner += "'";
            j += 2;
            continue;
          }
          break;
        }
        inner += c;
        j++;
      }
      if (text[j + 1] !== '!') throw new CalcParseError("Expected '!' after a quoted sheet name", j + 1);
      i = pushPrefixed(prefixFromQuoted(inner), j + 2, start, spaceBefore);
      continue;
    }

    if (ch === '#') {
      // `A1#` is the spill-range operator; anywhere else `#` starts an error literal.
      if (!spaceBefore && prev !== undefined && (prev.kind === 'ref' || prev.kind === 'rparen' || prev.kind === 'name')) {
        tokens.push({ kind: 'op', op: '#', start, end: i + 1, spaceBefore });
        i++;
        continue;
      }
      const err = sticky(ERROR_LITERAL, text, i);
      if (err === null) throw new CalcParseError('Unknown error literal', i);
      const error = errorFromCode(err[0]);
      if (error === undefined) throw new CalcParseError('Unknown error literal', i);
      tokens.push({ kind: 'error', error, start, end: i + err[0].length, spaceBefore });
      i += err[0].length;
      continue;
    }

    if (ch === '[') {
      const ext = sticky(EXTERNAL_PREFIX, text, i);
      if (ext !== null) {
        const prefix: SheetPrefix = {
          external: ext[1] ?? '',
          sheet: ext[2] ?? '',
          ...(ext[3] !== undefined ? { sheet2: ext[3] } : {}),
        };
        i = pushPrefixed(prefix, i + ext[0].length, start, spaceBefore);
        continue;
      }
      const end = readBrackets(text, i);
      tokens.push({ kind: 'struct', table: undefined, body: text.slice(i + 1, end - 1), start, end, spaceBefore });
      i = end;
      continue;
    }

    if (ch === '$' || (ch >= '0' && ch <= '9') || ch === '.') {
      const ref = matchReference(text, i);
      if (ref !== undefined) {
        tokens.push({ kind: 'ref', prefix: undefined, prefixEnd: i, area: ref.area, start, end: ref.end, spaceBefore });
        i = ref.end;
        continue;
      }
      const num = sticky(NUMBER, text, i);
      if (num === null) throw new CalcParseError(`Unexpected character '${ch}'`, i);
      tokens.push({ kind: 'number', value: numberLiteral(num[0]), start, end: i + num[0].length, spaceBefore });
      i += num[0].length;
      continue;
    }

    const word = sticky(WORD, text, i);
    if (word !== null) {
      const w = word[0];
      const after = i + w.length;
      const next = text[after];
      if (next === '(') {
        tokens.push({ kind: 'func', name: w, start, end: after, spaceBefore });
        i = after;
        continue;
      }
      if (next === '!') {
        i = pushPrefixed({ sheet: w }, after + 1, start, spaceBefore);
        continue;
      }
      if (next === ':') {
        const tail = sticky(SHEET_3D_TAIL, text, after);
        if (tail !== null) {
          i = pushPrefixed({ sheet: w, sheet2: tail[1] ?? '' }, after + tail[0].length, start, spaceBefore);
          continue;
        }
      }
      const ref = matchReference(text, i);
      if (ref !== undefined) {
        tokens.push({ kind: 'ref', prefix: undefined, prefixEnd: i, area: ref.area, start, end: ref.end, spaceBefore });
        i = ref.end;
        continue;
      }
      if (next === '[') {
        const end = readBrackets(text, after);
        tokens.push({ kind: 'struct', table: w, body: text.slice(after + 1, end - 1), start, end, spaceBefore });
        i = end;
        continue;
      }
      const upper = w.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        tokens.push({ kind: 'bool', value: upper === 'TRUE', start, end: after, spaceBefore });
      } else {
        tokens.push({ kind: 'name', name: w, prefix: undefined, prefixEnd: start, start, end: after, spaceBefore });
      }
      i = after;
      continue;
    }

    const punct = PUNCTUATION.get(ch);
    if (punct !== undefined) {
      tokens.push({ kind: punct, start, end: i + 1, spaceBefore });
      i++;
      continue;
    }

    const op = operatorAt(text, i);
    if (op !== undefined) {
      tokens.push({ kind: 'op', op, start, end: i + op.length, spaceBefore });
      i += op.length;
      continue;
    }
    throw new CalcParseError(`Unexpected character '${ch}'`, i);
  }
}
