// Formula AutoComplete and the argument ScreenTip shown while typing a
// formula. Both scan the text left of the caret once; string literals and
// array constants are skipped so their commas and letters don't count.

import { FUNCTION_CATALOG, type FunctionInfo } from '../calc/index.ts';

const MAX_ITEMS = 12;
const IDENT_TAIL = /[A-Za-z_\\][A-Za-z0-9_.]*$/;
/** Characters that make a trailing word part of something else (a reference, sheet name, number). */
const NOT_A_START = /[A-Za-z0-9_.$:!'\]]/;

export interface CompletionItem {
  readonly name: string;
  readonly kind: 'function' | 'name';
  readonly detail: string;
}

export interface Completion {
  /** Span of the partial word the chosen item replaces. */
  readonly start: number;
  readonly end: number;
  readonly items: readonly CompletionItem[];
}

/** True when `caret` sits outside any string literal of formula `text`. */
function outsideString(text: string, caret: number): boolean {
  let inString = false;
  for (let i = 0; i < caret; i++) if (text[i] === '"') inString = !inString;
  return !inString;
}

export function completionAt(text: string, caret: number, names: readonly string[]): Completion | undefined {
  if (!text.startsWith('=') || caret < 2 || !outsideString(text, caret)) return undefined;
  const before = text.slice(0, caret);
  const word = IDENT_TAIL.exec(before)?.[0];
  if (!word) return undefined;
  const start = caret - word.length;
  const prev = text[start - 1] ?? '';
  if (start > 1 && NOT_A_START.test(prev)) return undefined;
  const upper = word.toUpperCase();
  const items: CompletionItem[] = [];
  for (const f of FUNCTION_CATALOG) {
    if (f.name.startsWith(upper)) items.push({ name: f.name, kind: 'function', detail: f.description });
    if (items.length >= MAX_ITEMS) break;
  }
  for (const n of names) {
    if (items.length >= MAX_ITEMS) break;
    if (n.toUpperCase().startsWith(upper)) items.push({ name: n, kind: 'name', detail: '' });
  }
  if (items.length === 0) return undefined;
  // Excel hides the list once the word already is the only match, spelled out.
  if (items.length === 1 && items[0]?.name.toUpperCase() === upper) return undefined;
  return { start, end: caret, items };
}

/** Text and caret after accepting `item`: functions get their opening parenthesis. */
export function applyCompletion(text: string, c: Completion, item: CompletionItem): { text: string; caret: number } {
  const insert = item.kind === 'function' ? `${item.name}(` : item.name;
  const rest = text.slice(c.end);
  // Don't double the parenthesis when the user already typed one.
  const ins = item.kind === 'function' && rest.startsWith('(') ? item.name : insert;
  return { text: text.slice(0, c.start) + ins + rest, caret: c.start + insert.length };
}

export interface ArgumentHint {
  readonly fn: FunctionInfo;
  /** Zero-based argument the caret is in. */
  readonly argIndex: number;
}

const CATALOG_BY_NAME = new Map(FUNCTION_CATALOG.map((f) => [f.name, f]));

/** The innermost known function call whose argument list contains the caret. */
export function argumentHintAt(text: string, caret: number): ArgumentHint | undefined {
  if (!text.startsWith('=')) return undefined;
  const stack: Array<{ name: string; arg: number }> = [];
  let inString = false;
  let braces = 0;
  for (let i = 1; i < caret; i++) {
    const ch = text[i];
    if (ch === '"') inString = !inString;
    if (inString) continue;
    if (ch === '{') braces++;
    else if (ch === '}') braces = Math.max(0, braces - 1);
    else if (ch === '(') {
      const name = IDENT_TAIL.exec(text.slice(0, i))?.[0] ?? '';
      stack.push({ name: name.toUpperCase().replace(/^_XLFN\./, ''), arg: 0 });
    } else if (ch === ')') stack.pop();
    else if (ch === ',' && braces === 0) {
      const top = stack[stack.length - 1];
      if (top) top.arg++;
    }
  }
  if (inString) return undefined;
  for (let i = stack.length - 1; i >= 0; i--) {
    const frame = stack[i];
    const fn = frame && CATALOG_BY_NAME.get(frame.name);
    if (frame && fn) return { fn, argIndex: frame.arg };
  }
  return undefined;
}

/**
 * Split a syntax string like `SUM(number1, [number2], ...)` into the prefix,
 * its arguments and the closing part, so the current argument can be bolded.
 * A trailing `...` keeps the last named argument current for later positions.
 */
export function syntaxParts(syntax: string): { head: string; args: string[] } {
  const open = syntax.indexOf('(');
  if (open < 0 || !syntax.endsWith(')')) return { head: syntax, args: [] };
  const inner = syntax.slice(open + 1, -1).trim();
  return { head: syntax.slice(0, open + 1), args: inner ? inner.split(/,\s*/) : [] };
}

export function currentArgIndex(args: readonly string[], argIndex: number): number {
  if (argIndex < args.length && args[argIndex] !== '...') return argIndex;
  const last = args[args.length - 1] === '...' ? args.length - 2 : args.length - 1;
  return args[args.length - 1] === '...' ? last : -1;
}

/**
 * Ctrl+Shift+A: with the caret right after a function name (or its opening
 * parenthesis), spell out the argument names, e.g. `=ROUND(` → `=ROUND(number, num_digits)`.
 */
export function insertArgumentNames(text: string, caret: number): { text: string; caret: number } | undefined {
  if (!text.startsWith('=')) return undefined;
  const m = /([A-Za-z_][A-Za-z0-9_.]*)(\(?)$/.exec(text.slice(0, caret));
  const fn = m?.[1] && CATALOG_BY_NAME.get(m[1].toUpperCase());
  if (!m || !fn) return undefined;
  const { args } = syntaxParts(fn.syntax);
  const insert = `${m[2] ? '' : '('}${args.join(', ')})`;
  return { text: text.slice(0, caret) + insert + text.slice(caret), caret: caret + insert.length };
}

const CELL_PART = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;

/**
 * Re-spell a cell/area reference for a new rectangle, keeping its sheet
 * prefix and `$` anchoring (the range finder's drag). Whole-row / whole-column
 * references aren't draggable in Excel either, so they return undefined.
 */
export function rewriteReference(refText: string, range: { r1: number; c1: number; r2: number; c2: number }): string | undefined {
  const bang = refText.lastIndexOf('!');
  const prefix = refText.slice(0, bang + 1);
  const parts = refText.slice(bang + 1).split(':');
  const first = CELL_PART.exec(parts[0] ?? '');
  const second = parts[1] === undefined ? first : CELL_PART.exec(parts[1]);
  if (!first || !second || parts.length > 2) return undefined;
  const spell = (m: RegExpExecArray, row: number, col: number) => `${m[1]}${columnLetters(col)}${m[3]}${row}`;
  const single = range.r1 === range.r2 && range.c1 === range.c2;
  if (single && parts.length === 1) return prefix + spell(first, range.r1, range.c1);
  return `${prefix}${spell(first, range.r1, range.c1)}:${spell(second, range.r2, range.c2)}`;
}

function columnLetters(col: number): string {
  let s = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
