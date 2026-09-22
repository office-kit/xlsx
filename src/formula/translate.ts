// Shared / array formula translator. TS port of openpyxl
// `formula/translate.py`. Walks the tokens produced by `tokenize`, shifts
// the OPERAND-RANGE cell refs by `(row_delta, col_delta)`, and re-renders
// the formula. Absolute (`$`-prefixed) anchors stay put; falling off any edge
// of the grid raises `TranslatorError`. **Never evaluates.**

import { columnIndexFromLetter, columnLetterFromIndex, coordinateToTuple, MAX_ROW } from '../utils/coordinate.js';
import { OpenXmlError } from '../utils/exceptions.js';
import { LITERAL, OPERAND, RANGE, renderTokens, type Token, tokenize } from './tokenizer.js';

export class TranslatorError extends OpenXmlError {
  override readonly name = 'TranslatorError';
}

/** `1:1`, `$1234:78910`, etc. — pure-row range. */
export const ROW_RANGE_RE = /^(\$?[1-9][0-9]{0,6}):(\$?[1-9][0-9]{0,6})$/;
/** `A:A`, `$ABC:AZZ` — pure-column range. */
export const COL_RANGE_RE = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})$/;
/** `A1`, `$AB$15` — single cell ref. */
export const CELL_REF_RE = /^(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$/;

/**
 * Shift a row-snippet (`"3"` or `"$3"`) by `rdelta` rows. Absolute
 * anchors return verbatim; falling off either end of the grid raises
 * `TranslatorError`.
 *
 * The bottom bound is where this leaves openpyxl, which checks only for row 0
 * and below. {@link translateCol} has always raised at both ends, because
 * `columnLetterFromIndex` has no letters past XFD to hand back, so without the
 * check here a translation past row 1048576 was the one direction that quietly
 * produced a reference no spreadsheet can resolve.
 */
export function translateRow(rowStr: string, rdelta: number): string {
  if (rowStr.startsWith('$')) return rowStr;
  const newRow = Number.parseInt(rowStr, 10) + rdelta;
  if (newRow <= 0 || newRow > MAX_ROW) {
    throw new TranslatorError('Formula out of range');
  }
  return String(newRow);
}

/**
 * Shift a column-snippet (`"A"` or `"$A"`) by `cdelta` columns. Absolute
 * anchors return verbatim; out-of-range raises `TranslatorError`.
 */
export function translateCol(colStr: string, cdelta: number): string {
  if (colStr.startsWith('$')) return colStr;
  let idx: number;
  try {
    idx = columnIndexFromLetter(colStr);
  } catch {
    throw new TranslatorError('Formula out of range');
  }
  const newIdx = idx + cdelta;
  try {
    return columnLetterFromIndex(newIdx);
  } catch {
    throw new TranslatorError('Formula out of range');
  }
}

/**
 * Split `Sheet!A1` into `["Sheet!", "A1"]`. Multi-`!` sheet names are not
 * supported by Excel itself, so we just `rsplit('!', 1)`.
 */
export function stripWsName(rangeStr: string): [string, string] {
  const idx = rangeStr.lastIndexOf('!');
  if (idx === -1) return ['', rangeStr];
  return [`${rangeStr.slice(0, idx + 1)}`, rangeStr.slice(idx + 1)];
}

/**
 * Translate an A1-style range reference (potentially worksheet-prefixed,
 * potentially a named range) by `(rdelta, cdelta)`. Mirrors openpyxl
 * `Translator.translate_range` exactly:
 *
 * - `1:1` / `$1234:78910` → row-range, only rows shift
 * - `A:A` / `$ABC:AZZ` → col-range, only cols shift
 * - `A1:B2` (with `:`) → recurse on each side, allowing named-range endpoints
 * - `A1` → cell ref
 * - anything else → assumed named range, returned verbatim
 */
export function translateRange(rangeStr: string, rdelta: number, cdelta: number): string {
  const [wsPart, rest] = stripWsName(rangeStr);
  const rowMatch = ROW_RANGE_RE.exec(rest);
  if (rowMatch !== null) {
    return `${wsPart}${translateRow(rowMatch[1] as string, rdelta)}:${translateRow(rowMatch[2] as string, rdelta)}`;
  }
  const colMatch = COL_RANGE_RE.exec(rest);
  if (colMatch !== null) {
    return `${wsPart}${translateCol(colMatch[1] as string, cdelta)}:${translateCol(colMatch[2] as string, cdelta)}`;
  }
  if (rest.indexOf(':') !== -1) {
    const pieces = rest.split(':').map((p) => translateRange(p, rdelta, cdelta));
    return wsPart + pieces.join(':');
  }
  const cellMatch = CELL_REF_RE.exec(rest);
  if (cellMatch === null) return rest; // assume named range
  return `${wsPart}${translateCol(cellMatch[1] as string, cdelta)}${translateRow(cellMatch[2] as string, rdelta)}`;
}

export interface TranslateOptions {
  /** Destination cell address ("B2"). When set, derives `rowDelta`/`colDelta` from `origin`. */
  dest?: string;
  /** Explicit row delta. Ignored if `dest` is provided. */
  rowDelta?: number;
  /** Explicit col delta. Ignored if `dest` is provided. */
  colDelta?: number;
}

/**
 * Translate `formula` (defined at `origin`, e.g. "A1") to its destination
 * cell. Pass either `{ dest }` or `{ rowDelta, colDelta }`. LITERAL
 * formulas (input that does not start with `=`) and empty input pass
 * through untouched.
 */
export function translateFormula(formula: string, origin: string, opts: TranslateOptions = {}): string {
  const { col: originCol, row: originRow } = coordinateToTuple(origin);
  let rowDelta = opts.rowDelta ?? 0;
  let colDelta = opts.colDelta ?? 0;
  if (opts.dest !== undefined) {
    const { col: destCol, row: destRow } = coordinateToTuple(opts.dest);
    rowDelta = destRow - originRow;
    colDelta = destCol - originCol;
  }
  const tokens = tokenize(formula);
  if (tokens.length === 0) return '';
  const first = tokens[0] as Token;
  if (first.type === LITERAL) return first.value;
  let out = '=';
  for (const t of tokens) {
    if (t.type === OPERAND && t.subtype === RANGE) {
      out += translateRange(t.value, rowDelta, colDelta);
    } else {
      out += t.value;
    }
  }
  return out;
}

/**
 * Convenience: tokenize once, expose the parsed list + the same translate
 * helpers bound to a fixed origin. Mirrors openpyxl's `Translator` object.
 */
export interface Translator {
  formula: string;
  origin: string;
  row: number;
  col: number;
  tokens: Token[];
}

export function makeTranslator(formula: string, origin: string): Translator {
  const { row, col } = coordinateToTuple(origin);
  return { formula, origin, row, col, tokens: tokenize(formula) };
}

/** Render the translator's tokens back to the source string (sanity check). */
export function translatorRender(t: Translator): string {
  return renderTokens(t.tokens);
}
