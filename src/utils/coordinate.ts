// Worksheet coordinate utilities. Mirrors openpyxl/openpyxl/utils/cell.py.
//
// These functions are on the worksheet read/write hot path (millions of calls
// when streaming a sheet) so the implementations stay branch-light, regex-based
// only at the entry points, with bounded Map caches for the bidirectional
// column letter <-> index mapping.

import { OpenXmlSchemaError } from './exceptions.js';

/** Maximum column index Excel accepts (XFD). */
export const MAX_COL = 16384;
/** Maximum row index Excel accepts. */
export const MAX_ROW = 1048576;

// ---- column letter <-> 1-based index ---------------------------------------

const indexByLetter = new Map<string, number>();
const letterByIndex = new Map<number, string>();

/**
 * 1-based column index → spreadsheet column letter ("A", "Z", "AA", "XFD").
 * Throws OpenXmlSchemaError when out of range.
 */
export function columnLetterFromIndex(n: number): string {
  const cached = letterByIndex.get(n);
  if (cached !== undefined) return cached;
  if (!Number.isInteger(n) || n < 1 || n > MAX_COL) {
    throw new OpenXmlSchemaError(`column index ${n} is out of range [1, ${MAX_COL}]`);
  }
  let m = n;
  let out = '';
  while (m > 0) {
    m -= 1; // shift to 0-based for the modulo
    out = String.fromCharCode(65 + (m % 26)) + out;
    m = Math.floor(m / 26);
  }
  letterByIndex.set(n, out);
  return out;
}

/**
 * Column letter → 1-based column index. Case-insensitive but at most 3 letters
 * (the spec ceiling). Throws on empty / non-A-Z / over-range.
 */
export function columnIndexFromLetter(letter: string): number {
  const cached = indexByLetter.get(letter);
  if (cached !== undefined) return cached;
  if (letter.length === 0 || letter.length > 3) {
    throw new OpenXmlSchemaError(`column letter "${letter}" is empty or too long`);
  }
  let n = 0;
  for (let i = 0; i < letter.length; i++) {
    const c = letter.charCodeAt(i);
    let v: number;
    if (c >= 65 && c <= 90)
      v = c - 64; // 'A' = 65
    else if (c >= 97 && c <= 122)
      v = c - 96; // 'a' = 97
    else throw new OpenXmlSchemaError(`column letter "${letter}" contains non-letter char`);
    n = n * 26 + v;
  }
  if (n < 1 || n > MAX_COL) {
    throw new OpenXmlSchemaError(`column letter "${letter}" expands to out-of-range index ${n}`);
  }
  // Normalise the cache key to upper-case so 'a' / 'A' share a slot.
  const key = letter.toUpperCase();
  indexByLetter.set(key, n);
  if (key !== letter) indexByLetter.set(letter, n);
  return n;
}

// ---- coordinate parsing ----------------------------------------------------

/** A single-cell coordinate split into its letter and 1-based row. */
export interface CellCoordinate {
  column: string;
  row: number;
}

/** A single-cell coordinate split into 1-based numeric (col, row). */
export interface CellCoordinateNumeric {
  col: number;
  row: number;
}

/** A 1-based rectangular boundary. */
export interface CellRangeBoundaries {
  minCol: number;
  minRow: number;
  maxCol: number;
  maxRow: number;
}

/**
 * What every range-taking helper accepts: an A1 expression (`"A4:H20"`,
 * `"A:A"`, `"1:1"`) or already-computed 1-based bounds. Code that walks rows
 * and columns as integers passes the bounds directly instead of formatting a
 * string for the callee to parse straight back.
 */
export type RangeRef = string | CellRangeBoundaries;

const COORD_RE = /^[$]?([A-Za-z]{1,3})[$]?([1-9][0-9]*)$/;
const COL_RANGE_RE = /^[$]?([A-Za-z]{1,3}):[$]?([A-Za-z]{1,3})$/;
const ROW_RANGE_RE = /^[$]?([1-9][0-9]*):[$]?([1-9][0-9]*)$/;
const SHEET_RANGE_RE = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/;

/**
 * Parse a single-cell coordinate string ("A1", "$XFD$1048576") into its column
 * letter (always uppercased) and 1-based row.
 */
export function coordinateFromString(coord: string): CellCoordinate {
  const m = COORD_RE.exec(coord);
  if (m === null || m[1] === undefined || m[2] === undefined) {
    throw new OpenXmlSchemaError(`coordinateFromString: invalid coordinate "${coord}"`);
  }
  const column = m[1].toUpperCase();
  const row = Number.parseInt(m[2], 10);
  if (row < 1 || row > MAX_ROW) {
    throw new OpenXmlSchemaError(`coordinateFromString: row ${row} out of range`);
  }
  // Validate column upper bound through the cached helper.
  columnIndexFromLetter(column);
  return { column, row };
}

/**
 * Same as {@link coordinateFromString} but returning the column as its 1-based
 * numeric index. Thin convenience for the worksheet read path.
 */
export function coordinateToTuple(coord: string): CellCoordinateNumeric {
  const c = coordinateFromString(coord);
  return { col: columnIndexFromLetter(c.column), row: c.row };
}

/** `$` placement for a composed A1 reference. */
export interface AbsoluteRefOptions {
  /** Prefix the column letter with `$` so it survives a fill-across. */
  absoluteCol?: boolean;
  /** Prefix the row number with `$` so it survives a fill-down. */
  absoluteRow?: boolean;
}

/**
 * Compose `"A1"` from a 1-based (col, row). Set `absoluteCol` / `absoluteRow`
 * to get the `$` forms formula text needs (`"$B$5"`, `"B$5"`), so callers
 * tracking integer coordinates never have to concatenate a column letter
 * themselves.
 */
export function tupleToCoordinate(col: number, row: number, opts: AbsoluteRefOptions = {}): string {
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW) {
    throw new OpenXmlSchemaError(`tupleToCoordinate: row ${row} out of range`);
  }
  const colMarker = opts.absoluteCol === true ? '$' : '';
  const rowMarker = opts.absoluteRow === true ? '$' : '';
  return `${colMarker}${columnLetterFromIndex(col)}${rowMarker}${row}`;
}

/**
 * Predicate: true iff `s` is a valid single-cell A1 coordinate (`"A1"`,
 * `"XFD1048576"`). Strings with `$` absolute markers, surrounding whitespace,
 * ranges (`A1:B2`), or out-of-bound row / column return false. Useful for
 * sanitising user input before passing to {@link coordinateToTuple} or
 * `setCellByCoord`.
 */
export function isValidCellRef(s: unknown): s is string {
  if (typeof s !== 'string') return false;
  const m = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(s);
  if (!m) return false;
  const colStr = m[1];
  const rowStr = m[2];
  if (colStr === undefined || rowStr === undefined) return false;
  const col = columnIndexFromLetterUnchecked(colStr);
  if (col < 1 || col > MAX_COL) return false;
  const row = Number.parseInt(rowStr, 10);
  if (!Number.isFinite(row) || row < 1 || row > MAX_ROW) return false;
  return true;
}

/**
 * Predicate: true iff `s` is a valid A1-style range expression — single cell,
 * two-corner range, whole column (`A:A`), or whole row (`1:1`). `$` markers,
 * whitespace, and out-of-bound bounds fail. Sanity-check before {@link
 * rangeBoundaries} / `parseRange`.
 */
export function isValidRangeRef(s: unknown): s is string {
  if (typeof s !== 'string' || s.length === 0) return false;
  if (isValidCellRef(s)) return true;
  const ab = /^([A-Za-z]{1,3})([1-9][0-9]*):([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(s);
  if (ab) {
    const [, c1, r1, c2, r2] = ab;
    if (!c1 || !r1 || !c2 || !r2) return false;
    const col1 = columnIndexFromLetterUnchecked(c1);
    const col2 = columnIndexFromLetterUnchecked(c2);
    if (col1 < 1 || col1 > MAX_COL || col2 < 1 || col2 > MAX_COL) return false;
    const row1 = Number.parseInt(r1, 10);
    const row2 = Number.parseInt(r2, 10);
    return row1 >= 1 && row1 <= MAX_ROW && row2 >= 1 && row2 <= MAX_ROW;
  }
  const cc = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(s);
  if (cc) {
    const [, c1, c2] = cc;
    if (!c1 || !c2) return false;
    const col1 = columnIndexFromLetterUnchecked(c1);
    const col2 = columnIndexFromLetterUnchecked(c2);
    return col1 >= 1 && col1 <= MAX_COL && col2 >= 1 && col2 <= MAX_COL;
  }
  const rr = /^([1-9][0-9]*):([1-9][0-9]*)$/.exec(s);
  if (rr) {
    const [, r1, r2] = rr;
    if (!r1 || !r2) return false;
    const row1 = Number.parseInt(r1, 10);
    const row2 = Number.parseInt(r2, 10);
    return row1 >= 1 && row1 <= MAX_ROW && row2 >= 1 && row2 <= MAX_ROW;
  }
  return false;
}

const columnIndexFromLetterUnchecked = (letters: string): number => {
  let n = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    const ch = upper.charCodeAt(i);
    if (ch < 65 || ch > 90) return -1;
    n = n * 26 + (ch - 64);
  }
  return n;
};

/**
 * Predicate: true iff `s` is a valid 1..3-char column letter (`"A"` through
 * `"XFD"`, case-insensitive). Empty / over-long / out-of-bound / non-string
 * fails.
 */
export function isValidColumnLetter(s: unknown): s is string {
  if (typeof s !== 'string' || s.length === 0 || s.length > 3) return false;
  if (!/^[A-Za-z]+$/.test(s)) return false;
  const col = columnIndexFromLetterUnchecked(s);
  return col >= 1 && col <= MAX_COL;
}

/**
 * Predicate: true iff `n` is a valid 1-based row index in `[1, 1048576]`.
 * Non-finite / non-integer / out-of-bound fails.
 */
export function isValidRowNumber(n: unknown): n is number {
  return typeof n === 'number' && Number.isInteger(n) && n >= 1 && n <= MAX_ROW;
}

/**
 * Predicate: true iff `n` is a valid 1-based column index in `[1, 16384]`.
 * Non-finite / non-integer / out-of-bound fails.
 */
export function isValidColumnNumber(n: unknown): n is number {
  return typeof n === 'number' && Number.isInteger(n) && n >= 1 && n <= MAX_COL;
}

/**
 * Parse "A1:B5" / "A:A" / "1:1" / single-cell into 1-based (minCol, minRow,
 * maxCol, maxRow). Whole-column ranges fill rows to [1, MAX_ROW]; whole-row
 * ranges fill cols to [1, MAX_COL].
 */
export function rangeBoundaries(range: string): CellRangeBoundaries {
  const trimmed = range.trim();
  if (trimmed.length === 0) throw new OpenXmlSchemaError('rangeBoundaries: empty range');

  const colon = trimmed.indexOf(':');
  if (colon < 0) {
    // Single-cell shorthand: "A1" → A1:A1.
    const c = coordinateFromString(trimmed);
    const col = columnIndexFromLetter(c.column);
    return { minCol: col, minRow: c.row, maxCol: col, maxRow: c.row };
  }

  const left = trimmed.slice(0, colon);
  const right = trimmed.slice(colon + 1);

  const colsOnly = COL_RANGE_RE.exec(trimmed);
  if (colsOnly !== null && colsOnly[1] !== undefined && colsOnly[2] !== undefined) {
    const minCol = columnIndexFromLetter(colsOnly[1]);
    const maxCol = columnIndexFromLetter(colsOnly[2]);
    return {
      minCol: Math.min(minCol, maxCol),
      minRow: 1,
      maxCol: Math.max(minCol, maxCol),
      maxRow: MAX_ROW,
    };
  }

  const rowsOnly = ROW_RANGE_RE.exec(trimmed);
  if (rowsOnly !== null && rowsOnly[1] !== undefined && rowsOnly[2] !== undefined) {
    const minRow = Number.parseInt(rowsOnly[1], 10);
    const maxRow = Number.parseInt(rowsOnly[2], 10);
    if (minRow < 1 || maxRow < 1 || minRow > MAX_ROW || maxRow > MAX_ROW) {
      throw new OpenXmlSchemaError(`rangeBoundaries: row out of range in "${trimmed}"`);
    }
    return {
      minCol: 1,
      minRow: Math.min(minRow, maxRow),
      maxCol: MAX_COL,
      maxRow: Math.max(minRow, maxRow),
    };
  }

  const a = coordinateFromString(left);
  const b = coordinateFromString(right);
  const ac = columnIndexFromLetter(a.column);
  const bc = columnIndexFromLetter(b.column);
  return {
    minCol: Math.min(ac, bc),
    minRow: Math.min(a.row, b.row),
    maxCol: Math.max(ac, bc),
    maxRow: Math.max(a.row, b.row),
  };
}

/**
 * Inverse of {@link rangeBoundaries} for the rectangular case. `opts` forwards
 * to {@link tupleToCoordinate}, so `{ absoluteCol: true, absoluteRow: true }`
 * yields the `"$A$4:$H$20"` form a formula or defined name wants.
 */
export function boundariesToRangeString(b: CellRangeBoundaries, opts: AbsoluteRefOptions = {}): string {
  const tl = tupleToCoordinate(b.minCol, b.minRow, opts);
  if (b.minCol === b.maxCol && b.minRow === b.maxRow) return tl;
  const br = tupleToCoordinate(b.maxCol, b.maxRow, opts);
  return `${tl}:${br}`;
}

/**
 * Parse a sheet-qualified range ("Sheet1!A1:B5" / "'Quarter 1'!A1"). Sheet
 * names with single quotes inside use SQL-style doubling ("'Bob''s Sheet'!A1")
 * — we unescape on the way out.
 */
export function parseSheetRange(input: string): {
  sheet: string;
  range: string;
  bounds: CellRangeBoundaries;
} {
  const m = SHEET_RANGE_RE.exec(input);
  if (m === null) throw new OpenXmlSchemaError(`parseSheetRange: missing "!" delimiter in "${input}"`);
  const quoted = m[1];
  const bare = m[2];
  const range = m[3];
  if (range === undefined) throw new OpenXmlSchemaError(`parseSheetRange: empty range part in "${input}"`);
  const sheet = quoted !== undefined ? quoted.replace(/''/g, "'") : (bare ?? '');
  return { sheet, range, bounds: rangeBoundaries(range) };
}

const BARE_SHEET_NAME = /^[A-Za-z_][A-Za-z0-9_]*$/;

/**
 * Inverse of {@link parseSheetRange}: format a sheet title + range (or
 * single-cell ref) as `Sheet1!A1` or `'Quarter 1'!A1` per Excel's
 * sheet-qualified syntax. Single quotes inside the title are escaped by
 * doubling (`'Bob''s Sheet'`).
 *
 * Quoting rule: sheet titles consisting only of `[A-Za-z_][A-Za-z0-9_]*` are
 * emitted bare; everything else (spaces, digits-leading, hyphens, apostrophes,
 * punctuation…) gets wrapped in single quotes.
 */
export function formatSheetQualifiedRef(sheet: string, ref: string): string {
  if (BARE_SHEET_NAME.test(sheet)) return `${sheet}!${ref}`;
  return `'${sheet.replace(/'/g, "''")}'!${ref}`;
}
