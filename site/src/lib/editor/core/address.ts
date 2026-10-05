// Grid coordinates shared by every editor module. Rows and columns are 1-based
// like the library model, and a Range is always normalised (r1 <= r2,
// c1 <= c2) so consumers never re-check orientation.

import { columnIndexFromLetter, columnLetterFromIndex, MAX_COL, MAX_ROW } from '@office-kit/xlsx/utils';

export { MAX_COL, MAX_ROW };

export interface CellPos {
  readonly row: number;
  readonly col: number;
}

export interface Range {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

export const colLetter = columnLetterFromIndex;

export function clampPos(row: number, col: number): CellPos {
  return { row: Math.min(MAX_ROW, Math.max(1, row)), col: Math.min(MAX_COL, Math.max(1, col)) };
}

export function rangeOf(a: CellPos, b: CellPos = a): Range {
  return {
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.col, b.col),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.col, b.col),
  };
}

export function cellAddress(row: number, col: number, absolute = false): string {
  return absolute ? `$${colLetter(col)}$${row}` : `${colLetter(col)}${row}`;
}

export function isWholeColumns(r: Range): boolean {
  return r.r1 === 1 && r.r2 === MAX_ROW;
}

export function isWholeRows(r: Range): boolean {
  return r.c1 === 1 && r.c2 === MAX_COL;
}

/** A1-style text for a range the way Excel's Name Box prints it (`A:A`, `1:3`, `B2:C4`, `B2`). */
export function rangeAddress(r: Range, absolute = false): string {
  const d = absolute ? '$' : '';
  if (isWholeColumns(r) && !isWholeRows(r)) return `${d}${colLetter(r.c1)}:${d}${colLetter(r.c2)}`;
  if (isWholeRows(r) && !isWholeColumns(r)) return `${d}${r.r1}:${d}${r.r2}`;
  if (r.r1 === r.r2 && r.c1 === r.c2) return cellAddress(r.r1, r.c1, absolute);
  return `${cellAddress(r.r1, r.c1, absolute)}:${cellAddress(r.r2, r.c2, absolute)}`;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const COLS_RE = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/;
const ROWS_RE = /^\$?(\d{1,7}):\$?(\d{1,7})$/;

function parseCell(text: string): CellPos | undefined {
  const m = CELL_RE.exec(text);
  if (!m?.[1] || !m[2]) return undefined;
  const col = columnIndexFromLetter(m[1].toUpperCase());
  const row = Number(m[2]);
  if (row < 1 || row > MAX_ROW || col < 1 || col > MAX_COL) return undefined;
  return { row, col };
}

/**
 * Parse what a user types into the Name Box / Go To dialog: `B3`, `B3:D9`,
 * `C:E`, `4:9`. Sheet-qualified references return the sheet separately.
 * Returns undefined for anything else (defined names are resolved by the
 * caller, which knows the workbook).
 */
export function parseRangeAddress(input: string): { sheet?: string; range: Range } | undefined {
  let text = input.trim();
  let sheet: string | undefined;
  const bang = text.lastIndexOf('!');
  if (bang > 0) {
    sheet = text.slice(0, bang);
    if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replaceAll("''", "'");
    text = text.slice(bang + 1);
  }
  const withSheet = (range: Range) => (sheet === undefined ? { range } : { sheet, range });
  const cols = COLS_RE.exec(text);
  if (cols?.[1] && cols[2]) {
    const a = columnIndexFromLetter(cols[1].toUpperCase());
    const b = columnIndexFromLetter(cols[2].toUpperCase());
    if (a > MAX_COL || b > MAX_COL) return undefined;
    return withSheet({ r1: 1, r2: MAX_ROW, c1: Math.min(a, b), c2: Math.max(a, b) });
  }
  const rows = ROWS_RE.exec(text);
  if (rows?.[1] && rows[2]) {
    const a = Number(rows[1]);
    const b = Number(rows[2]);
    if (a < 1 || b < 1 || a > MAX_ROW || b > MAX_ROW) return undefined;
    return withSheet({ r1: Math.min(a, b), r2: Math.max(a, b), c1: 1, c2: MAX_COL });
  }
  const [first, second, ...rest] = text.split(':');
  if (first === undefined || rest.length > 0) return undefined;
  const a = parseCell(first);
  if (!a) return undefined;
  const b = second === undefined ? a : parseCell(second);
  if (!b) return undefined;
  return withSheet(rangeOf(a, b));
}

export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name
    : `'${name.replaceAll("'", "''")}'`;
}

export function inRange(r: Range, row: number, col: number): boolean {
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

export function rangesIntersect(a: Range, b: Range): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

export function unionRange(a: Range, b: Range): Range {
  return {
    r1: Math.min(a.r1, b.r1),
    c1: Math.min(a.c1, b.c1),
    r2: Math.max(a.r2, b.r2),
    c2: Math.max(a.c2, b.c2),
  };
}

export function rangeCellCount(r: Range): number {
  return (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
}

export function sameRange(a: Range, b: Range): boolean {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

/** The library's range boundary shape (`mergedCells`, `RangeRef` objects). */
export interface Boundaries {
  minRow: number;
  minCol: number;
  maxRow: number;
  maxCol: number;
}

export function fromBoundaries(b: Boundaries): Range {
  return { r1: b.minRow, c1: b.minCol, r2: b.maxRow, c2: b.maxCol };
}

export function toBoundaries(r: Range): Boundaries {
  return { minRow: r.r1, minCol: r.c1, maxRow: r.r2, maxCol: r.c2 };
}
