// A1-notation helpers shared by the lexer and the text-rewriting utilities.

import type { RefArea, SheetPrefix } from './ast.ts';
import { MAX_COL, MAX_ROW } from './types.ts';

export function columnFromLetters(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + ((letters.charCodeAt(i) | 0x20) - 96);
  return n;
}

export function lettersFromColumn(col: number): string {
  let s = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function cellAddress(row: number, col: number): string {
  return `${lettersFromColumn(col)}${row}`;
}

const colPart = (col: number, abs: boolean): string => `${abs ? '$' : ''}${lettersFromColumn(col)}`;
const rowPart = (row: number, abs: boolean): string => `${abs ? '$' : ''}${row}`;

/** Render a reference rectangle back to A1 text, keeping its anchors and written corner order. */
export function renderArea(a: RefArea): string {
  switch (a.kind) {
    case 'cell':
      return colPart(a.c1, a.c1Abs) + rowPart(a.r1, a.r1Abs);
    case 'area':
      return `${colPart(a.c1, a.c1Abs)}${rowPart(a.r1, a.r1Abs)}:${colPart(a.c2, a.c2Abs)}${rowPart(a.r2, a.r2Abs)}`;
    case 'cols':
      return `${colPart(a.c1, a.c1Abs)}:${colPart(a.c2, a.c2Abs)}`;
    case 'rows':
      return `${rowPart(a.r1, a.r1Abs)}:${rowPart(a.r2, a.r2Abs)}`;
  }
}

const PLAIN_SHEET_NAME = /^[A-Za-z_ -￿][\w. -￿]*$/u;
const LOOKS_LIKE_CELL = /^[A-Za-z]{1,3}\d+$/;
const LOOKS_LIKE_R1C1 = /^(?:R\d*C?\d*|C\d*)$/i;

/** Whether Excel requires `'…'` around `name` in a reference. */
export function sheetNameNeedsQuotes(name: string): boolean {
  if (!PLAIN_SHEET_NAME.test(name)) return true;
  if (LOOKS_LIKE_CELL.test(name) && columnFromLetters(name.replace(/\d+$/, '')) <= MAX_COL) return true;
  if (LOOKS_LIKE_R1C1.test(name)) return true;
  const upper = name.toUpperCase();
  return upper === 'TRUE' || upper === 'FALSE';
}

export function quoteSheetName(name: string): string {
  return sheetNameNeedsQuotes(name) ? `'${name.replace(/'/g, "''")}'` : name;
}

/** `Sheet1!`, `'My Sheet'!`, `Sheet1:Sheet3!`, `[1]Sheet1!`. */
export function renderPrefix(p: SheetPrefix): string {
  const external = p.external !== undefined ? `[${p.external}]` : '';
  const body = p.sheet2 !== undefined ? `${p.sheet}:${p.sheet2}` : p.sheet;
  const needsQuotes =
    p.external !== undefined
      ? /[^\w.]/.test(body) || sheetNameNeedsQuotes(p.sheet)
      : sheetNameNeedsQuotes(p.sheet) || (p.sheet2 !== undefined && sheetNameNeedsQuotes(p.sheet2));
  return needsQuotes ? `'${(external + body).replace(/'/g, "''")}'!` : `${external}${body}!`;
}

export function isValidRow(row: number): boolean {
  return Number.isInteger(row) && row >= 1 && row <= MAX_ROW;
}

export function isValidCol(col: number): boolean {
  return Number.isInteger(col) && col >= 1 && col <= MAX_COL;
}
