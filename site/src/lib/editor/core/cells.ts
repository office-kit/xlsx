// Sparse-aware iteration over the library's `rows: Map<row, Map<col, Cell>>`.
//
// Ranges in a spreadsheet UI are routinely whole columns (1,048,576 rows) or
// whole rows (16,384 columns). Walking such a range index by index would stall
// the tab, so every helper here picks the cheaper of "walk the range" and "walk
// the populated cells", and never materialises a cell that does not exist.

import type { Cell, CellValue } from '@office-kit/xlsx/cell';
import { makeCell } from '@office-kit/xlsx/cell';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import type { Range } from './address.ts';

export function getCellAt(ws: Worksheet, row: number, col: number): Cell | undefined {
  return ws.rows.get(row)?.get(col);
}

/** Put `value` at (row, col), keeping an existing cell's format; a new cell takes `styleId`. */
export function setValueAt(ws: Worksheet, row: number, col: number, value: CellValue, styleId = 0): Cell {
  let rowMap = ws.rows.get(row);
  if (!rowMap) {
    rowMap = new Map();
    ws.rows.set(row, rowMap);
  }
  const existing = rowMap.get(col);
  if (existing) {
    existing.value = value;
    return existing;
  }
  const cell = makeCell(row, col, value, styleId);
  rowMap.set(col, cell);
  return cell;
}

/** Visit every existing cell inside `range`, row-major within the iteration order of the sparse maps. */
export function forEachCellInRange(ws: Worksheet, range: Range, fn: (cell: Cell) => void): void {
  const rowSpan = range.r2 - range.r1 + 1;
  const colSpan = range.c2 - range.c1 + 1;
  const visitRow = (rowMap: Map<number, Cell>): void => {
    if (colSpan < rowMap.size) {
      for (let c = range.c1; c <= range.c2; c++) {
        const cell = rowMap.get(c);
        if (cell) fn(cell);
      }
    } else {
      for (const [c, cell] of rowMap) if (c >= range.c1 && c <= range.c2) fn(cell);
    }
  };
  if (rowSpan < ws.rows.size) {
    for (let r = range.r1; r <= range.r2; r++) {
      const rowMap = ws.rows.get(r);
      if (rowMap) visitRow(rowMap);
    }
  } else {
    for (const [r, rowMap] of ws.rows) if (r >= range.r1 && r <= range.r2) visitRow(rowMap);
  }
}

export function cellsInRange(ws: Worksheet, range: Range): Cell[] {
  const out: Cell[] = [];
  forEachCellInRange(ws, range, (c) => out.push(c));
  return out;
}

/** Remove every cell object inside `range` (value, style and links alike). */
export function deleteCellsInRange(ws: Worksheet, range: Range): void {
  const doomed = cellsInRange(ws, range);
  for (const cell of doomed) {
    const rowMap = ws.rows.get(cell.row);
    if (!rowMap) continue;
    rowMap.delete(cell.col);
    if (rowMap.size === 0) ws.rows.delete(cell.row);
  }
}

/** True when the cell shows nothing: no value (a style alone does not count, as in Excel). */
export function isBlank(cell: Cell | undefined): boolean {
  return cell === undefined || cell.value === null || cell.value === '';
}

/** Populated extent of the sheet (cells with a value or a style), or undefined for an empty sheet. */
export function usedRange(ws: Worksheet): Range | undefined {
  let r1 = Infinity;
  let r2 = 0;
  let c1 = Infinity;
  let c2 = 0;
  for (const [r, rowMap] of ws.rows) {
    if (rowMap.size === 0) continue;
    if (r < r1) r1 = r;
    if (r > r2) r2 = r;
    for (const c of rowMap.keys()) {
      if (c < c1) c1 = c;
      if (c > c2) c2 = c;
    }
  }
  return r2 === 0 ? undefined : { r1, c1, r2, c2 };
}
