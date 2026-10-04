import { describe, expect, test } from 'vitest';
import { addWorksheet, createWorkbook, makeDefinedName } from '@office-kit/xlsx/workbook';
import { addExcelTable, setCell } from '@office-kit/xlsx/worksheet';
import { MAX_ROW } from './address.ts';
import { currentRegion, dataEdge, lastUsedCell, navigationTree } from './navigation.ts';

describe('navigation pane', () => {
  test('lists tables and user names under the sheet they point at', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Summary');
    const data = addWorksheet(wb, 'Data');
    setCell(data, 1, 1, 'Item');
    setCell(data, 1, 2, 'Qty');
    setCell(data, 2, 1, 'a');
    setCell(data, 2, 2, 1);
    addExcelTable(wb, data, { name: 'Sales', ref: 'A1:B2', columns: ['Item', 'Qty'] });
    wb.definedNames.push(makeDefinedName({ name: 'Rate', value: 'Data!$D$1' }));
    // Built-in names are not shown, as in Excel.
    wb.definedNames.push(makeDefinedName({ name: '_xlnm.Print_Area', value: 'Data!$A$1:$B$2', scope: 1 }));

    const [summary, sheet] = navigationTree(wb);
    expect(summary?.elements).toEqual([]);
    expect(sheet?.elements.map((e) => [e.kind, e.label])).toEqual([
      ['table', 'Sales'],
      ['name', 'Rate'],
    ]);
    expect(sheet?.elements[1]?.target).toEqual({ kind: 'range', range: { r1: 1, c1: 4, r2: 1, c2: 4 } });
  });
});

describe('keyboard navigation over cell data', () => {
  // A1:B3 is a block, C4 touches it diagonally, then A6, A12 and F2 stand alone.
  function sheet() {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    for (const [r, c] of [[1, 1], [1, 2], [2, 1], [2, 2], [3, 1], [3, 2], [4, 3], [6, 1], [12, 1], [2, 6]] as const) setCell(ws, r, c, 'x');
    return ws;
  }

  test('Ctrl+Arrow runs to the end of a block, then jumps to the next filled cell or the sheet edge', () => {
    const ws = sheet();
    expect(dataEdge(ws, { row: 1, col: 1 }, 1, 0)).toEqual({ row: 3, col: 1 });
    expect(dataEdge(ws, { row: 3, col: 1 }, 1, 0)).toEqual({ row: 6, col: 1 });
    expect(dataEdge(ws, { row: 6, col: 1 }, 1, 0)).toEqual({ row: 12, col: 1 });
    expect(dataEdge(ws, { row: 12, col: 1 }, 1, 0)).toEqual({ row: MAX_ROW, col: 1 });
    expect(dataEdge(ws, { row: 6, col: 1 }, -1, 0)).toEqual({ row: 3, col: 1 });
    expect(dataEdge(ws, { row: 1, col: 1 }, -1, 0)).toEqual({ row: 1, col: 1 });
    expect(dataEdge(ws, { row: 2, col: 1 }, 0, 1)).toEqual({ row: 2, col: 2 });
    expect(dataEdge(ws, { row: 2, col: 2 }, 0, 1)).toEqual({ row: 2, col: 6 });
  });

  test('the current region takes in diagonal neighbours and stops at blank rows and columns', () => {
    const ws = sheet();
    expect(currentRegion(ws, { row: 2, col: 2 })).toEqual({ r1: 1, c1: 1, r2: 4, c2: 3 });
    expect(currentRegion(ws, { row: 6, col: 1 })).toEqual({ r1: 6, c1: 1, r2: 6, c2: 1 });
  });

  test('Ctrl+End goes to the last used row and the last used column, even when that cell is empty', () => {
    expect(lastUsedCell(sheet())).toEqual({ row: 12, col: 6 });
  });
});
