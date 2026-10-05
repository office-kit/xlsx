import { describe, expect, test } from 'vitest';
import { addWorksheet, createWorkbook, makeDefinedName } from '@office-kit/xlsx/workbook';
import { addExcelTable, setCell } from '@office-kit/xlsx/worksheet';
import { MAX_ROW } from './address.ts';
import { EditorController } from './controller.svelte.ts';
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

  test('the current region matches a full rescan on scattered data', () => {
    // Reference: grow until the border ring is blank, rescanning everything each pass.
    const naive = (ws: ReturnType<typeof addWorksheet>, row: number, col: number) => {
      let [r1, r2, c1, c2] = [row, row, col, col];
      const has = (r: number, c: number) => ws.rows.get(r)?.get(c) !== undefined;
      for (let grew = true; grew; ) {
        grew = false;
        const ring = (a1: number, a2: number, b1: number, b2: number) => {
          for (let r = a1; r <= a2; r++) for (let c = b1; c <= b2; c++) if (r >= 1 && c >= 1 && has(r, c)) return true;
          return false;
        };
        if (ring(r1 - 1, r1 - 1, c1 - 1, c2 + 1)) [r1, grew] = [r1 - 1, true];
        if (ring(r2 + 1, r2 + 1, c1 - 1, c2 + 1)) [r2, grew] = [r2 + 1, true];
        if (ring(r1 - 1, r2 + 1, c1 - 1, c1 - 1)) [c1, grew] = [c1 - 1, true];
        if (ring(r1 - 1, r2 + 1, c2 + 1, c2 + 1)) [c2, grew] = [c2 + 1, true];
      }
      return { r1: Math.max(1, r1), c1: Math.max(1, c1), r2, c2 };
    };
    let seed = 7;
    const rand = () => (seed = (seed * 48271) % 2147483647) / 2147483647;
    for (let trial = 0; trial < 40; trial++) {
      const ws = addWorksheet(createWorkbook(), 'S');
      for (let i = 0; i < 60; i++) setCell(ws, 1 + Math.floor(rand() * 14), 1 + Math.floor(rand() * 14), 1);
      for (let r = 1; r <= 15; r++) {
        for (let c = 1; c <= 15; c++) expect(currentRegion(ws, { row: r, col: c })).toEqual(naive(ws, r, c));
      }
    }
  });

  test('the current region of a tall block is found in linear time', () => {
    const ws = addWorksheet(createWorkbook(), 'S');
    for (let r = 1; r <= 100_000; r++) for (let c = 1; c <= 5; c++) setCell(ws, r, c, r);
    const t0 = performance.now();
    expect(currentRegion(ws, { row: 1, col: 1 })).toEqual({ r1: 1, c1: 1, r2: 100_000, c2: 5 });
    expect(performance.now() - t0).toBeLessThan(2000);
  });

  test('Ctrl+End goes to the last used row and the last used column, even when that cell is empty', () => {
    expect(lastUsedCell(sheet())).toEqual({ row: 12, col: 6 });
  });
});

describe('navigation skips hidden rows and columns', () => {
  test('Ctrl+Arrow stops at the last shown cell of a block', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    for (let r = 1; r <= 10; r++) setCell(ws, r, 1, r);
    const hidden = (r: number) => r >= 6 && r <= 10;
    expect(dataEdge(ws, { row: 1, col: 1 }, 1, 0, hidden)).toEqual({ row: 5, col: 1 });
    // A hidden gap inside a block does not split it.
    expect(dataEdge(ws, { row: 1, col: 1 }, 1, 0, (r) => r === 3)).toEqual({ row: 10, col: 1 });
    // Jumping from the block end over hidden data lands on the sheet edge.
    expect(dataEdge(ws, { row: 5, col: 1 }, 1, 0, hidden)).toEqual({ row: MAX_ROW, col: 1 });
  });

  test('Home and Ctrl+Home land on the first shown column and row', () => {
    const ctl = new EditorController();
    ctl.doc.ws.columnDimensions.set(1, { min: 1, max: 1, hidden: true });
    ctl.doc.ws.rowDimensions.set(1, { hidden: true });
    ctl.doc.layoutVersion++;
    ctl.selectCell({ row: 3, col: 4 });
    ctl.home(false, false);
    expect(ctl.doc.selection.active).toEqual({ row: 3, col: 2 });
    ctl.home(true, false);
    expect(ctl.doc.selection.active).toEqual({ row: 2, col: 2 });
  });
});
