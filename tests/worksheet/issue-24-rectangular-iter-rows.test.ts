// Coverage for the #24 redesign: iterRows / iterValues yield rectangular
// rows over the populated bounding box, not just the populated cells.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { iterCells, iterRows, iterValues, setCell } from '../../src/worksheet/worksheet.js';

describe('iterRows — rectangular default (#24)', () => {
  it('default extent is the populated bounding box', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 3, 3, 'c');
    // Populated extent: rows 1..3, cols 1..3 → 3x3 rectangle.
    const rows = [...iterRows(ws)].map((row) => row.map((c) => c?.value ?? null));
    expect(rows).toEqual([
      ['a', null, null],
      [null, null, null],
      [null, null, 'c'],
    ]);
  });

  it('yields one row per row in [minRow, maxRow] including entirely empty rows', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 4, 1, 'd');
    const rows = [...iterRows(ws, { minRow: 1, maxRow: 4, minCol: 1, maxCol: 1 })].map((row) =>
      row.map((c) => c?.value ?? null),
    );
    expect(rows.length).toBe(4);
    expect(rows).toEqual([['a'], [null], [null], ['d']]);
  });

  it('rectangular rows match maxCol - minCol + 1 in width', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 3, 'c');
    const rows = [...iterRows(ws, { minRow: 1, maxRow: 1, minCol: 1, maxCol: 5 })];
    expect(rows.length).toBe(1);
    expect(rows[0]?.length).toBe(5);
  });

  it('explicit maxRow past populated extent emits trailing empty rows', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    const rows = [...iterRows(ws, { minRow: 1, maxRow: 4, minCol: 1, maxCol: 1 })].map((row) =>
      row.map((c) => c?.value ?? null),
    );
    expect(rows).toEqual([['a'], [null], [null], [null]]);
  });

  it('inverted ranges yield nothing without throwing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    expect([...iterRows(ws, { minRow: 5, maxRow: 2 })]).toEqual([]);
    expect([...iterRows(ws, { minCol: 5, maxCol: 2 })]).toEqual([]);
  });

  it('empty worksheet yields no rows under default extent', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    expect([...iterRows(ws)]).toEqual([]);
  });
});

describe('iterValues — rectangular (#24)', () => {
  it('fills missing cells with null, not undefined', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 3, 'c');
    setCell(ws, 3, 2, 'b');
    const values = [...iterValues(ws)];
    expect(values).toEqual([
      ['a', null, 'c'],
      [null, null, null],
      [null, 'b', null],
    ]);
    // null is a valid CellValue; undefined is not.
    for (const row of values) {
      for (const v of row) expect(v === undefined).toBe(false);
    }
  });

  it('round-trips: re-applying via setCell reproduces the same iterValues output', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 3, 'c');
    setCell(ws, 3, 2, 'b');
    const before = [...iterValues(ws)];

    const wb2 = createWorkbook();
    const ws2 = addWorksheet(wb2, 'S2');
    for (let r = 0; r < before.length; r++) {
      const row = before[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (v !== null && v !== undefined) setCell(ws2, r + 1, c + 1, v);
      }
    }
    const after = [...iterValues(ws2)];
    expect(after).toEqual(before);
  });
});

describe('iterCells — populated only, unchanged behavior (#24)', () => {
  it('still skips undefined positions; yields only populated cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 3, 3, 'c');
    const cells = [...iterCells(ws)];
    expect(cells.map((c) => [c.row, c.col, c.value])).toEqual([
      [1, 1, 'a'],
      [3, 3, 'c'],
    ]);
  });
});
