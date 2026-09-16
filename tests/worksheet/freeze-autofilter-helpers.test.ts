// Tests for freezeRows / freezeColumns / setFreezePanes / unfreezePanes
// and addAutoFilter / addAutoFilterColumn helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  addAutoFilter,
  addAutoFilterColumn,
  removeAutoFilter,
} from '../../src/worksheet/auto-filter.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import {
  freezeColumns,
  freezeRows,
  getFreezePanes,
  setCell,
  setFreezePanes,
  unfreezePanes,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('freezeRows / freezeColumns / setFreezePanes', () => {
  it('freezeRows(1) freezes the first row, ref="A2"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezeRows(ws, 1);
    expect(getFreezePanes(ws)).toBe('A2');
  });

  it('freezeColumns(2) freezes A and B, ref="C1"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezeColumns(ws, 2);
    expect(getFreezePanes(ws)).toBe('C1');
  });

  it('setFreezePanes({ rows: 2, cols: 3 }) freezes top 2 rows and left 3 cols, ref="D3"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setFreezePanes(ws, { rows: 2, cols: 3 });
    expect(getFreezePanes(ws)).toBe('D3');
  });

  it('setFreezePanes accepts a zero on either axis; A1 and negatives throw', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setFreezePanes(ws, { rows: 2, cols: 0 });
    expect(getFreezePanes(ws)).toBe('A3');
    setFreezePanes(ws, { rows: 0, cols: 2 });
    expect(getFreezePanes(ws)).toBe('C1');
    expect(() => setFreezePanes(ws, { rows: 0, cols: 0 })).toThrow(OpenXmlSchemaError);
    expect(() => setFreezePanes(ws, { rows: -1, cols: 1 })).toThrow(OpenXmlSchemaError);
    expect(() => setFreezePanes(ws, { rows: 1, cols: 1.5 })).toThrow(OpenXmlSchemaError);
  });

  it('unfreezePanes drops the freeze', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezeRows(ws, 1);
    expect(getFreezePanes(ws)).toBe('A2');
    unfreezePanes(ws);
    expect(getFreezePanes(ws)).toBeUndefined();
  });

  it('rejects non-positive counts', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => freezeRows(ws, 0)).toThrow();
    expect(() => freezeColumns(ws, -1)).toThrow();
  });

  it('full save → load round-trip preserves the freeze ref', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 1);
    setFreezePanes(ws, { rows: 1, cols: 1 });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(getFreezePanes(ws2)).toBe('B2');
  });
});

describe('addAutoFilter / addAutoFilterColumn / removeAutoFilter', () => {
  it('addAutoFilter sets the filter ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    const af = addAutoFilter(ws, 'A1:E10');
    expect(af.ref).toBe('A1:E10');
    expect(af.filterColumns).toEqual([]);
    expect(ws.autoFilter?.ref).toBe('A1:E10');
  });

  it('addAutoFilterColumn appends a value-list filter', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    addAutoFilter(ws, 'A1:E10');
    addAutoFilterColumn(ws, 0, ['Open', 'In Progress']);
    addAutoFilterColumn(ws, 2, ['1', '2', '3'], { blank: true });
    expect(ws.autoFilter?.filterColumns.length).toBe(2);
    expect(ws.autoFilter?.filterColumns[0]?.colId).toBe(0);
    expect(ws.autoFilter?.filterColumns[0]?.values).toEqual(['Open', 'In Progress']);
    expect(ws.autoFilter?.filterColumns[1]?.blank).toBe(true);
  });

  it('addAutoFilterColumn throws OpenXmlSchemaError if no autoFilter is set', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    expect(() => addAutoFilterColumn(ws, 0, ['x'])).toThrow(OpenXmlSchemaError);
  });

  it('removeAutoFilter drops the autoFilter', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    addAutoFilter(ws, 'A1:C10');
    removeAutoFilter(ws);
    expect(ws.autoFilter).toBeUndefined();
  });

  it('survives a save → load round-trip', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setCell(ws, 1, 1, 'Status');
    addAutoFilter(ws, 'A1:A5');
    addAutoFilterColumn(ws, 0, ['Open', 'Closed']);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.autoFilter?.ref).toBe('A1:A5');
    expect(ws2.autoFilter?.filterColumns[0]?.values).toEqual(['Open', 'Closed']);
  });
});