// Tests for the addExcelTable builder helper.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import { setCell, type Worksheet, writeRange } from '../../src/worksheet/worksheet.js';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('addExcelTable', () => {
  it('builds a table with auto-assigned id and string-array columns', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sales');
    writeRange(ws, 'A1', [['Product', 'Region', 'Quantity', 'Price']]);
    const t = addExcelTable(wb, ws, {
      name: 'tblSales',
      ref: 'A1:D6',
      columns: ['Product', 'Region', 'Quantity', 'Price'],
      style: 'TableStyleMedium2',
    });
    expect(t.id).toBe(1);
    expect(t.displayName).toBe('tblSales');
    expect(t.name).toBe('tblSales');
    expect(t.ref).toBe('A1:D6');
    expect(t.columns).toHaveLength(4);
    expect(t.columns[0]).toEqual({ id: 1, name: 'Product' });
    expect(t.columns[3]).toEqual({ id: 4, name: 'Price' });
    expect(t.styleInfo?.name).toBe('TableStyleMedium2');
    expect(t.styleInfo?.showRowStripes).toBe(true);
    expect(ws.tables).toHaveLength(1);
  });

  it('auto-id ascends across multiple sheets in the same workbook', () => {
    const wb = createWorkbook();
    const ws1 = addWorksheet(wb, 'A');
    const ws2 = addWorksheet(wb, 'B');
    writeRange(ws1, 'A1', [['x', 'y']]);
    writeRange(ws1, 'D1', [['p', 'q']]);
    writeRange(ws2, 'A1', [['x', 'y']]);
    const t1 = addExcelTable(wb, ws1, { name: 't1', ref: 'A1:B5', columns: ['x', 'y'] });
    const t2 = addExcelTable(wb, ws2, { name: 't2', ref: 'A1:B5', columns: ['x', 'y'] });
    const t3 = addExcelTable(wb, ws1, { name: 't3', ref: 'D1:E5', columns: ['p', 'q'] });
    expect(t1.id).toBe(1);
    expect(t2.id).toBe(2);
    expect(t3.id).toBe(3);
  });

  it('full styleInfo override beats the simple `style` shortcut', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRange(ws, 'A1', [['x', 'y']]);
    const t = addExcelTable(wb, ws, {
      name: 't',
      ref: 'A1:B5',
      columns: ['x', 'y'],
      style: 'TableStyleMedium2', // ignored when styleInfo is set
      styleInfo: {
        name: 'TableStyleLight9',
        showRowStripes: false,
        showColumnStripes: true,
        showFirstColumn: true,
      },
    });
    expect(t.styleInfo?.name).toBe('TableStyleLight9');
    expect(t.styleInfo?.showColumnStripes).toBe(true);
    expect(t.styleInfo?.showFirstColumn).toBe(true);
  });

  it('TableColumn array passes through unchanged', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRange(ws, 'A1', [['Custom1', 'Custom2', 'Custom3']]);
    const t = addExcelTable(wb, ws, {
      name: 't',
      ref: 'A1:C5',
      columns: [
        { id: 10, name: 'Custom1' },
        { id: 20, name: 'Custom2', totalsRowFunction: 'sum' },
        { id: 30, name: 'Custom3' },
      ],
    });
    expect(t.columns[0]?.id).toBe(10);
    expect(t.columns[1]?.totalsRowFunction).toBe('sum');
  });

  it('rejects a column count that does not match the ref width', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRange(ws, 'A1', [['x', 'y']]);
    expect(() =>
      addExcelTable(wb, ws, { name: 't', ref: 'A1:C5', columns: ['x', 'y'] }),
    ).toThrow(OpenXmlSchemaError);
    expect(ws.tables).toHaveLength(0);
  });

  it('rejects header cells that disagree with the column names', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRange(ws, 'A1', [['x', 'WRONG']]);
    expect(() =>
      addExcelTable(wb, ws, { name: 't', ref: 'A1:B5', columns: ['x', 'y'] }),
    ).toThrow(/header cell B1 holds "WRONG" but column 2 is named "y"/);
  });

  it('rejects a missing header row, naming the escape hatch', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() =>
      addExcelTable(wb, ws, { name: 't', ref: 'A1:B5', columns: ['x', 'y'] }),
    ).toThrow(/headerRowCount: 0/);
  });

  it('skips the header check for a header-less table', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRange(ws, 'A1', [[1, 2]]);
    const t = addExcelTable(wb, ws, {
      name: 't',
      ref: 'A1:B5',
      columns: ['x', 'y'],
      headerRowCount: 0,
    });
    expect(t.headerRowCount).toBe(0);
  });

  it('full save → load round-trip preserves the table', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sales');
    setCell(ws, 1, 1, 'Product');
    setCell(ws, 1, 2, 'Quantity');
    setCell(ws, 2, 1, 'Apples');
    setCell(ws, 2, 2, 100);
    addExcelTable(wb, ws, {
      name: 'tblSales',
      ref: 'A1:B2',
      columns: ['Product', 'Quantity'],
      style: 'TableStyleMedium2',
    });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.tables).toHaveLength(1);
    expect(ws2.tables[0]?.displayName).toBe('tblSales');
    expect(ws2.tables[0]?.styleInfo?.name).toBe('TableStyleMedium2');
    expect(ws2.tables[0]?.columns).toHaveLength(2);
  });
});