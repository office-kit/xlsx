import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { makeAutoFilter, makeFilterColumn } from '../../src/worksheet/auto-filter.js';
import { makeTableColumn, makeTableDefinition } from '../../src/worksheet/table.js';
import { addTable, getTable, removeTable, setCell, type Worksheet, writeRange } from '../../src/worksheet/worksheet.js';

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('addTable / getTable / removeTable', () => {
  it('starts empty and accepts an addTable', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'T');
    expect(ws.tables.length).toBe(0);
    writeRange(ws, 'A1', [['Region', 'Quarter', 'Total']]);
    addTable(
      ws,
      makeTableDefinition({
        id: 1,
        displayName: 'Sales',
        ref: 'A1:C10',
        columns: [
          makeTableColumn({ id: 1, name: 'Region' }),
          makeTableColumn({ id: 2, name: 'Quarter' }),
          makeTableColumn({ id: 3, name: 'Total' }),
        ],
      }),
    );
    expect(getTable(ws, 'Sales')?.columns.length).toBe(3);
  });

  it('rejects a definition that disagrees with the cells under it', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'T');
    writeRange(ws, 'A1', [['Region']]);
    expect(() =>
      addTable(
        ws,
        makeTableDefinition({
          id: 1,
          displayName: 'Sales',
          ref: 'A1:A2',
          columns: [makeTableColumn({ id: 1, name: 'Area' })],
        }),
      ),
    ).toThrow(OpenXmlSchemaError);
    expect(ws.tables).toHaveLength(0);
  });

  it('removeTable returns true on hit, false on miss', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'T');
    setCell(ws, 1, 1, 'x');
    addTable(
      ws,
      makeTableDefinition({
        id: 1,
        displayName: 'X',
        ref: 'A1:A2',
        columns: [makeTableColumn({ id: 1, name: 'x' })],
      }),
    );
    expect(removeTable(ws, 'X')).toBe(true);
    expect(removeTable(ws, 'X')).toBe(false);
    expect(ws.tables.length).toBe(0);
  });
});

describe('table round-trip through saveWorkbook → loadWorkbook', () => {
  it('preserves a single-sheet table with columns + style + autoFilter', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'T');
    setCell(ws, 1, 1, 'Region');
    setCell(ws, 1, 2, 'Total');
    setCell(ws, 2, 1, 'EU');
    setCell(ws, 2, 2, 100);
    addTable(
      ws,
      makeTableDefinition({
        id: 1,
        displayName: 'Sales',
        ref: 'A1:B2',
        columns: [makeTableColumn({ id: 1, name: 'Region' }), makeTableColumn({ id: 2, name: 'Total' })],
        styleInfo: { name: 'TableStyleMedium2', showRowStripes: true },
        autoFilter: makeAutoFilter({
          ref: 'A1:B2',
          filterColumns: [makeFilterColumn({ colId: 0, values: ['EU'] })],
        }),
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const t = getTable(ws2, 'Sales');
    expect(t?.ref).toBe('A1:B2');
    expect(t?.columns.length).toBe(2);
    expect(t?.columns[0]?.name).toBe('Region');
    expect(t?.styleInfo?.name).toBe('TableStyleMedium2');
    expect(t?.styleInfo?.showRowStripes).toBe(true);
    expect(t?.autoFilter?.filterColumns[0]).toHaveProperty('values', ['EU']);
    expect(t?.rId).toMatch(/^rId\d+$/);
  });

  it('preserves multiple tables across the workbook with workbook-global ids', async () => {
    const wb = createWorkbook();
    const wsA = addWorksheet(wb, 'A');
    const wsB = addWorksheet(wb, 'B');
    writeRange(wsA, 'A1', [['x', 'y']]);
    writeRange(wsA, 'D1', [['p', 'q']]);
    writeRange(wsB, 'A1', [['v']]);
    addTable(
      wsA,
      makeTableDefinition({
        id: 1,
        displayName: 'TableA1',
        ref: 'A1:B2',
        columns: [makeTableColumn({ id: 1, name: 'x' }), makeTableColumn({ id: 2, name: 'y' })],
      }),
    );
    addTable(
      wsA,
      makeTableDefinition({
        id: 2,
        displayName: 'TableA2',
        ref: 'D1:E2',
        columns: [makeTableColumn({ id: 1, name: 'p' }), makeTableColumn({ id: 2, name: 'q' })],
      }),
    );
    addTable(
      wsB,
      makeTableDefinition({
        id: 3,
        displayName: 'TableB1',
        ref: 'A1:A5',
        columns: [makeTableColumn({ id: 1, name: 'v' })],
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const tablesA = (wb2.sheets[0]?.sheet as Worksheet).tables;
    const tablesB = (wb2.sheets[1]?.sheet as Worksheet).tables;
    expect(tablesA.map((t) => t.displayName).sort()).toEqual(['TableA1', 'TableA2']);
    expect(tablesB.map((t) => t.displayName)).toEqual(['TableB1']);
    // ids on the wire are unique workbook-wide.
    const allIds = [...tablesA, ...tablesB].map((t) => t.id);
    expect(new Set(allIds).size).toBe(3);
  });

  it('preserves headerRowCount + totalsRowCount + totalsRowShown', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'T');
    writeRange(ws, 'A1', [['a', 'b']]);
    addTable(
      ws,
      makeTableDefinition({
        id: 1,
        displayName: 'Q',
        ref: 'A1:B5',
        columns: [makeTableColumn({ id: 1, name: 'a' }), makeTableColumn({ id: 2, name: 'b' })],
        headerRowCount: 1,
        totalsRowCount: 1,
        totalsRowShown: true,
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const t = getTable(expectSheet(wb2.sheets[0]?.sheet), 'Q');
    expect(t?.headerRowCount).toBe(1);
    expect(t?.totalsRowCount).toBe(1);
    expect(t?.totalsRowShown).toBe(true);
  });

  it('omits <tableParts> when the worksheet has no tables', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'NoTable');
    const bytes = await workbookToBytes(wb);
    const txt = new TextDecoder().decode(bytes);
    expect(txt).not.toContain('<tableParts');
    expect(txt).not.toContain('<tablePart ');
  });
});
