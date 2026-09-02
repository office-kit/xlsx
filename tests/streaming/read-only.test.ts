// Phase 4 read-only streaming acceptance.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { workbookToBytes } from '../../src/io/save';
import { loadWorkbookStream } from '../../src/streaming/read-only';
import { setCellAsDate } from '../../src/styles/cell-style';
import { builtinFormatCode, isDateFormat } from '../../src/styles/numbers';
import { dateToExcel } from '../../src/utils/datetime';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

const buildSampleWorkbook = async (rowCount: number): Promise<Uint8Array> => {
  const wb = createWorkbook();
  const a = addWorksheet(wb, 'Alpha');
  const b = addWorksheet(wb, 'Beta');
  for (let r = 1; r <= rowCount; r++) {
    setCell(a, r, 1, r);
    setCell(a, r, 2, `row-${r}`);
    setCell(a, r, 3, r % 2 === 0);
  }
  setCell(b, 1, 1, 'only');
  setCell(b, 1, 2, 42);
  return workbookToBytes(wb);
};

describe('loadWorkbookStream — sheet metadata', () => {
  it('lists sheet names without iterating cells', async () => {
    const bytes = await buildSampleWorkbook(3);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    expect(wb.sheetNames).toEqual(['Alpha', 'Beta']);
    await wb.close();
  });

  it('throws OpenXmlSchemaError for an unknown sheet name', async () => {
    const bytes = await buildSampleWorkbook(1);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    expect(() => wb.openWorksheet('NoSuchSheet')).toThrowError(/no worksheet named/);
    await wb.close();
  });
});

describe('loadWorkbookStream — iterRows', () => {
  it('streams every row + cell with correct row/col coordinates', async () => {
    const bytes = await buildSampleWorkbook(5);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('Alpha');
    const rows: { row: number; col: number; value: unknown }[][] = [];
    for await (const row of ws.iterRows()) {
      rows.push(row.map((c) => ({ row: c.row, col: c.col, value: c.value })));
    }
    expect(rows.length).toBe(5);
    expect(rows[0]).toEqual([
      { row: 1, col: 1, value: 1 },
      { row: 1, col: 2, value: 'row-1' },
      { row: 1, col: 3, value: false },
    ]);
    expect(rows[4]).toEqual([
      { row: 5, col: 1, value: 5 },
      { row: 5, col: 2, value: 'row-5' },
      { row: 5, col: 3, value: false },
    ]);
    await wb.close();
  });

  it('honours minRow / maxRow bounds', async () => {
    const bytes = await buildSampleWorkbook(10);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('Alpha');
    const rows: number[] = [];
    for await (const row of ws.iterRows({ minRow: 3, maxRow: 5 })) {
      const first = row[0];
      if (first) rows.push(first.row);
    }
    expect(rows).toEqual([3, 4, 5]);
    await wb.close();
  });

  it('honours minCol / maxCol bounds', async () => {
    const bytes = await buildSampleWorkbook(2);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('Alpha');
    const colsPerRow: number[][] = [];
    for await (const row of ws.iterRows({ minCol: 2, maxCol: 2 })) {
      colsPerRow.push(row.map((c) => c.col));
    }
    expect(colsPerRow).toEqual([[2], [2]]);
    await wb.close();
  });

  it('parses a Date cell as a Windows-epoch serial and preserves its date style', async () => {
    const source = createWorkbook();
    const sheet = addWorksheet(source, 'Dates');
    const date = new Date(Date.UTC(2026, 4, 5, 9, 30));
    const cell = setCell(sheet, 1, 1, date);
    setCellAsDate(source, cell);

    const wb = await loadWorkbookStream(fromBuffer(await workbookToBytes(source)));
    const rows = [];
    for await (const row of wb.openWorksheet('Dates').iterRows()) rows.push(row);

    expect(wb.date1904).toBe(false);
    expect(rows).toHaveLength(1);
    expect(rows[0]?.[0]).toEqual({
      row: 1,
      col: 1,
      value: dateToExcel(date, { epoch: 'windows' }),
      styleId: cell.styleId,
    });
    const numFmtId = wb.styles.cellXfs[cell.styleId]?.numFmtId;
    expect(numFmtId).toBeDefined();
    const formatCode =
      numFmtId === undefined ? undefined : wb.styles.numFmts.get(numFmtId) ?? builtinFormatCode(numFmtId);
    expect(isDateFormat(formatCode)).toBe(true);
    await wb.close();
  });

  it('exposes the Mac epoch when parsing a Date cell from a date1904 workbook', async () => {
    const source = createWorkbook({ date1904: true });
    const sheet = addWorksheet(source, 'Dates');
    const date = new Date(Date.UTC(2026, 4, 5, 9, 30));
    setCell(sheet, 1, 1, date);

    const wb = await loadWorkbookStream(fromBuffer(await workbookToBytes(source)));
    const rows = [];
    for await (const row of wb.openWorksheet('Dates').iterRows()) rows.push(row);

    expect(wb.date1904).toBe(true);
    expect(rows[0]?.[0]?.value).toBe(dateToExcel(date, { epoch: 'mac' }));
    await wb.close();
  });

  it('iterValues drops the cell envelope', async () => {
    const bytes = await buildSampleWorkbook(2);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('Beta');
    const rows: unknown[][] = [];
    for await (const row of ws.iterValues()) rows.push([...row]);
    expect(rows).toEqual([['only', 42]]);
    await wb.close();
  });

  it('two sheets iterated independently produce independent streams', async () => {
    const bytes = await buildSampleWorkbook(3);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const a = wb.openWorksheet('Alpha');
    const b = wb.openWorksheet('Beta');
    const aRows: unknown[][] = [];
    const bRows: unknown[][] = [];
    for await (const row of a.iterValues()) aRows.push([...row]);
    for await (const row of b.iterValues()) bRows.push([...row]);
    expect(aRows.length).toBe(3);
    expect(bRows.length).toBe(1);
    await wb.close();
  });
});

describe('loadWorkbookStream — shared strings + styles', () => {
  it('resolves t="s" cells via the sharedStrings table', async () => {
    const bytes = await buildSampleWorkbook(2);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    const ws = wb.openWorksheet('Alpha');
    let saw = '';
    for await (const row of ws.iterRows()) {
      for (const cell of row) {
        if (cell.col === 2 && cell.row === 1) saw = String(cell.value);
      }
    }
    expect(saw).toBe('row-1');
    await wb.close();
  });

  it('exposes the parsed Stylesheet for downstream lookups', async () => {
    const bytes = await buildSampleWorkbook(1);
    const wb = await loadWorkbookStream(fromBuffer(bytes));
    expect(wb.styles.fonts.length).toBeGreaterThan(0);
    expect(wb.styles.cellXfs.length).toBeGreaterThan(0);
    await wb.close();
  });
});
