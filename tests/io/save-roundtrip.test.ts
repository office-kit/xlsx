import { describe, expect, it } from 'vitest';
import type { FormulaValue } from '../../src/cell/cell.js';
import { setFormula } from '../../src/cell/cell.js';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { saveWorkbook, workbookToBytes } from '../../src/io/save.js';
import { setCellFont } from '../../src/styles/cell-style.js';
import { makeFont } from '../../src/styles/fonts.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { ensureCell, getCell, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected worksheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('saveWorkbook → loadWorkbook round-trip', () => {
  it('preserves an empty single-sheet workbook', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Only');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.length).toBe(1);
    expect(wb2.sheets[0]?.sheet.title).toBe('Only');
    expect(expectSheet(wb2.sheets[0]?.sheet).rows.size).toBe(0);
  });

  it('preserves number / string / boolean / formula cells', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Mixed');
    setCell(ws, 1, 1, 42);
    setCell(ws, 1, 2, 'hello');
    setCell(ws, 1, 3, true);
    setCell(ws, 2, 1, 'with " < > & symbols');
    const cF = ensureCell(ws, 2, 2);
    setFormula(cF, 'A1+1', { cachedValue: 43 });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);

    expect(getCell(ws2, 1, 1)?.value).toBe(42);
    expect(getCell(ws2, 1, 2)?.value).toBe('hello');
    expect(getCell(ws2, 1, 3)?.value).toBe(true);
    expect(getCell(ws2, 2, 1)?.value).toBe('with " < > & symbols');

    const f = getCell(ws2, 2, 2)?.value as FormulaValue;
    expect(f.kind).toBe('formula');
    expect(f.formula).toBe('A1+1');
    expect(f.cachedValue).toBe(43);
  });

  it('preserves multiple sheets, including their order and titles', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Alpha');
    addWorksheet(wb, 'Beta');
    addWorksheet(wb, 'Gamma');
    setCell(wb.sheets[1]?.sheet as Worksheet, 1, 1, 100);

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));

    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(['Alpha', 'Beta', 'Gamma']);
    const beta = expectSheet(wb2.sheets[1]?.sheet);
    expect(getCell(beta, 1, 1)?.value).toBe(100);
  });

  it('preserves stylesheet pool + cell styleId references', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Styled');
    const c = setCell(ws, 1, 1, 'bold-text');
    setCellFont(wb, c, makeFont({ bold: true }));
    const styleIdBefore = c.styleId;

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);

    expect(wb2.styles.fonts.length).toBe(wb.styles.fonts.length);
    expect(wb2.styles.cellXfs.length).toBe(wb.styles.cellXfs.length);
    const c2 = getCell(ws2, 1, 1);
    expect(c2?.styleId).toBe(styleIdBefore);
    expect(wb2.styles.fonts[wb2.styles.cellXfs[c2?.styleId ?? 0]?.fontId ?? 0]?.bold).toBe(true);
  });

  it('saveWorkbook through an XlsxSink completes without error', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'S');
    setCell(wb.sheets[0]?.sheet as Worksheet, 1, 1, 7);
    const { toBuffer } = await import('../../src/io/node.js');
    const sink = toBuffer();
    await saveWorkbook(wb, sink);
    const wb2 = await loadWorkbook(fromBuffer(sink.result()));
    expect(getCell(expectSheet(wb2.sheets[0]?.sheet), 1, 1)?.value).toBe(7);
  });
});
