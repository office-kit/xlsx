// Tests for registerCellStyle: build a style once, apply it by id on write.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import type { Cell } from '../../src/cell/cell.js';
import { makeBorder, makeSide } from '../../src/styles/borders.js';
import {
  getCellBorder,
  getCellFont,
  getCellNumberFormat,
  registerCellStyle,
  setCellStyle,
} from '../../src/styles/cell-style.js';
import { makeFont } from '../../src/styles/fonts.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import { appendRow, getCell, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const THIN = makeBorder({
  left: makeSide({ style: 'thin' }),
  right: makeSide({ style: 'thin' }),
  top: makeSide({ style: 'thin' }),
  bottom: makeSide({ style: 'thin' }),
});

const sheetOrThrow = (ws: Worksheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected worksheet');
  return ws;
};

const cellOrThrow = (c: Cell | undefined): Cell => {
  if (!c) throw new Error('expected cell');
  return c;
};

describe('registerCellStyle', () => {
  it('returns an id that setCell applies verbatim', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const id = registerCellStyle(wb, { numberFormat: '#,##0', border: THIN });
    const c = setCell(ws, 1, 1, 1234, id);
    expect(c.styleId).toBe(id);
    expect(getCellNumberFormat(wb, c)).toBe('#,##0');
    expect(getCellBorder(wb, c).left?.style).toBe('thin');
  });

  it('dedups equal specs to one xf and reserves slot 0 for the default', () => {
    const wb = createWorkbook();
    const spec = { font: makeFont({ name: 'Calibri', size: 11, bold: true }) };
    const first = registerCellStyle(wb, spec);
    const second = registerCellStyle(wb, spec);
    expect(second).toBe(first);
    expect(first).not.toBe(0);
    expect(wb.styles.cellXfs).toHaveLength(2);
  });

  it('an empty spec resolves to the default xf', () => {
    const wb = createWorkbook();
    expect(registerCellStyle(wb, {})).toBe(0);
  });

  it('describes a whole style, not a patch over the target cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 1);
    setCellStyle(wb, c, { numberFormat: '0.00', font: makeFont({ bold: true }) });
    // The id carries only a number format, so the bold is not inherited.
    const numberOnly = registerCellStyle(wb, { numberFormat: '#,##0' });
    setCell(ws, 1, 1, 2, numberOnly);
    expect(getCellNumberFormat(wb, c)).toBe('#,##0');
    expect(getCellFont(wb, c).bold).toBeUndefined();
  });

  it('appendRow applies per-column ids and materialises styled blanks', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    const intId = registerCellStyle(wb, { numberFormat: '#,##0', border: THIN });
    const inputId = registerCellStyle(wb, { border: THIN });

    appendRow(ws, ['Widgets', 1200], { styleIds: [undefined, intId, inputId] });

    expect(getCell(ws, 1, 2)?.styleId).toBe(intId);
    const blank = getCell(ws, 1, 3);
    expect(blank?.value).toBeNull();
    expect(blank?.styleId).toBe(inputId);

    const reloaded = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    const sheet = sheetOrThrow(getSheet(reloaded, 'Data'));
    expect(getCellNumberFormat(reloaded, cellOrThrow(getCell(sheet, 1, 2)))).toBe('#,##0');
    expect(getCellBorder(reloaded, cellOrThrow(getCell(sheet, 1, 3))).top?.style).toBe('thin');
  });
});
