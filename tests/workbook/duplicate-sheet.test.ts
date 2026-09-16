// Tests for the duplicateSheet workbook helper.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { setCellFont } from '../../src/styles/cell-style.js';
import { makeFont } from '../../src/styles/fonts.js';
import { addWorksheet, createWorkbook, duplicateSheet, getSheet } from '../../src/workbook/workbook.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import {
  setCell,
  setColumnWidth,
  setComment,
  type Worksheet,
  writeRange,
} from '../../src/worksheet/worksheet.js';

describe('duplicateSheet', () => {
  it('clones cells, dimensions, comments verbatim', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'header');
    setCell(a, 2, 1, 42);
    setColumnWidth(a, 1, 14);
    setComment(a, { ref: 'A1', author: 'Alice', text: 'note' });

    const b = duplicateSheet(wb, 'A', 'A copy');
    expect(b.title).toBe('A copy');
    expect(b.rows.get(1)?.get(1)?.value).toBe('header');
    expect(b.rows.get(2)?.get(1)?.value).toBe(42);
    expect(b.columnDimensions.get(1)?.width).toBe(14);
    expect(b.legacyComments[0]?.text).toBe('note');
  });

  it('clone is independent: mutations on copy do not affect source', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'orig');
    const b = duplicateSheet(wb, 'A', 'B');
    setCell(b, 1, 1, 'changed');
    expect(a.rows.get(1)?.get(1)?.value).toBe('orig');
    expect(b.rows.get(1)?.get(1)?.value).toBe('changed');
  });

  it('renumbers tables + suffixes displayName to keep workbook uniqueness', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'h');
    writeRange(a, 'A1', [['c1', 'c2']]);
    addExcelTable(wb, a, { name: 'Tbl', ref: 'A1:B2', columns: ['c1', 'c2'] });

    const b = duplicateSheet(wb, 'A', 'B');
    const aTable = a.tables[0];
    const bTable = b.tables[0];
    expect(aTable?.id).not.toBe(bTable?.id);
    expect(aTable?.displayName).toBe('Tbl');
    expect(bTable?.displayName).toBe('Tbl_2');
    expect(bTable?.id).toBe(2);
  });

  it('custom tableSuffix is honoured', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'h');
    writeRange(a, 'A1', [['c1', 'c2']]);
    addExcelTable(wb, a, { name: 'Tbl', ref: 'A1:B2', columns: ['c1', 'c2'] });
    const b = duplicateSheet(wb, 'A', 'B', { tableSuffix: '_dupe' });
    expect(b.tables[0]?.displayName).toBe('Tbl_dupe');
  });

  it('throws on duplicate target title or missing source', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    expect(() => duplicateSheet(wb, 'A', 'B')).toThrow(/already in use/);
    expect(() => duplicateSheet(wb, 'Missing', 'C')).toThrow(/no worksheet/);
  });

  it('insert at custom index', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    duplicateSheet(wb, 'A', 'A copy', { index: 1 });
    expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['A', 'A copy', 'B']);
  });

  it('cloned styled cells share styleId via the workbook stylesheet', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const cell = setCell(a, 1, 1, 'styled');
    setCellFont(wb, cell, makeFont({ bold: true }));
    const b = duplicateSheet(wb, 'A', 'B');
    const clonedCell = b.rows.get(1)?.get(1);
    expect(clonedCell?.styleId).toBe(cell.styleId);
  });

  it('duplicated sheet round-trips through saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'orig');
    duplicateSheet(wb, 'A', 'B');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const b2 = getSheet(wb2, 'B');
    expect(b2).toBeDefined();
    const sheet = b2 as Worksheet;
    expect(sheet.rows.get(1)?.get(1)?.value).toBe('orig');
  });
});
