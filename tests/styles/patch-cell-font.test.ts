// Tests for patchCellFont, the merge path the single-field setters are built on.

import { describe, expect, it } from 'vitest';
import { getCellFont, patchCellFont, setCellFont } from '../../src/styles/cell-style.js';
import { DEFAULT_FONT, makeFont } from '../../src/styles/fonts.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';

describe('patchCellFont', () => {
  it('keeps the workbook default for the fields the patch omits', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'Header');
    patchCellFont(wb, c, { bold: true, size: 13 });
    const font = getCellFont(wb, c);
    expect(font.bold).toBe(true);
    expect(font.size).toBe(13);
    expect(font.name).toBe(DEFAULT_FONT.name);
    expect(font.scheme).toBe(DEFAULT_FONT.scheme);
  });

  it('is the merge counterpart to setCellFont, which replaces outright', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'Note');
    setCellFont(wb, c, makeFont({ bold: true }));
    expect(getCellFont(wb, c).name).toBeUndefined();
    expect(getCellFont(wb, c).size).toBeUndefined();
  });

  it('accumulates across calls', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'Note');
    patchCellFont(wb, c, { italic: true });
    patchCellFont(wb, c, { size: 9 });
    const font = getCellFont(wb, c);
    expect(font.italic).toBe(true);
    expect(font.size).toBe(9);
    expect(font.name).toBe(DEFAULT_FONT.name);
  });

  it('an empty patch leaves the font untouched', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 'x');
    patchCellFont(wb, c, { bold: true });
    const before = getCellFont(wb, c);
    patchCellFont(wb, c, {});
    expect(getCellFont(wb, c)).toEqual(before);
  });
});
