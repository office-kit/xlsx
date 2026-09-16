// Tests for passing numeric bounds where an A1 range string is accepted.
//
// A caller that already tracks rows and columns as integers should not have to
// format a string for the callee to parse straight back.

import { describe, expect, it } from 'vitest';
import { makeBorder, makeSide } from '../../src/styles/borders.js';
import {
  getCellBorder,
  getCellNumberFormat,
  setRangeNumberFormat,
  setRangeStyle,
} from '../../src/styles/cell-style.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  applyToRange,
  clearRange,
  getCell,
  getRangeAddress,
  getRangeValues,
  setCell,
  setRangeValues,
  writeRange,
} from '../../src/worksheet/worksheet.js';

const bounds = { minRow: 2, minCol: 2, maxRow: 3, maxCol: 3 };

describe('numeric bounds as a range ref', () => {
  it('setRangeValues / getRangeValues agree with the A1 form', () => {
    const wb = createWorkbook();
    const numeric = addWorksheet(wb, 'N');
    const a1 = addWorksheet(wb, 'S');
    const values = [
      [1, 2],
      [3, 4],
    ];
    setRangeValues(numeric, bounds, values);
    setRangeValues(a1, 'B2:C3', values);
    expect(getRangeValues(numeric, bounds)).toEqual(getRangeValues(a1, 'B2:C3'));
    expect(getRangeValues(numeric, bounds)).toEqual(values);
  });

  it('applyToRange visits every coordinate and allocates on first touch', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const seen: string[] = [];
    applyToRange(ws, bounds, (_cell, row, col) => seen.push(`${row},${col}`));
    expect(seen).toEqual(['2,2', '2,3', '3,2', '3,3']);
    expect(getCell(ws, 3, 3)?.value).toBeNull();
  });

  it('clearRange counts the same cells either way', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeValues(ws, bounds, [
      [1, 2],
      [3, 4],
    ]);
    expect(clearRange(ws, bounds)).toBe(4);
  });

  it('style helpers accept bounds', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setRangeNumberFormat(wb, ws, bounds, '#,##0');
    setRangeStyle(wb, ws, bounds, { border: makeBorder({ top: makeSide({ style: 'thin' }) }) });
    const c = getCell(ws, 2, 2);
    expect(c).toBeDefined();
    if (!c) return;
    expect(getCellNumberFormat(wb, c)).toBe('#,##0');
    expect(getCellBorder(wb, c).top?.style).toBe('thin');
  });

  it('normalises inverted bounds like the A1 parser does', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'x');
    setCell(ws, 2, 2, 'y');
    expect(clearRange(ws, { minRow: 2, minCol: 2, maxRow: 1, maxCol: 1 })).toBe(2);
  });

  it('writeRange takes a numeric anchor', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const written = writeRange(ws, { row: 2, col: 3 }, [['a', 'b']]);
    expect(written).toEqual({ minRow: 2, maxRow: 2, minCol: 3, maxCol: 4 });
    expect(getCell(ws, 2, 4)?.value).toBe('b');
  });

  it('getRangeAddress formats bounds as a sheet-qualified rectangle', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Quarter 1');
    expect(getRangeAddress(ws, bounds)).toBe("'Quarter 1'!B2:C3");
  });
});
