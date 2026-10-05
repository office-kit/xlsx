// A sweep of the whole built-in format catalogue, so a change to the renderer
// cannot quietly move any of the 36 codes Excel ships with. One numeric value
// and one date serial exercise both halves of the catalogue.

import { describe, expect, it } from 'vitest';
import { BUILTIN_FORMATS, getCellDisplayText, setCellNumberFormat } from '../../src/styles/index.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/index.js';
import { setCell } from '../../src/worksheet/index.js';

const display = (code: string, value: number): string => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Sheet1');
  const cell = setCell(ws, 1, 1, value);
  setCellNumberFormat(wb, cell, code);
  return getCellDisplayText(wb, cell);
};

/** A plain number, and 2024-03-14 12:00 as a Windows-epoch serial. */
const NUMBER = 1234.5;
const SERIAL = 45_365.5;

const EXPECTED: Record<number, [string, string]> = {
  0: ['1234.5', '45365.5'],
  1: ['1235', '45366'],
  2: ['1234.50', '45365.50'],
  3: ['1,235', '45,366'],
  4: ['1,234.50', '45,365.50'],
  5: ['$1,235 ', '$45,366 '],
  6: ['$1,235 ', '$45,366 '],
  7: ['$1,234.50 ', '$45,365.50 '],
  8: ['$1,234.50 ', '$45,365.50 '],
  9: ['123450%', '4536550%'],
  10: ['123450.00%', '4536550.00%'],
  11: ['1.23E+03', '4.54E+04'],
  12: ['1234 1/2', '45365 1/2'],
  13: ['1234  1/2 ', '45365  1/2 '],
  14: ['05-18-03', '03-14-24'],
  15: ['18-May-03', '14-Mar-24'],
  16: ['18-May', '14-Mar'],
  17: ['May-03', 'Mar-24'],
  18: ['12:00 PM', '12:00 PM'],
  19: ['12:00:00 PM', '12:00:00 PM'],
  20: ['12:00', '12:00'],
  21: ['12:00:00', '12:00:00'],
  22: ['5/18/03 12:00', '3/14/24 12:00'],
  37: ['1,235 ', '45,366 '],
  38: ['1,235 ', '45,366 '],
  39: ['1,234.50 ', '45,365.50 '],
  40: ['1,234.50 ', '45,365.50 '],
  41: [' 1,235 ', ' 45,366 '],
  42: [' $1,235 ', ' $45,366 '],
  43: [' 1,234.50 ', ' 45,365.50 '],
  44: [' $1,234.50 ', ' $45,365.50 '],
  45: ['00:00', '00:00'],
  46: ['29628:00:00', '1088772:00:00'],
  47: ['00:00.0', '00:00.0'],
  48: ['1.2E+3', '45.4E+3'],
  49: ['1234.5', '45365.5'],
};

describe('getCellDisplayText over the built-in format catalogue', () => {
  it('covers every id the catalogue defines', () => {
    expect(Object.keys(EXPECTED).sort()).toEqual(Object.keys(BUILTIN_FORMATS).sort());
  });

  it.each(Object.entries(BUILTIN_FORMATS))('numFmtId %s (%s)', (id, code) => {
    const expected = EXPECTED[Number(id)];
    expect(expected).toBeDefined();
    if (expected === undefined) return;
    expect([display(code, NUMBER), display(code, SERIAL)]).toEqual(expected);
  });

  it('renders numFmtId 44 like Excel for negatives and zero too', () => {
    const code = BUILTIN_FORMATS[44] ?? '';
    expect(display(code, -NUMBER)).toBe(' $(1,234.50)');
    expect(display(code, 0)).toBe(' $-   ');
  });
});
