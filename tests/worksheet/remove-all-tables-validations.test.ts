// Tests for removeAllTables / removeAllDataValidations.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { addListValidation } from '../../src/worksheet/data-validations.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import {
  listDataValidations,
  listTables,
  removeAllDataValidations,
  removeAllTables,
  writeRange,
} from '../../src/worksheet/worksheet.js';

describe('removeAllTables', () => {
  it('drops every table and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    writeRange(ws, 'A1', [['c1', 'c2']]);
    writeRange(ws, 'D1', [['c1', 'c2']]);
    addExcelTable(wb, ws, { name: 'T1', ref: 'A1:B2', columns: ['c1', 'c2'] });
    addExcelTable(wb, ws, { name: 'T2', ref: 'D1:E2', columns: ['c1', 'c2'] });
    expect(removeAllTables(ws)).toBe(2);
    expect(listTables(ws)).toEqual([]);
  });

  it('returns 0 when no tables exist', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(removeAllTables(ws)).toBe(0);
  });
});

describe('removeAllDataValidations', () => {
  it('drops every validation and returns the count', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addListValidation(ws, 'A1:A10', ['Yes', 'No']);
    addListValidation(ws, 'B1:B10', ['1', '2', '3']);
    expect(removeAllDataValidations(ws)).toBe(2);
    expect(listDataValidations(ws)).toEqual([]);
  });

  it('returns 0 when no validations exist', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(removeAllDataValidations(ws)).toBe(0);
  });
});
