// Tests for getAllTables / getAllDataValidations aggregators.

import { describe, expect, it } from 'vitest';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getAllDataValidations,
  getAllTables,
} from '../../src/workbook/workbook.js';
import { addListValidation } from '../../src/worksheet/data-validations.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import { writeRange } from '../../src/worksheet/worksheet.js';

describe('getAllTables', () => {
  it('aggregates tables across every worksheet in tab-strip order', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    writeRange(a, 'A1', [['c1', 'c2']]);
    writeRange(a, 'D1', [['c1', 'c2']]);
    writeRange(b, 'A1', [['c1', 'c2']]);
    addExcelTable(wb, a, { name: 'TblA', ref: 'A1:B2', columns: ['c1', 'c2'] });
    addExcelTable(wb, b, { name: 'TblB', ref: 'A1:B2', columns: ['c1', 'c2'] });
    addExcelTable(wb, a, { name: 'TblA2', ref: 'D1:E2', columns: ['c1', 'c2'] });
    const out = getAllTables(wb).map(({ sheet, table }) => `${sheet.title}:${table.name}`);
    expect(out).toEqual(['A:TblA', 'A:TblA2', 'B:TblB']);
  });

  it('skips chartsheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    addChartsheet(wb, 'Chart');
    expect(getAllTables(wb)).toEqual([]);
  });
});

describe('getAllDataValidations', () => {
  it('aggregates validations across every worksheet', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    addListValidation(a, 'A1:A10', ['Yes', 'No']);
    addListValidation(b, 'B1:B10', ['1', '2', '3']);
    const out = getAllDataValidations(wb);
    expect(out.length).toBe(2);
    expect(out[0]?.sheet.title).toBe('A');
    expect(out[1]?.sheet.title).toBe('B');
  });

  it('empty workbook → empty array', () => {
    const wb = createWorkbook();
    expect(getAllDataValidations(wb)).toEqual([]);
  });
});
