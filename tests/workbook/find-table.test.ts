// Tests for workbook-wide findTable.

import { describe, expect, it } from 'vitest';
import {
  addWorksheet,
  createWorkbook,
  findTable,
} from '../../src/workbook/workbook.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import { writeRange } from '../../src/worksheet/worksheet.js';

describe('findTable', () => {
  it('locates a table by displayName across multiple sheets', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    writeRange(a, 'A1', [['c1', 'c2']]);
    writeRange(b, 'A1', [['c1', 'c2']]);
    addExcelTable(wb, a, { name: 'TblA', ref: 'A1:B2', columns: ['c1', 'c2'] });
    addExcelTable(wb, b, { name: 'TblB', ref: 'A1:B2', columns: ['c1', 'c2'] });
    const hit = findTable(wb, 'TblB');
    expect(hit?.sheet.title).toBe('B');
    expect(hit?.table.name).toBe('TblB');
  });

  it('returns undefined when the table is not registered', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    writeRange(a, 'A1', [['c1', 'c2']]);
    addExcelTable(wb, a, { name: 'Tbl', ref: 'A1:B2', columns: ['c1', 'c2'] });
    expect(findTable(wb, 'Missing')).toBeUndefined();
  });

  it('empty workbook → undefined', () => {
    const wb = createWorkbook();
    expect(findTable(wb, 'Anything')).toBeUndefined();
  });
});
