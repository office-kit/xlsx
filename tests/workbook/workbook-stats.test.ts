// Tests for getWorkbookStats summary helper.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell.js';
import { addDefinedName } from '../../src/workbook/defined-names.js';
import { addChartsheet, addWorksheet, createWorkbook, getWorkbookStats } from '../../src/workbook/workbook.js';
import { setCustomStringProperty } from '../../src/packaging/custom.js';
import { addUrlHyperlink } from '../../src/worksheet/hyperlinks.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import { ensureCell, mergeCells, setCell, setComment, writeRange } from '../../src/worksheet/worksheet.js';

describe('getWorkbookStats', () => {
  it('empty workbook → zero everything', () => {
    const wb = createWorkbook();
    expect(getWorkbookStats(wb)).toEqual({
      worksheetCount: 0,
      chartsheetCount: 0,
      cellCount: 0,
      formulaCount: 0,
      commentCount: 0,
      hyperlinkCount: 0,
      mergedRangeCount: 0,
      tableCount: 0,
      definedNameCount: 0,
      customPropertyCount: 0,
    });
  });

  it('counts populated cells + formulas + comments + hyperlinks per worksheet', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'plain');
    const fc = ensureCell(a, 1, 2);
    setFormula(fc, 'A1+1');
    setCell(a, 2, 1, 42);
    setComment(a, { ref: 'A1', author: 'Alice', text: 'note' });
    addUrlHyperlink(a, 'A2', 'https://example.com');
    mergeCells(a, 'C1:D1');

    const b = addWorksheet(wb, 'B');
    setCell(b, 1, 1, 'b1');

    const stats = getWorkbookStats(wb);
    expect(stats.worksheetCount).toBe(2);
    expect(stats.cellCount).toBe(4);
    expect(stats.formulaCount).toBe(1);
    expect(stats.commentCount).toBe(1);
    expect(stats.hyperlinkCount).toBe(1);
    expect(stats.mergedRangeCount).toBe(1);
  });

  it('separates chartsheets from worksheets', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    addChartsheet(wb, 'Chart1');
    const stats = getWorkbookStats(wb);
    expect(stats.worksheetCount).toBe(1);
    expect(stats.chartsheetCount).toBe(1);
  });

  it('counts tables + definedNames + customProperties', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    setCell(a, 1, 1, 'h');
    writeRange(a, 'A1', [['c1', 'c2']]);
    addExcelTable(wb, a, { name: 'Tbl', ref: 'A1:B2', columns: ['c1', 'c2'] });
    addDefinedName(wb, { name: 'Wb', value: '$A$1' });
    addDefinedName(wb, { name: 'Sheet', value: '$A$1', scope: 0 });
    setCustomStringProperty(wb, 'project', 'Apollo');
    const stats = getWorkbookStats(wb);
    expect(stats.tableCount).toBe(1);
    expect(stats.definedNameCount).toBe(2);
    expect(stats.customPropertyCount).toBe(1);
  });
});
