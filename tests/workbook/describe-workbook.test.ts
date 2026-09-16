// Tests for describeWorkbook — workbook-level overview snapshot.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell.js';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  describeWorkbook,
  hideSheet,
} from '../../src/workbook/workbook.js';
import { ensureCell, setCell, mergeCells, addTable, setHyperlink, setComment } from '../../src/worksheet/worksheet.js';

describe('describeWorkbook', () => {
  it('returns all-zero counts + empty sheets list for an empty workbook', () => {
    const wb = createWorkbook();
    const o = describeWorkbook(wb);
    expect(o.worksheetCount).toBe(0);
    expect(o.chartsheetCount).toBe(0);
    expect(o.cellCount).toBe(0);
    expect(o.formulaCount).toBe(0);
    expect(o.sheets).toEqual([]);
    expect(o.cellsByKind.string).toBe(0);
  });

  it('aggregates a single populated worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 1, 1);
    const f = ensureCell(ws, 3, 1);
    setFormula(f, 'A1+1');
    mergeCells(ws, 'A1:B1');
    addTable(ws, { id: 1, displayName: 'T', ref: 'A1:A2', columns: [{ id: 1, name: 'a' }] });
    setHyperlink(ws, 'A1', { target: 'https://example.com' });
    setComment(ws, { ref: 'A2', text: 'note', author: 'me' });
    const o = describeWorkbook(wb);
    expect(o.worksheetCount).toBe(1);
    expect(o.cellCount).toBe(3);
    expect(o.formulaCount).toBe(1);
    expect(o.commentCount).toBe(1);
    expect(o.hyperlinkCount).toBe(1);
    expect(o.mergedRangeCount).toBe(1);
    expect(o.tableCount).toBe(1);
    expect(o.cellsByKind).toMatchObject({ string: 1, number: 1, formula: 1 });
    expect(o.sheets[0]).toMatchObject({
      title: 'A',
      kind: 'worksheet',
      state: 'visible',
      cellCount: 3,
      formulaCount: 1,
      tableCount: 1,
    });
  });

  it('handles a mix of worksheets + chartsheet, including hidden', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    addWorksheet(wb, 'B');
    addChartsheet(wb, 'Chart1');
    setCell(a, 1, 1, 'x');
    hideSheet(wb, 'B');
    const o = describeWorkbook(wb);
    expect(o.worksheetCount).toBe(2);
    expect(o.chartsheetCount).toBe(1);
    expect(o.sheets.length).toBe(3);
    expect(o.sheets.find((s) => s.title === 'A')?.cellCount).toBe(1);
    expect(o.sheets.find((s) => s.title === 'B')?.state).toBe('hidden');
    expect(o.sheets.find((s) => s.title === 'Chart1')?.kind).toBe('chartsheet');
    expect(o.sheets.find((s) => s.title === 'Chart1')?.cellCount).toBe(0);
  });

  it('preserves tab-strip order in sheets[]', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'first');
    addWorksheet(wb, 'second');
    addChartsheet(wb, 'middle-chart');
    addWorksheet(wb, 'third');
    expect(describeWorkbook(wb).sheets.map((s) => s.title)).toEqual([
      'first',
      'second',
      'middle-chart',
      'third',
    ]);
  });
});
