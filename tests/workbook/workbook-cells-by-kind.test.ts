// Tests for getWorkbookCellsByKind — workbook-wide kind histogram.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell.js';
import {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  getWorkbookCellsByKind,
} from '../../src/workbook/workbook.js';
import { ensureCell, setCell } from '../../src/worksheet/worksheet.js';

const zeros = {
  null: 0,
  string: 0,
  number: 0,
  boolean: 0,
  date: 0,
  duration: 0,
  error: 0,
  'rich-text': 0,
  formula: 0,
};

describe('getWorkbookCellsByKind', () => {
  it('returns all-zero buckets for an empty workbook', () => {
    const wb = createWorkbook();
    expect(getWorkbookCellsByKind(wb)).toEqual(zeros);
  });

  it('aggregates a single sheet correctly', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, true);
    expect(getWorkbookCellsByKind(wb)).toEqual({ ...zeros, string: 1, number: 1, boolean: 1 });
  });

  it('sums across multiple sheets', () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'A');
    const b = addWorksheet(wb, 'B');
    setCell(a, 1, 1, 'x');
    setCell(a, 2, 1, 'y');
    const f = ensureCell(b, 1, 1);
    setFormula(f, 'A1+1');
    setCell(b, 2, 1, 42);
    expect(getWorkbookCellsByKind(wb)).toEqual({ ...zeros, string: 2, number: 1, formula: 1 });
  });

  it('skips chartsheets (they hold no cells)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'plain');
    addChartsheet(wb, 'Chart1');
    expect(getWorkbookCellsByKind(wb)).toEqual({ ...zeros, string: 1 });
  });
});
