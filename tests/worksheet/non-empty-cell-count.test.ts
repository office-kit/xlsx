// Tests for getNonEmptyCellCount — null/empty exclusion and formula/rich-text opts.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { ensureCell, getNonEmptyCellCount, setCell } from '../../src/worksheet/worksheet.js';

describe('getNonEmptyCellCount', () => {
  it('counts every cell whose value is non-null', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, 1);
    setCell(ws, 1, 3, true);
    expect(getNonEmptyCellCount(ws)).toBe(3);
  });

  it('skips cells with value === null (but they remain in the row map)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 1, 2, null);
    setCell(ws, 1, 3, 'c');
    expect(getNonEmptyCellCount(ws)).toBe(2);
  });

  it('counts the empty string "" as a real value (it is not null)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, '');
    setCell(ws, 1, 2, 'b');
    expect(getNonEmptyCellCount(ws)).toBe(2);
  });

  it('opts.includeFormulas: false drops FormulaValue cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'plain');
    const c = ensureCell(ws, 1, 2);
    setFormula(c, 'A1');
    expect(getNonEmptyCellCount(ws)).toBe(2);
    expect(getNonEmptyCellCount(ws, { includeFormulas: false })).toBe(1);
  });

  it('opts.includeRichText: false drops rich-text cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'plain');
    setCell(ws, 1, 2, { kind: 'rich-text', runs: [] });
    expect(getNonEmptyCellCount(ws)).toBe(2);
    expect(getNonEmptyCellCount(ws, { includeRichText: false })).toBe(1);
  });
});
