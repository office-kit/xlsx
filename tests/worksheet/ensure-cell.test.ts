// Tests for ensureCell and the mandatory-value contract on setCell.
//
// The pair exists because a styling / formula pass that walks already-populated
// rows must be able to reach a cell without wiping it.

import { describe, expect, it } from 'vitest';
import { getFormulaText, makeFormula, setFormula } from '../../src/cell/cell.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { ensureCell, getCell, setCell } from '../../src/worksheet/worksheet.js';

describe('ensureCell', () => {
  it('allocates a blank cell at an unpopulated coordinate', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = ensureCell(ws, 3, 2);
    expect(c.row).toBe(3);
    expect(c.col).toBe(2);
    expect(c.value).toBeNull();
    expect(c.styleId).toBe(0);
    expect(getCell(ws, 3, 2)).toBe(c);
  });

  it('returns the existing cell with its value and styleId intact', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'keep me', 4);
    const c = ensureCell(ws, 1, 1);
    expect(c.value).toBe('keep me');
    expect(c.styleId).toBe(4);
  });

  it('leaves formulas alone where setCell would blank them', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setFormula(ensureCell(ws, 1, 1), 'SUM(B1:D1)');
    expect(getFormulaText(ensureCell(ws, 1, 1))).toBe('SUM(B1:D1)');
    setCell(ws, 1, 1, null);
    expect(getCell(ws, 1, 1)?.value).toBeNull();
  });

  it('advances the append cursor so appendRow does not land on it', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    ensureCell(ws, 5, 1);
    expect(ws._appendRowCursor).toBe(5);
  });
});

describe('setCell with an explicit value', () => {
  it('replaces the value and keeps the styleId unless overridden', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'first', 3);
    setCell(ws, 1, 1, 'second');
    expect(getCell(ws, 1, 1)?.value).toBe('second');
    expect(getCell(ws, 1, 1)?.styleId).toBe(3);
    setCell(ws, 1, 1, 'third', 7);
    expect(getCell(ws, 1, 1)?.styleId).toBe(7);
  });

  it('places a formula and a style in one write via makeFormula', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 2, 1, makeFormula('SUM(A1:A1)', { cachedValue: 5 }), 2);
    expect(getFormulaText(c)).toBe('SUM(A1:A1)');
    expect(c.styleId).toBe(2);
  });
});
