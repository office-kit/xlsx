// Tests for getFormulaText / getCachedFormulaValue introspection helpers.

import { describe, expect, it } from 'vitest';
import {
  getCachedFormulaValue,
  getFormulaText,
  setArrayFormula,
  setFormula,
  setSharedFormula,
} from '../../src/cell/cell.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { ensureCell, setCell } from '../../src/worksheet/worksheet.js';

describe('getFormulaText', () => {
  it('returns the formula string for a normal formula cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = ensureCell(ws, 1, 1);
    setFormula(c, 'A2+B2');
    expect(getFormulaText(c)).toBe('A2+B2');
  });

  it('returns the formula for an array formula', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = ensureCell(ws, 1, 1);
    setArrayFormula(c, 'A1:A3', 'TRANSPOSE(B1:D1)');
    expect(getFormulaText(c)).toBe('TRANSPOSE(B1:D1)');
  });

  it('returns the empty string for a shared follower (no formula text)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = ensureCell(ws, 2, 1);
    setSharedFormula(c, 0); // follower — no formula text, just si index
    expect(getFormulaText(c)).toBe('');
  });

  it('returns undefined for non-formula cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const numeric = setCell(ws, 1, 1, 42);
    const text = setCell(ws, 2, 1, 'hi');
    const empty = setCell(ws, 3, 1, null);
    expect(getFormulaText(numeric)).toBeUndefined();
    expect(getFormulaText(text)).toBeUndefined();
    expect(getFormulaText(empty)).toBeUndefined();
  });
});

describe('getCachedFormulaValue', () => {
  it('returns the cachedValue when set on a formula', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = ensureCell(ws, 1, 1);
    setFormula(c, 'A2+B2', { cachedValue: 7 });
    expect(getCachedFormulaValue(c)).toBe(7);
  });

  it('returns string + boolean cached values verbatim', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const sCell = ensureCell(ws, 1, 1);
    setFormula(sCell, '"ok"', { cachedValue: 'ok' });
    const bCell = ensureCell(ws, 2, 1);
    setFormula(bCell, 'TRUE()', { cachedValue: true });
    expect(getCachedFormulaValue(sCell)).toBe('ok');
    expect(getCachedFormulaValue(bCell)).toBe(true);
  });

  it('returns undefined when cachedValue is omitted', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = ensureCell(ws, 1, 1);
    setFormula(c, 'A2+B2');
    expect(getCachedFormulaValue(c)).toBeUndefined();
  });

  it('returns undefined for non-formula cells', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const c = setCell(ws, 1, 1, 42);
    expect(getCachedFormulaValue(c)).toBeUndefined();
  });
});
