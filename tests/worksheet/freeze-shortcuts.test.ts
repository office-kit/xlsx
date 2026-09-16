// Tests for the freezeFirst* shortcut helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  freezeFirstColumn,
  freezeFirstRow,
  freezeFirstRowAndColumn,
  getFreezePanes,
  setFreezePanes,
} from '../../src/worksheet/worksheet.js';

describe('freeze shortcuts', () => {
  it('freezeFirstRow: freezes row 1 (top-left of unfrozen pane = A2)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezeFirstRow(ws);
    expect(getFreezePanes(ws)).toBe('A2');
  });

  it('freezeFirstColumn: freezes column A (top-left of unfrozen pane = B1)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezeFirstColumn(ws);
    expect(getFreezePanes(ws)).toBe('B1');
  });

  it('freezeFirstRowAndColumn: freezes both (top-left of unfrozen pane = B2)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    freezeFirstRowAndColumn(ws);
    expect(getFreezePanes(ws)).toBe('B2');
  });

  it('overrides a pre-existing freeze without composing on top', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setFreezePanes(ws, { rows: 5, cols: 3 }); // freeze 5 rows + 3 cols → top-left D6
    expect(getFreezePanes(ws)).toBe('D6');
    freezeFirstRow(ws);
    expect(getFreezePanes(ws)).toBe('A2');
  });
});
