// Tests for the conditional-formatting builder helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  addAverageRule,
  addCellIsRule,
  addDuplicateValuesRule,
  addFormulaRule,
  addTextRule,
  addTopNRule,
} from '../../src/worksheet/conditional-formatting.js';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('addCellIsRule', () => {
  it('between two formulas builds a 2-formula cellIs rule', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    const r = addCellIsRule(ws, 'A1:A10', { operator: 'between', formula1: '0', formula2: '100', dxfId: 0 });
    expect(r.type).toBe('cellIs');
    expect(r.operator).toBe('between');
    expect(r.formulas).toEqual(['0', '100']);
    expect(r.dxfId).toBe(0);
    expect(r.priority).toBe(1);
  });

  it('single-formula form omits formula2', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    const r = addCellIsRule(ws, 'A1:A10', { operator: 'greaterThan', formula1: '15', dxfId: 1 });
    expect(r.formulas).toEqual(['15']);
  });

  it('priority auto-increments per worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    const r1 = addCellIsRule(ws, 'A1:A10', { operator: 'lessThan', formula1: '0' });
    const r2 = addFormulaRule(ws, 'B1:B10', { formula: '=ISNUMBER(B1)' });
    const r3 = addTopNRule(ws, 'C1:C10', { rank: 5 });
    expect(r1.priority).toBe(1);
    expect(r2.priority).toBe(2);
    expect(r3.priority).toBe(3);
  });
});

describe('addTopNRule', () => {
  it('defaults are sane (rank=10 implicit, top-N)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'T');
    const r = addTopNRule(ws, 'A1:A100', { dxfId: 0 });
    expect(r.type).toBe('top10');
    expect(r.bottom).toBeUndefined();
    expect(r.percent).toBeUndefined();
  });

  it('bottom: true + percent: true produces a bottom-percentile rule', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'T');
    const r = addTopNRule(ws, 'A1:A100', { rank: 10, bottom: true, percent: true });
    expect(r.bottom).toBe(true);
    expect(r.percent).toBe(true);
    expect(r.rank).toBe(10);
  });
});

describe('addAverageRule', () => {
  it('above-average rule defaults', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const r = addAverageRule(ws, 'A1:A100', { aboveAverage: true, dxfId: 2 });
    expect(r.type).toBe('aboveAverage');
    expect(r.aboveAverage).toBe(true);
  });

  it('±1 standard deviation form', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    const r = addAverageRule(ws, 'A1:A100', { aboveAverage: true, stdDev: 1 });
    expect(r.stdDev).toBe(1);
  });
});

describe('addDuplicateValuesRule + addFormulaRule + addTextRule', () => {
  it('duplicates default vs unique', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'D');
    const dup = addDuplicateValuesRule(ws, 'A1:A100');
    expect(dup.type).toBe('duplicateValues');
    const uniq = addDuplicateValuesRule(ws, 'B1:B100', { unique: true });
    expect(uniq.type).toBe('uniqueValues');
  });

  it('formula rule normalises the formula to the stored form', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    const r = addFormulaRule(ws, 'A1:A100', { formula: '=ISNUMBER(A1)*A1>0' });
    expect(r.type).toBe('expression');
    expect(r.formulas).toEqual(['ISNUMBER(A1)*A1>0']);
  });

  it('text rule maps the operator to the correct ECMA type token', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'T');
    expect(addTextRule(ws, 'A1:A10', { operator: 'containsText', text: 'foo' }).type).toBe('containsText');
    expect(addTextRule(ws, 'A1:A10', { operator: 'notContains', text: 'foo' }).type).toBe('notContainsText');
    expect(addTextRule(ws, 'A1:A10', { operator: 'beginsWith', text: 'foo' }).type).toBe('beginsWith');
    expect(addTextRule(ws, 'A1:A10', { operator: 'endsWith', text: 'foo' }).type).toBe('endsWith');
  });
});

describe('builders survive a save → load round-trip', () => {
  it('cellIs + topN + formula all preserve their attrs', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'CF');
    setCell(ws, 1, 1, 1);
    addCellIsRule(ws, 'A1:A10', { operator: 'greaterThan', formula1: '15', dxfId: 0 });
    addTopNRule(ws, 'B1:B10', { rank: 3, dxfId: 0 });
    addFormulaRule(ws, 'C1:C10', { formula: '=ISNUMBER(C1)' });

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.conditionalFormatting.length).toBe(3);
    expect(ws2.conditionalFormatting[0]?.rules[0]?.operator).toBe('greaterThan');
    expect(ws2.conditionalFormatting[1]?.rules[0]?.rank).toBe(3);
    expect(ws2.conditionalFormatting[2]?.rules[0]?.formulas?.[0]).toBe('ISNUMBER(C1)');
  });
});