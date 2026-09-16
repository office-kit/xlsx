// Tests for the data-validation builder helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  addCustomValidation,
  addDateValidation,
  addListValidation,
  addNumberValidation,
} from '../../src/worksheet/data-validations.js';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const expectSheet = (
  ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined,
): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('addListValidation', () => {
  it('inline list values become a quoted comma-separated formula1', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    const dv = addListValidation(ws, 'A1:A10', ['Red', 'Green', 'Blue']);
    expect(dv.type).toBe('list');
    expect(dv.formula1).toBe('"Red,Green,Blue"');
    expect(dv.allowBlank).toBe(true);
    expect(ws.dataValidations.length).toBe(1);
  });

  it('normalises a reference-string formula to the stored form', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    const dv = addListValidation(ws, 'A1:A10', '=Sheet2!$A$1:$A$10');
    expect(dv.formula1).toBe('Sheet2!$A$1:$A$10');
  });

  it('passes through prompt + error metadata', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    const dv = addListValidation(ws, 'A1:A10', ['Yes', 'No'], {
      prompt: 'Pick one',
      promptTitle: 'Choice',
      error: 'Must be Yes or No',
      errorStyle: 'stop',
    });
    expect(dv.prompt).toBe('Pick one');
    expect(dv.errorStyle).toBe('stop');
    expect(dv.showInputMessage).toBe(true);
    expect(dv.showErrorMessage).toBe(true);
  });
});

describe('addNumberValidation', () => {
  it('between(min, max) defaults to operator="between"', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    const dv = addNumberValidation(ws, 'B1:B5', { min: 0, max: 100 });
    expect(dv.type).toBe('whole');
    expect(dv.operator).toBe('between');
    expect(dv.formula1).toBe('0');
    expect(dv.formula2).toBe('100');
  });

  it('min only defaults to greaterThanOrEqual', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    const dv = addNumberValidation(ws, 'B1:B5', { min: 0 });
    expect(dv.operator).toBe('greaterThanOrEqual');
    expect(dv.formula2).toBeUndefined();
  });

  it('kind:"decimal" picks the decimal type', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    const dv = addNumberValidation(ws, 'B1:B5', { min: 0.1, max: 9.9, kind: 'decimal' });
    expect(dv.type).toBe('decimal');
  });
});

describe('addDateValidation + addCustomValidation', () => {
  it('addDateValidation accepts Excel serial dates', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    const dv = addDateValidation(ws, 'C1:C5', { min: 45000, max: 45100 });
    expect(dv.type).toBe('date');
    expect(dv.operator).toBe('between');
    expect(dv.formula1).toBe('45000');
    expect(dv.formula2).toBe('45100');
  });

  it('addCustomValidation normalises the formula to the stored form', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    const dv = addCustomValidation(ws, 'A1:A100', '=ISNUMBER(A1)');
    expect(dv.type).toBe('custom');
    expect(dv.formula1).toBe('ISNUMBER(A1)');
  });
});

describe('builders survive a save → load round-trip', () => {
  it('list + number + custom all preserve their attrs', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'V');
    setCell(ws, 1, 1, 1);
    addListValidation(ws, 'A1:A10', ['Open', 'Closed']);
    addNumberValidation(ws, 'B1:B10', { min: 0, max: 100 }, { errorStyle: 'warning' });
    addCustomValidation(ws, 'C1:C5', '=ISNUMBER(C1)');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.dataValidations.length).toBe(3);
    expect(ws2.dataValidations[0]?.type).toBe('list');
    expect(ws2.dataValidations[0]?.formula1).toBe('"Open,Closed"');
    expect(ws2.dataValidations[1]?.type).toBe('whole');
    expect(ws2.dataValidations[1]?.errorStyle).toBe('warning');
    expect(ws2.dataValidations[2]?.type).toBe('custom');
    expect(ws2.dataValidations[2]?.formula1).toBe('ISNUMBER(C1)');
  });
});