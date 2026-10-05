import { type CellValue, makeFormula } from '@office-kit/xlsx/cell';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { addExcelTable, getCell, setCell } from '@office-kit/xlsx/worksheet';
import { expect, test } from 'vitest';
import { CalcEngine } from './index.ts';

test('a header ending in a space is its own column', () => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  const rows: CellValue[][] = [
    ['Amount ', 'Amount'],
    [5, 1],
    [6, 2],
  ];
  rows.forEach((r, i) => r.forEach((v, j) => setCell(ws, i + 1, j + 1, v)));
  addExcelTable(wb, ws, { name: 'T', ref: 'A1:B3', columns: ['Amount ', 'Amount'] });
  setCell(ws, 5, 1, makeFormula('SUM(T[Amount ])'));
  setCell(ws, 5, 2, makeFormula('SUM(T[Amount])'));
  setCell(ws, 5, 3, makeFormula('SUM(T[[#Data],[Amount ]])'));
  new CalcEngine(wb).recalculateAll();
  const shown = (c: number) => {
    const v = getCell(ws, 5, c)?.value;
    return v && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.cachedValue : v;
  };
  expect([shown(1), shown(2), shown(3)]).toEqual([11, 3, 11]);
});
