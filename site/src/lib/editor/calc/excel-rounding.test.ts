// Results read from Excel for Mac (Microsoft 365). Excel rounds a result that
// only cancels to floating-point noise down to 0, but only in a formula's last
// addition / subtraction and in SUM's last addition.

import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { expect, test } from 'vitest';
import { CalcEngine } from './index.ts';

const engine = (() => {
  const wb = createWorkbook();
  addWorksheet(wb, 'Sheet1');
  return new CalcEngine(wb);
})();

test.each([
  ['0.5-0.4-0.1', 0],
  ['0.1+0.2-0.3', 0],
  ['1+1E-16-1', 0],
  ['1E+20+1-1E+20', 0],
  ['SUM(0.1,0.2,-0.3)', 0],
  ['SUM(0.3,-0.1,-0.2)', 0],
  ['SUM(0.1,0.2,-0.3)*1', 0],
  ['AVERAGE(0.1,0.2,-0.3)', 0],
  // Not the last operation, or written in parentheses: the noise stays.
  ['1*(0.5-0.4-0.1)', 0.5 - 0.4 - 0.1],
  ['(0.5-0.4-0.1)', 0.5 - 0.4 - 0.1],
  ['0.3-0.1-0.2+0', 0.3 - 0.1 - 0.2],
  ['IF(TRUE,0.5-0.4-0.1)', 0.5 - 0.4 - 0.1],
  ['SUM(0.1,0.2,-0.3,0)', 0.1 + 0.2 - 0.3],
  // Just past the 2^-48 tolerance.
  ['10.1-10-0.1', 10.1 - 10 - 0.1],
])('%s', (formula, expected) => {
  expect(engine.evaluate(formula, 'Sheet1', 1, 1)).toBe(expected);
});

test.each([
  ['VALUE("1 1/2")', 1.5],
  ['"1 1/2"+0', 1.5],
  ['VALUE("-1 1/2")', -1.5],
  ['VALUE("0 3/4")', 0.75],
])('mixed number %s', (formula, expected) => {
  expect(engine.evaluate(formula, 'Sheet1', 1, 1)).toBe(expected);
});

test('a serial past 12/31/9999 is not a date', () => {
  expect(engine.evaluate('YEAR(2958465.5)', 'Sheet1', 1, 1)).toBe(9999);
  expect(engine.evaluate('YEAR(2958466)', 'Sheet1', 1, 1)).toEqual({ kind: 'error', code: '#NUM!' });
});
