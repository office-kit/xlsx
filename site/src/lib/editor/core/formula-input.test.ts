import { describe, expect, test } from 'vitest';
import { getCell } from '@office-kit/xlsx/worksheet';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';
import { parseInput } from './input.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

function cellValue(ctl: EditorController, row: number, col: number): unknown {
  const v = getCell(ctl.doc.ws, row, col)?.value;
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.cachedValue : v;
}

function formula(ctl: EditorController, row: number, col: number): string | undefined {
  const v = getCell(ctl.doc.ws, row, col)?.value;
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.formula : undefined;
}

describe('formula entry, compared with Excel', () => {
  test('a LAMBDA held in a LET variable can be called', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '=LET(f,LAMBDA(a,a*2),f(4))');
    type(ctl, 2, 1, '=LET(sq,LAMBDA(n,n*n),sq(3)+1)');
    expect(cellValue(ctl, 1, 1)).toBe(8);
    expect(cellValue(ctl, 2, 1)).toBe(10);
  });

  test('known function names and references are upper-cased; a misspelt function is left as typed', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '=sum(a2:b3)+xlookup(1,c1:c2,d1:d2)');
    type(ctl, 2, 1, '=sumx(a1)');
    expect(formula(ctl, 1, 1)).toBe('SUM(A2:B3)+_xlfn.XLOOKUP(1,C1:C2,D1:D2)');
    expect(formula(ctl, 2, 1)).toBe('sumx(A1)');
  });

  test('a leading + or - that is not a plain number starts a formula', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '+1+2');
    type(ctl, 2, 1, '-3*2');
    type(ctl, 3, 1, '--5');
    type(ctl, 4, 1, '- item');
    expect(cellValue(ctl, 1, 1)).toBe(3);
    expect(cellValue(ctl, 2, 1)).toBe(-6);
    expect(cellValue(ctl, 3, 1)).toBe(5);
    expect(cellValue(ctl, 4, 1)).toBe('- item');
  });

  test('a General cell takes a date or time format from TODAY, DATE, NOW and TIME', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '=DATE(2024,1,5)');
    type(ctl, 2, 1, '=NOW()');
    type(ctl, 3, 1, '=TIME(13,30,0)');
    type(ctl, 4, 1, '=YEAR(TODAY())');
    const fmt = (row: number) => {
      ctl.selectCell({ row, col: 1 });
      return A.activeStyle(ctl).numFmt;
    };
    expect([fmt(1), fmt(2), fmt(3), fmt(4)]).toEqual(['m/d/yy', 'm/d/yy h:mm', 'h:mm AM/PM', 'General']);
  });

  test('typed January and February 1900 dates follow the 1900 leap-year bug', () => {
    expect(parseInput('1/1/1900').value).toBe(1);
    expect(parseInput('2/28/1900').value).toBe(59);
    expect(parseInput('2/29/1900').value).toBe(60);
    expect(parseInput('3/1/1900').value).toBe(61);
  });

  // Each expectation was checked by entering the text in Excel for Mac (en-US).
  test('typed dates, date-times, month-years, mixed fractions and long numbers parse like Excel', () => {
    const today = new Date(2024, 5, 1);
    expect(parseInput('1/31/2024', { today })).toEqual({ value: 45322, impliedFormat: 'm/d/yy' });
    expect(parseInput('2024-01-31', { today })).toEqual({ value: 45322, impliedFormat: 'm/d/yy' });
    expect(parseInput('1/31/2024 1:30 PM', { today })).toEqual({ value: 45322.5625, impliedFormat: 'm/d/yy h:mm' });
    expect(parseInput('Jan 2024', { today })).toEqual({ value: 45292, impliedFormat: 'mmm-yy' });
    expect(parseInput('1/2024', { today })).toEqual({ value: 45292, impliedFormat: 'mmm-yy' });
    expect(parseInput('-1 1/2', { today })).toEqual({ value: -1.5, impliedFormat: '# ?/?' });
    expect(parseInput('1234567890123456789', { today }).value).toBe(1234567890123450000);
    expect(parseInput('0.12345678901234567', { today }).value).toBe(0.123456789012345);
  });
});

