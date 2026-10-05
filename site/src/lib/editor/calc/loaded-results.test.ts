import { type CellValue, makeArrayFormula, makeFormula } from '@office-kit/xlsx/cell';
import { addDefinedName, addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { getCell, setCell, type Worksheet } from '@office-kit/xlsx/worksheet';
import { describe, expect, test } from 'vitest';
import { CalcEngine } from './index.ts';

const shown = (ws: Worksheet, row: number, col: number): CellValue | undefined => {
  const v = getCell(ws, row, col)?.value;
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.cachedValue : v;
};

const setup = (cells: ReadonlyArray<[number, number, CellValue]>) => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  for (const [r, c, v] of cells) setCell(ws, r, c, v);
  const engine = new CalcEngine(wb);
  engine.recalculateAll();
  return { wb, ws, engine };
};

describe('legacy (Ctrl+Shift+Enter) arrays keep their size', () => {
  test('a scalar result fills the whole area', () => {
    const { ws } = setup([
      [1, 1, 1],
      [2, 1, 2],
      [3, 1, 3],
      [1, 3, makeArrayFormula('C1:C3', 'SUM(A1:A3)', { cachedValue: 6 })],
      [2, 3, 6],
      [3, 3, 6],
    ]);
    expect([1, 2, 3].map((r) => shown(ws, r, 3))).toEqual([6, 6, 6]);
    expect(getCell(ws, 1, 3)?.value).toMatchObject({ t: 'array', ref: 'C1:C3' });
  });

  test('a larger result is clipped and a smaller one pads with #N/A', () => {
    const { ws } = setup([
      [1, 1, 1],
      [2, 1, 2],
      [3, 1, 3],
      [1, 4, makeArrayFormula('D1:D2', 'A1:A3*2')],
      [1, 5, makeArrayFormula('E1:F4', 'A1:A3*10')],
    ]);
    expect([1, 2].map((r) => shown(ws, r, 4))).toEqual([2, 4]);
    expect(shown(ws, 3, 4)).toBeUndefined();
    expect(getCell(ws, 1, 4)?.value).toMatchObject({ ref: 'D1:D2' });
    // One column repeats across; the fourth row is past the result.
    expect([1, 2, 3].map((r) => [shown(ws, r, 5), shown(ws, r, 6)])).toEqual([
      [10, 10],
      [20, 20],
      [30, 30],
    ]);
    expect(shown(ws, 4, 5)).toEqual({ kind: 'error', code: '#N/A' });
  });
});

describe('results this engine cannot compute keep what Excel saved', () => {
  test('macro functions, external links and CELL("filename")', () => {
    const { ws } = setup([
      [1, 1, makeFormula('MyUdf(2)', { cachedValue: 42 })],
      [2, 1, makeFormula('[1]Sheet1!$A$5', { cachedValue: 'linked' })],
      [3, 1, makeFormula('CELL("filename")', { cachedValue: '/tmp/[a.xlsx]S' })],
      [4, 1, makeFormula('A1+1', { cachedValue: 0 })],
      [5, 1, makeFormula('_xll.ADDIN(1)', { cachedValue: 7 })],
    ]);
    expect([1, 2, 3, 5].map((r) => shown(ws, r, 1))).toEqual([42, 'linked', '/tmp/[a.xlsx]S', 7]);
    expect(shown(ws, 4, 1)).toBe(43);
  });

  test('a newly typed unknown function still shows #NAME?', () => {
    const { ws } = setup([[1, 1, makeFormula('MyUdf(2)')]]);
    expect(getCell(ws, 1, 1)?.value).toMatchObject({ cachedValue: '#NAME?', cachedValueType: 'error' });
  });

  test('LET and defined-name lambdas are still computed', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    addDefinedName(wb, { name: 'Twice', value: 'LAMBDA(x,x*2)' });
    setCell(ws, 1, 1, makeFormula('LET(f,LAMBDA(x,x+1),f(2))', { cachedValue: 0 }));
    setCell(ws, 2, 1, makeFormula('Twice(5)', { cachedValue: 0 }));
    new CalcEngine(wb).recalculateAll();
    expect(shown(ws, 1, 1)).toBe(3);
    expect(shown(ws, 2, 1)).toBe(10);
  });
});
