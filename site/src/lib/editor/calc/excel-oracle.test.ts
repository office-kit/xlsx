// Checks the engine against values recorded from real Excel by
// excel-oracle.mjs. Each case was entered there as a single-cell array
// formula, so the expected value is the top-left of the formula's result.

import { type CellValue, type ExcelErrorCode, makeErrorValue } from '@office-kit/xlsx/cell';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { setCell } from '@office-kit/xlsx/worksheet';
import { describe, expect, test } from 'vitest';
import fixture from './excel-oracle.fixture.json' with { type: 'json' };
import { CalcEngine, type CalcScalar } from './index.ts';

type Expected = { type: 'number'; value: number } | { type: 'string'; value: string } | { type: 'boolean'; value: boolean } | { type: 'error'; value: string } | { type: 'blank' };
type Cell = [number, number, string | number | boolean | { error: ExcelErrorCode } | null];

const toModel = (v: Cell[2]): CellValue => (v !== null && typeof v === 'object' ? makeErrorValue(v.error) : v);

const engine = (() => {
  const wb = createWorkbook();
  const data = addWorksheet(wb, 'Data');
  for (const [r, c, v] of fixture.data as Cell[]) setCell(data, r, c, toModel(v));
  const db = addWorksheet(wb, 'Db');
  for (const [r, c, v] of fixture.db as Cell[]) setCell(db, r, c, toModel(v));
  addWorksheet(wb, 'Cases');
  const e = new CalcEngine(wb);
  e.recalculateAll();
  return e;
})();

// Excel caches the post-2018 errors as #VALUE! in the file (the real code
// lives in rich-value metadata), so that is all the fixture can hold for them;
// ERROR.TYPE(UNIQUE(...)) = 14 in the fixture confirms #CALC! is what Excel computed.
const FILE_ERROR: Readonly<Record<string, string>> = { '#CALC!': '#VALUE!', '#SPILL!': '#VALUE!' };

// The oracle Excel runs in a Japanese locale (dates read year-first, ¥ is the
// currency, the Mac code page backs CHAR/CODE) while the engine implements
// en-US; CELL("address") on another sheet also embeds the workbook's file name.
const ENVIRONMENT_DEPENDENT = new Set([
  '"1/15/2024"+0',
  'COUNTIF(Data!C1:C10,">3/1/2024")',
  'CHAR(128)',
  'CODE("€")',
  'DOLLAR(1234.567)',
  'DOLLAR(-1234.567,1)',
  'DOLLAR(1234.567,-2)',
  'DATEVALUE("1/15/2024")',
  'DATEVALUE("January 15 2024")',
  'DATEVALUE("1/1/29")',
  'DATEVALUE("1/1/30")',
  'CELL("address",Data!B7)',
]);

const actual = (v: CalcScalar): Expected => {
  if (v === null) return { type: 'blank' };
  if (typeof v === 'object') return { type: 'error', value: FILE_ERROR[v.code] ?? v.code };
  if (typeof v === 'number') return { type: 'number', value: v };
  if (typeof v === 'string') return { type: 'string', value: v };
  return { type: 'boolean', value: v };
};

// Closed-form results agree to the last bit or two; Excel's iterative
// solvers (RATE, IRR, XIRR) stop a little short of the root, which this allows.
const RELATIVE_TOLERANCE = 1e-8;

describe(`engine matches Excel ${fixture.excelVersion}`, () => {
  const cases = (fixture.cases as Array<{ formula: string; expected: Expected }>)
    .map((c, i) => ({ formula: c.formula, expected: c.expected, row: i + 1 }))
    .filter((c) => !ENVIRONMENT_DEPENDENT.has(c.formula));
  test.each(cases)('$formula', ({ formula, expected, row }) => {
    const got = actual(engine.evaluate(formula, 'Cases', row, 1));
    if (expected.type === 'number' && got.type === 'number') {
      const scale = Math.max(1, Math.abs(expected.value));
      expect(Math.abs(got.value - expected.value) / scale).toBeLessThanOrEqual(RELATIVE_TOLERANCE);
    } else {
      expect(got).toEqual(expected);
    }
  });
});
