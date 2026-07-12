// Scenario 30: classic single-value formulas. Output: 30-modern-formulas.xlsx
//
// The original scenario covered Excel 365 / 2021 modern formulas (LET,
// LAMBDA, FILTER, SORT, UNIQUE, SEQUENCE, XLOOKUP, BYROW). This scenario
// deliberately stays on formulas that resolve in every Excel since 2007
// (AVERAGE / VLOOKUP) for one reason only: version compatibility. Bare
// `LET(...)` resolves only on Excel 365 and `XLOOKUP(...)` only on Excel
// 2019+; older Excel shows #NAME?. AVERAGE / VLOOKUP keep the same
// demonstrative shape (single-value arithmetic + lookup) and evaluate
// everywhere.
//
// Note: the `_xlfn.` / `_xlfn._xlws.` *name* prefix is not what blocks
// this — that prefix is preserved verbatim on round-trip (see
// `handleFormula` in src/worksheet/reader.ts) and Excel accepts it in an
// ordinary cell with no metadata. The `cm="N"` + `xl/metadata.xml`
// coupling is a *separate* concern that only marks a spill as a resizing
// dynamic array; we don't emit that metadata yet, so a round-tripped
// spill survives as a CSE array (values correct, no auto-resize).
//
// What to verify in Excel:
// - E2 = AVERAGE(C2:C8) returns the average salary (84428.57).
// - F2 = VLOOKUP("Bob",A2:C8,3,FALSE) returns Bob's salary (88000).

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../../src/workbook/index';
import { setCell } from '../../../src/worksheet/index';
import { setFormula } from '../../../src/cell/cell';
import { writeWorkbook } from '../_helpers';

describe('e2e 30 — single-value formulas', () => {
  it('writes 30-modern-formulas.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Modern');

    // Source table A1:C8
    setCell(ws, 1, 1, 'Name');
    setCell(ws, 1, 2, 'Dept');
    setCell(ws, 1, 3, 'Salary');
    const rows: Array<[string, string, number]> = [
      ['Alice', 'Eng', 95000],
      ['Bob', 'Eng', 88000],
      ['Carol', 'Sales', 79000],
      ['Dan', 'Ops', 71000],
      ['Eve', 'Eng', 102000],
      ['Frank', 'Sales', 76000],
      ['Grace', 'Ops', 80000],
    ];
    rows.forEach((r, i) => r.forEach((v, c) => setCell(ws, i + 2, c + 1, v)));

    setCell(ws, 1, 5, 'AVERAGE');
    setFormula(setCell(ws, 2, 5, ''), 'AVERAGE(C2:C8)', { cachedValue: 84428.57 });

    setCell(ws, 1, 6, 'VLOOKUP');
    setFormula(setCell(ws, 2, 6, ''), 'VLOOKUP("Bob",A2:C8,3,FALSE)', { cachedValue: 88000 });

    const result = await writeWorkbook('30-modern-formulas.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
