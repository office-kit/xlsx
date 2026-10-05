// Guards the engine's complexity, not its constant factors: the budgets are
// several times what a laptop needs, so they only trip on an O(n²) regression
// (per-cell range expansion, rescanning the criteria range per SUMIF, ...).

import { makeFormula } from '@office-kit/xlsx/cell';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { getCell, setCell } from '@office-kit/xlsx/worksheet';
import { expect, test } from 'vitest';
import { CalcEngine } from './index.ts';

const time = (fn: () => void): number => {
  const start = performance.now();
  fn();
  return performance.now() - start;
};

test('a 50k-cell dependency chain', () => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  const n = 50_000;
  setCell(ws, 1, 1, 1);
  for (let r = 2; r <= n; r++) setCell(ws, r, 1, makeFormula(`A${r - 1}+1`));
  const engine = new CalcEngine(wb);
  const full = time(() => engine.recalculateAll());
  setCell(ws, 1, 1, 2);
  const incremental = time(() => engine.update([{ sheet: 'S', row: 1, col: 1 }]));
  expect(getCell(ws, n, 1)?.value).toMatchObject({ cachedValue: n + 1 });
  expect(full).toBeLessThan(5_000);
  expect(incremental).toBeLessThan(5_000);
});

test('10k SUMIFs over a 10k-row column', () => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  const n = 10_000;
  for (let r = 1; r <= n; r++) {
    setCell(ws, r, 1, `k${r % 100}`);
    setCell(ws, r, 2, r);
    setCell(ws, r, 3, makeFormula(`SUMIF($A$1:$A$${n},A${r},$B$1:$B$${n})`));
  }
  const engine = new CalcEngine(wb);
  const full = time(() => engine.recalculateAll());
  // One value edit dirties every SUMIF reading column B.
  setCell(ws, 1, 2, 0);
  const incremental = time(() => engine.update([{ sheet: 'S', row: 1, col: 2 }]));
  // Key k0 sits on rows 100, 200, ..., 10000.
  expect(getCell(ws, 100, 3)?.value).toMatchObject({ cachedValue: 505_000 });
  expect(full).toBeLessThan(5_000);
  expect(incremental).toBeLessThan(5_000);
});
