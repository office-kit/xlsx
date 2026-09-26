// Cost gate for the range walks that delete cells.
//
// A range can name far more coordinates than the sheet holds cells. `'A:A'` is
// 1_048_576 coordinates, `'A1:XFD1048576'` is seventeen billion, and Excel lets
// a caller merge either. `mergeCells` walked every coordinate of the rectangle
// with two Map lookups each, so merging a whole-column band on a sheet holding
// one cell cost time proportional to the band: `'A:J'` took 43 ms on an empty
// sheet here, which extrapolates to about 70 s for a whole-sheet merge.
//
// The walk now enumerates each axis whichever way is smaller, so the cost
// follows the cells rather than the rectangle. This measures a whole-column
// merge in absolute terms because the point is that it no longer scales with
// the rectangle at all; the ceiling is loose enough to survive a slow machine
// and would still catch the area walk coming back.
//
// Excluded from the default `pnpm test` run (see vitest.config.ts). Run
// explicitly:
//
//   pnpm test:perf
//   PERF_GATE=1 pnpm test:perf   # asserts the ceilings, as CI does

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { clearRange, mergeCells, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const PERF_GATE = process.env['PERF_GATE'] === '1';

/** Comfortably above a sparse walk, far below an area walk of these bands. */
const CEILING_MS = 50;

const sparseSheet = (): Worksheet => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  for (let r = 1; r <= 20; r++) setCell(ws, r, 1, r);
  return ws;
};

const time = (label: string, run: () => void): number => {
  const start = performance.now();
  run();
  const elapsed = performance.now() - start;
  // eslint-disable-next-line no-console
  console.log(`${label}: ${elapsed.toFixed(1)}ms`);
  return elapsed;
};

describe('range walks follow the cells, not the rectangle', () => {
  it('merges a whole-column band over a sparse sheet', () => {
    const ws = sparseSheet();
    const elapsed = time("mergeCells('A:J') over 20 cells", () => mergeCells(ws, 'A:J'));
    if (PERF_GATE) expect(elapsed).toBeLessThan(CEILING_MS);
  });

  it('merges the whole grid over a sparse sheet', () => {
    // Seventeen billion coordinates. An area walk never finishes this.
    const ws = sparseSheet();
    const elapsed = time("mergeCells('A1:XFD1048576') over 20 cells", () =>
      mergeCells(ws, 'A1:XFD1048576'),
    );
    if (PERF_GATE) expect(elapsed).toBeLessThan(CEILING_MS);
  });

  it('clears the whole grid over a sparse sheet', () => {
    const ws = sparseSheet();
    let removed = 0;
    const elapsed = time("clearRange('A1:XFD1048576') over 20 cells", () => {
      removed = clearRange(ws, 'A1:XFD1048576');
    });
    expect(removed).toBe(20);
    if (PERF_GATE) expect(elapsed).toBeLessThan(CEILING_MS);
  });

  it('stays cheap for a small range on a sheet with many rows', () => {
    // The other direction: enumerating the sparse store instead of the range
    // would make this scale with the sheet rather than with the two cells
    // being cleared.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    for (let r = 1; r <= 200_000; r++) setCell(ws, r, 1, r);
    const elapsed = time("clearRange('A1:B2') over 200_000 rows", () => clearRange(ws, 'A1:B2'));
    if (PERF_GATE) expect(elapsed).toBeLessThan(CEILING_MS);
  });
});
