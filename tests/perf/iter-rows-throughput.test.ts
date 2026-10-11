// Throughput gate for `loadWorkbookStream`'s `iterRows` on a wide sheet whose
// cells are almost all empty but styled (`<c r="X1" s="n"/>`), the shape
// business exports tend to have. Such a sheet is nearly all SAX events and
// almost no values, so any per-event overhead in the row walk shows up here
// at full strength (#302).
//
// Excluded from the default `pnpm test` run (see vitest.config.ts). Run
// explicitly:
//
//     pnpm test:perf
//     PERF_GATE=1 pnpm test:perf  # also assert the ratio
//
// The gate compares `iterRows` against `loadWorkbook` over the same bytes
// rather than against a cells/s floor, so the machine divides out. Holding
// one row at a time has no reason to be slower than building the whole
// worksheet model: before #302 the walk took 1.5–3× as long as the load, and
// after it a fraction.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { makeAlignment } from '../../src/styles/alignment.js';

const ROWS = 5_000;
const COLS = 227;
const TOTAL_CELLS = ROWS * COLS;
// Columns that hold a value; every other column is a styled empty cell.
const VALUE_COLS = new Set([0, 1, 3, 11, 89]);

const PERF_GATE = process.env['PERF_GATE'] === '1';
const MAX_ITER_TO_LOAD_RATIO = 1;

const ITERATIONS = 3;

const buildArchive = async (): Promise<Uint8Array> => {
  const sink = toBuffer();
  const wb = await createWriteOnlyWorkbook(sink);
  const ws = await wb.addWorksheet('Data');
  const style = { alignment: makeAlignment({ vertical: 'center', horizontal: 'left' }) };
  for (let r = 1; r <= ROWS; r++) {
    const row = Array.from({ length: COLS }, (_, c) => ({
      value: VALUE_COLS.has(c) ? (c === 1 ? `S${r}` : r + c) : null,
      style,
    }));
    await ws.appendRow(row);
  }
  await ws.close();
  await wb.finalize();
  return sink.result();
};

const timeIterRows = async (bytes: Uint8Array): Promise<{ seconds: number; cells: number }> => {
  const t0 = performance.now();
  const wb = await loadWorkbookStream(fromBuffer(bytes));
  let cells = 0;
  for await (const row of wb.openWorksheet('Data').iterRows()) cells += row.length;
  await wb.close();
  return { seconds: (performance.now() - t0) / 1000, cells };
};

const timeLoad = async (bytes: Uint8Array): Promise<number> => {
  const t0 = performance.now();
  await loadWorkbook(fromBuffer(bytes));
  return (performance.now() - t0) / 1000;
};

describe('perf: iterRows over styled empty cells', () => {
  it(
    `walks ${ROWS} × ${COLS} = ${TOTAL_CELLS.toLocaleString()} mostly-empty styled cells no slower than loadWorkbook`,
    async () => {
      const bytes = await buildArchive();
      // Warmup, discarded: the first pass pays JIT compile for both paths.
      await timeIterRows(bytes);
      await timeLoad(bytes);
      const iterRuns: Array<{ seconds: number; cells: number }> = [];
      const loadRuns: number[] = [];
      for (let i = 0; i < ITERATIONS; i++) {
        iterRuns.push(await timeIterRows(bytes));
        loadRuns.push(await timeLoad(bytes));
      }
      const bestIter = Math.min(...iterRuns.map((r) => r.seconds));
      const bestLoad = Math.min(...loadRuns);
      const ratio = bestIter / bestLoad;
      process.stderr.write(
        `[perf] ${TOTAL_CELLS.toLocaleString()} styled cells → iterRows best ${bestIter.toFixed(2)}s, loadWorkbook best ${bestLoad.toFixed(2)}s, ratio ${ratio.toFixed(2)} (max ${MAX_ITER_TO_LOAD_RATIO})\n`,
      );

      // A walk that dropped cells would post a flattering time.
      for (const run of iterRuns) expect(run.cells).toBe(TOTAL_CELLS);
      if (PERF_GATE) expect(ratio).toBeLessThanOrEqual(MAX_ITER_TO_LOAD_RATIO);
    },
    /* timeout */ 10 * 60_000,
  );
});
