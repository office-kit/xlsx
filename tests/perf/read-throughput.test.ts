// Throughput gate for the read path, the mirror of throughput.test.ts. Reading
// is the direction most callers spend their time in, and `loadWorkbook` is
// ~98% of the cost of a load-then-walk, so a regression in the `<sheetData>`
// walk is invisible to the write-side gate.
//
// Excluded from the default `pnpm test` run (see vitest.config.ts). Run
// explicitly:
//
//     pnpm test:perf
//     PERF_GATE=1 pnpm test:perf  # also assert the throughput floor
//
// 50k rows × 6 text columns is the shape the sheetData walk is tuned for: wide
// enough that per-cell cost dominates, small enough that one run is a few
// seconds.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { getSheet } from '../../src/workbook/workbook.js';
import { getNonEmptyCellCount } from '../../src/worksheet/worksheet.js';

const ROWS = 50_000;
const COLS = 6;
const TOTAL_CELLS = ROWS * COLS;

// PERF_GATE off by default: a laptop idling or CI under load can dip below the
// floor without the code being wrong. Set the env var when you want a hard
// assertion (release branches, perf-regression PRs).
const PERF_GATE = process.env['PERF_GATE'] === '1';
// Set to catch a return to building a node per `<c>`, which reads this shape at
// roughly 150k cells/s, and not to certify any one machine's ceiling. A shared
// CI runner has to clear it with room to spare or the gate is just flaky.
const FLOOR_CELLS_PER_SEC = 200_000;

const ITERATIONS = 3;

const buildArchive = async (): Promise<Uint8Array> => {
  const sink = toBuffer();
  const wb = await createWriteOnlyWorkbook(sink);
  const ws = await wb.addWorksheet('Data');
  for (let r = 0; r < ROWS; r++) {
    const row = new Array<string>(COLS);
    for (let c = 0; c < COLS; c++) row[c] = `r${r}c${c}`;
    await ws.appendRow(row);
  }
  await ws.close();
  await wb.finalize();
  return sink.result();
};

const measureOnce = async (bytes: Uint8Array): Promise<{ seconds: number; cells: number }> => {
  const t0 = performance.now();
  const wb = await loadWorkbook(fromBuffer(bytes));
  const t1 = performance.now();
  const ws = getSheet(wb, 'Data');
  if (ws === undefined) throw new Error('load produced no Data sheet');
  return { seconds: (t1 - t0) / 1000, cells: getNonEmptyCellCount(ws) };
};

describe('perf: loadWorkbook read throughput', () => {
  it(
    `reads ${ROWS} × ${COLS} = ${TOTAL_CELLS.toLocaleString()} cells and reports cells/s`,
    async () => {
      const bytes = await buildArchive();
      // Warmup, discarded: the first pass pays JIT compile and heap growth for
      // the whole read path and reads systematically low.
      await measureOnce(bytes);
      // Best-of-N: shared CPUs and thermal throttling create wide variance per
      // run, but the best run reflects the pipeline's real ceiling.
      const runs: Array<{ seconds: number; cells: number }> = [];
      for (let i = 0; i < ITERATIONS; i++) runs.push(await measureOnce(bytes));
      const bestSeconds = Math.min(...runs.map((r) => r.seconds));
      const bestCellsPerSec = Math.round(TOTAL_CELLS / bestSeconds);

      const summaries = runs
        .map((r, i) => `#${i + 1} ${(TOTAL_CELLS / r.seconds).toFixed(0)} cells/s (${r.seconds.toFixed(2)}s)`)
        .join(' · ');
      process.stderr.write(
        `[perf] ${TOTAL_CELLS.toLocaleString()} cells × ${ITERATIONS} runs → best ${bestCellsPerSec.toLocaleString()} cells/s; archive ${bytes.byteLength.toLocaleString()} bytes\n        runs: ${summaries}\n`,
      );

      // A load that dropped cells would post a flattering cells/s, so the gate
      // only means anything alongside the count.
      for (const run of runs) expect(run.cells).toBe(TOTAL_CELLS);
      if (PERF_GATE) {
        expect(bestCellsPerSec).toBeGreaterThanOrEqual(FLOOR_CELLS_PER_SEC);
      }
    },
    /* timeout */ 10 * 60_000,
  );
});
