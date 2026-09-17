import { describe, expect, it } from 'vitest';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';

const MIB = 1024 * 1024;
const GATE = process.env['PERF_GATE'] === '1';

describe('write-only string retention', () => {
  for (const rich of [false, true]) {
    it(`bounds retained heap for distinct ${rich ? 'rich' : 'plain'} strings`, async () => {
      const gc = globalThis.gc;
      expect(gc).toBeTypeOf('function');
      gc?.();
      const baseline = process.memoryUsage().heapUsed;
      let outputBytes = 0;
      const wb = await createWriteOnlyWorkbook({
        toBytes: () => ({
          write: (chunk) => { outputBytes += chunk.byteLength; },
          finish: async () => new Uint8Array(0),
        }),
      });
      const ws = await wb.addWorksheet('Distinct strings');
      const retained: number[] = [];
      for (let r = 0; r < 750_000; r++) {
        const a = `row-${r}`;
        const b = `value-${r}`;
        await ws.appendRow(rich
          ? [r, { kind: 'rich-text', runs: [{ text: a, font: { b: true } }, { text: b }] }]
          : [r, a, b]);
        if (r === 249_999 || r === 749_999) {
          // Measure while the workbook is live, before finalize can release it.
          gc?.();
          retained.push((process.memoryUsage().heapUsed - baseline) / MIB);
        }
      }
      await ws.close();
      await wb.finalize();
      expect(outputBytes).toBeGreaterThan(0);
      const early = retained[0] ?? Infinity;
      const late = retained[1] ?? Infinity;
      process.stderr.write(`[perf-strings] ${rich ? 'rich' : 'plain'} 250k/750k rows: ${early.toFixed(1)}/${late.toFixed(1)} MiB retained heap\n`);
      if (GATE) {
        // Leave room for Node/V8 differences, while rejecting linear retention.
        expect(late).toBeLessThan(64);
        expect(late - early).toBeLessThan(16);
      }
    }, 120_000);
  }
});
