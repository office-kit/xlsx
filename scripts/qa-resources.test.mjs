import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { runIsolated, limits } from './qa-resources.mjs';
function worker(code, check) {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-resource-calibration-'));
  const file = join(dir, 'worker.mjs'); writeFileSync(file, code);
  try { check(file); } finally { rmSync(dir, { recursive: true, force: true }); }
}
test('a timeout is a runner failure, never a successful rejection', () => {
  worker('Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0);', file => assert.throws(() => runIsolated(file, [], { ...limits, timeoutMs: 200 }), /ETIMEDOUT/));
});
test('heap exhaustion is a runner failure, never a successful rejection', () => {
  worker('const live = []; while (true) live.push(Array(100000).fill("live"));', file => assert.throws(() => runIsolated(file, [], { ...limits, heapMiB: 16, timeoutMs: 5000 }), /worker failed/));
});
test('missing reports, crashes and excessive RSS cannot pass', () => {
  for (const code of ['', 'process.exit(1);', 'console.log(JSON.stringify({peakRssBytes: 0}));', `console.log(JSON.stringify({peakRssBytes: ${limits.peakRssBytes + 1}}));`]) worker(code, file => assert.throws(() => runIsolated(file)));
});
