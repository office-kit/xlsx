import { spawnSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const output = '.qa/resources';
export const limits = { heapMiB: 128, timeoutMs: 15_000, peakRssBytes: 256 * 1024 * 1024 };
export function runIsolated(worker, args = [], budget = limits) {
  const start = performance.now();
  const result = spawnSync(process.execPath, [`--max-old-space-size=${budget.heapMiB}`, worker, ...args], {
    encoding: 'utf8', timeout: budget.timeoutMs, killSignal: 'SIGKILL', maxBuffer: 512 * 1024,
    env: { ...process.env, NODE_OPTIONS: '' },
  });
  if (result.error || result.status !== 0 || result.signal) throw new Error(`Isolated QA worker failed (${result.error?.code ?? result.signal ?? result.status}): ${result.stderr}`, { cause: result.error });
  const report = JSON.parse(result.stdout);
  if (!Number.isSafeInteger(report.peakRssBytes) || report.peakRssBytes <= 0 || report.peakRssBytes > budget.peakRssBytes) throw new Error('Missing or excessive peak RSS');
  return { ...report, elapsedMs: Math.round(performance.now() - start) };
}
function main() {
  rmSync(output, { recursive: true, force: true }); mkdirSync(output, { recursive: true });
  const results = { node: process.version, platform: process.platform, arch: process.arch, limits, cases: [], status: 'running' };
  const save = () => writeFileSync(`${output}/results.json`, JSON.stringify(results, null, 2) + '\n');
  save();
  try {
    for (const id of ['inflation', 'rows', 'cells', 'doctype', 'dimension']) for (const mode of ['model', 'stream']) {
      const report = runIsolated('scripts/qa-resource-worker.mjs', [id, mode]);
      const expected = ({ inflation: 'OpenXmlDecompressionBombError', rows: 'OpenXmlContentLimitError', cells: 'OpenXmlContentLimitError', doctype: 'OpenXmlSchemaError', dimension: 'accepted' })[id];
      if (report.id !== id || report.mode !== mode || report.outcome !== expected || !Number.isSafeInteger(report.inputBytes) || report.inputBytes <= 0) throw new Error(`Wrong worker evidence for ${id}/${mode}`);
      results.cases.push(report); save();
      console.info(`${id}/${mode}: ${report.outcome}, ${report.elapsedMs}ms, ${Math.round(report.peakRssBytes / 1024 / 1024)}MiB peak RSS`);
    }
    results.status = 'passed'; save();
  } catch (error) { results.status = 'failed'; results.error = String(error); save(); throw error; }
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
