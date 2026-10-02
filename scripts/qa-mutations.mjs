// Curated regressions, not an arbitrary mutation score. Each must fail a real
// assertion; infrastructure errors and anchor drift cannot count as detection.
import { spawnSync } from 'node:child_process';
import { cpSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
const root = resolve('.');
const mutants = JSON.parse(readFileSync(join(root, 'tests/conformance/mutations.json'), 'utf8'));
const scratch = mkdtempSync(join(tmpdir(), 'xlsx-mutations-'));
const output = resolve('.qa/mutations');
mkdirSync(output, { recursive: true });
const results = [];
function run(tests, id) {
  const report = join(scratch, `${id}.json`);
  const result = spawnSync(process.execPath, [join(root, 'node_modules/vitest/vitest.mjs'), 'run', ...tests, '--reporter=json', `--outputFile=${report}`], { cwd: scratch, encoding: 'utf8', timeout: 120_000, env: { ...process.env, QA_CORPUS_OUTPUT: '', QA_FUZZ_RUNS: '1' } });
  writeFileSync(join(output, `${id}.log`), `${result.stdout ?? ''}\n${result.stderr ?? ''}`);
  if (result.error || !existsSync(report)) throw new Error(`${id}: runner error ${result.error ?? result.status}`);
  const data = JSON.parse(readFileSync(report, 'utf8'));
  const failures = (data.testResults ?? []).flatMap(s => (s.assertionResults ?? []).filter(a => a.status === 'failed').map(a => a.fullName));
  if (result.status === 0 && data.success && data.numPassedTests > 0) return { status: 'passed', failures: [] };
  if (result.status === 1 && data.numFailedTests > 0 && failures.length && !data.numRuntimeErrorTestSuites) return { status: 'killed', failures };
  throw new Error(`${id}: infrastructure failure or empty test run; cannot count as killed`);
}
try {
  for (const path of ['src', 'tests', 'vitest.config.ts', 'tsconfig.json', 'package.json']) cpSync(join(root, path), join(scratch, path), { recursive: true });
  symlinkSync(join(root, 'node_modules'), join(scratch, 'node_modules'), process.platform === 'win32' ? 'junction' : 'dir');
  symlinkSync(join(root, 'reference'), join(scratch, 'reference'), process.platform === 'win32' ? 'junction' : 'dir');
  const tests = [...new Set(mutants.flatMap(m => m.tests))];
  if (run(tests, 'baseline').status !== 'passed') throw new Error('Baseline failed');
  for (const m of mutants) {
    const file = join(scratch, m.file);
    const original = readFileSync(file, 'utf8');
    if (original.split(m.before).length !== 2 || m.before === m.after) throw new Error(`${m.id}: mutation anchor drift`);
    writeFileSync(file, original.replace(m.before, m.after));
    let result;
    try { result = run(m.tests, m.id); } finally { writeFileSync(file, original); }
    results.push({ id: m.id, regression: m.regression, ...result });
    console.info(`${m.id}: ${result.status}`);
  }
  if (results.some(r => r.status !== 'killed')) throw new Error('A curated regression survived');
} finally {
  writeFileSync(join(output, 'results.json'), JSON.stringify({ profile: 'curated-regressions-v1', results }, null, 2));
  rmSync(scratch, { recursive: true, force: true });
}
