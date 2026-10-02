import { spawnSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
const output = resolve('.qa/render');
rmSync(output, { recursive: true, force: true });
mkdirSync(output, { recursive: true });
const report = { status: 'running', stages: [] };
const save = () => writeFileSync(resolve(output, 'runner.json'), JSON.stringify(report, null, 2) + '\n');
save();
for (const [command, args, extra] of [
  [process.execPath, ['node_modules/vitest/vitest.mjs', 'run', 'tests/conformance/render-roundtrip.test.ts'], { QA_RENDER_OUTPUT: output }],
  ['python3', ['scripts/qa-render-check.py', output], {}],
]) {
  const result = spawnSync(command, args, { stdio: 'inherit', env: { ...process.env, ...extra } });
  report.stages.push({ command, status: result.status, signal: result.signal, error: result.error?.message });
  if (result.error || result.status !== 0) {
    report.status = 'failed';
    process.exitCode = result.status ?? 1;
    save();
    break;
  }
  save();
}

if (!process.exitCode) { report.status = 'passed'; save(); }
