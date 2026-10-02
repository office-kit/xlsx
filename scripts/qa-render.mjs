import { spawnSync } from 'node:child_process';
import { rmSync } from 'node:fs';
import { resolve } from 'node:path';
const output = resolve('.qa/render');
rmSync(output, { recursive: true, force: true });
for (const [command, args, extra] of [
  [process.execPath, ['node_modules/vitest/vitest.mjs', 'run', 'tests/conformance/render-roundtrip.test.ts'], { QA_RENDER_OUTPUT: output }],
  ['python3', ['scripts/qa-render-check.py', output], {}],
]) {
  const result = spawnSync(command, args, { stdio: 'inherit', env: { ...process.env, ...extra } });
  if (result.error) throw result.error;
  if (result.status !== 0) { process.exitCode = result.status ?? 1; break; }
}
