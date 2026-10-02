import { spawnSync } from 'node:child_process';
import { rmSync } from 'node:fs';
import { resolve } from 'node:path';
const output = resolve('.qa/feature-interactions');
rmSync(output, { recursive: true, force: true });
const result = spawnSync(process.execPath, ['node_modules/vitest/vitest.mjs', 'run', 'tests/conformance/feature-interactions.test.ts'], {
  stdio: 'inherit', env: { ...process.env, QA_FEATURE_INTERACTIONS_OUTPUT: output },
});
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
