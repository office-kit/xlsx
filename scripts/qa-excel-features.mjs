// Generate the bounded feature corpus separately from the cell/browser corpus.
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
const output = resolve(process.argv[2] ?? '.qa/excel-feature-corpus');
const result = spawnSync(process.execPath, ['node_modules/vitest/vitest.mjs', 'run', 'tests/conformance/excel-features.test.ts'], { stdio: 'inherit', env: { ...process.env, QA_EXCEL_FEATURES_OUTPUT: output } });
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
