import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
const output = resolve(process.argv[2] ?? '.qa/corpus');
const result = spawnSync(process.execPath, ['node_modules/vitest/vitest.mjs', 'run', 'tests/conformance/corpus.test.ts'], { stdio: 'inherit', env: { ...process.env, QA_CORPUS_OUTPUT: output } });
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
