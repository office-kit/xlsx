// Replay a committed native-Excel fixture; this does not launch desktop Excel.
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
const result = spawnSync(process.execPath, ['node_modules/vitest/vitest.mjs', 'run', 'tests/conformance/native-excel.test.ts'], {
  stdio: 'inherit', env: { ...process.env, QA_NATIVE_EXCEL_OUTPUT: resolve('.qa/native-excel-corpus') },
});
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
