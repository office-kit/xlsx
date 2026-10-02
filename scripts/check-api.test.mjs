import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { collectApi, digest, validateApiChange } from './check-api.mjs';
const fixture = body => {
  const root = mkdtempSync(join(tmpdir(), 'xlsx-api-calibration-'));
  mkdirSync(join(root, 'dist'));
  writeFileSync(join(root, 'package.json'), JSON.stringify({ type: 'module', version: '0.23.3', exports: { './sample': { types: './dist/index.d.ts', import: './dist/sample.mjs' } } }));
  writeFileSync(join(root, 'dist/index.d.ts'), "export { call } from './model.js'; export type { Options } from './model.js';\n");
  writeFileSync(join(root, 'dist/model.d.ts'), body);
  return root;
};
const original = 'export interface Options { value: number; mode: "one" | "two"; } export declare function call(options: Options): number;';
for (const [name, body] of [
  ['parameter change', original.replace('options: Options', 'options: string')],
  ['transitive member change', original.replace('value: number', 'value: string')],
  ['union member removal', original.replace('"one" | "two"', '"one"')],
  ['export removal', original.replace('export declare function call(options: Options): number;', '')],
]) test(`detects ${name}`, () => {
  const root = fixture(original);
  try {
    const before = collectApi(root);
    writeFileSync(join(root, 'dist/model.d.ts'), body);
    assert.notEqual(digest(collectApi(root)), digest(before));
  } finally { rmSync(root, { recursive: true, force: true }); }
});
test('ignores documentation comments but fails a missing declaration entrypoint', () => {
  const root = fixture(original);
  try {
    const before = collectApi(root);
    writeFileSync(join(root, 'dist/model.d.ts'), `/** Documentation only. */\n${original}`);
    assert.equal(digest(collectApi(root)), digest(before));
    rmSync(join(root, 'dist/index.d.ts'));
    assert.throws(() => collectApi(root), /Missing public declaration/);
  } finally { rmSync(root, { recursive: true, force: true }); }
});
const before = { declarations: { a: 'old' } }; const after = { declarations: { a: 'new' } };
const approval = { baselineSha256: digest(after), kind: 'breaking', reason: 'Rename the old public API and document migration.' };
const changeset = bump => `---\n"@office-kit/xlsx": ${bump}\n---\n\nBreaking: replace the old API.\n`;
test('requires hash-bound explanation, intentional break and correct pre/post-1 bump', () => {
  assert.throws(() => validateApiChange(before, after, {}, '0.23.3', changeset('minor')), /explanation/);
  assert.throws(() => validateApiChange(before, after, { ...approval, baselineSha256: 'wrong' }, '0.23.3', changeset('minor')), /explanation/);
  assert.throws(() => validateApiChange(before, after, approval, '0.23.3', changeset('patch')), /minor/);
  assert.throws(() => validateApiChange(before, after, approval, '1.0.0', changeset('minor')), /major/);
  assert.throws(() => validateApiChange(before, after, approval, '0.23.3', changeset('minor').replace('Breaking:', 'Replace')), /flagged/);
  validateApiChange(before, after, approval, '0.23.3', changeset('minor'));
  validateApiChange(before, after, approval, '1.0.0', changeset('major'));
  validateApiChange(before, after, { ...approval, kind: 'correction' }, '1.0.0', changeset('patch'));
  assert.throws(() => validateApiChange(before, after, { ...approval, kind: 'additive' }, '1.0.0', changeset('patch')), /minor/);
});

test('CLI rejects refreshed baselines without a new changeset and approval record', () => {
  const root = fixture(original);
  const script = fileURLToPath(new URL('./check-api.mjs', import.meta.url));
  const run = (...args) => spawnSync(process.execPath, [script, ...args], { cwd: root, encoding: 'utf8' });
  const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8', stdio: 'pipe' });
  try {
    mkdirSync(join(root, 'tests/consumer'), { recursive: true });
    assert.equal(run('--write').status, 0);
    git('init'); git('add', '.');
    git('-c', 'user.name=QA', '-c', 'user.email=qa@example.invalid', 'commit', '-m', 'baseline');
    const base = git('rev-parse', 'HEAD').trim();
    writeFileSync(join(root, 'dist/model.d.ts'), original.replace('value: number', 'value: string'));
    assert.notEqual(run('--write', '--base', base).status, 0);
    const record = { baselineSha256: digest(collectApi(root)), kind: 'correction', reason: 'Correct the documented property type.', changeset: '.changeset/type-correction.md' };
    writeFileSync(join(root, 'tests/consumer/api-change.json'), JSON.stringify(record));
    assert.notEqual(run('--base', base).status, 0);
    mkdirSync(join(root, '.changeset'));
    writeFileSync(join(root, record.changeset), changeset('patch'));
    git('add', '.');
    const accepted = run('--base', base);
    assert.equal(accepted.status, 0, accepted.stderr);
    assert.notEqual(run('--bogus').status, 0);
  } finally { rmSync(root, { recursive: true, force: true }); }
});
