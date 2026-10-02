// Calibrate bookkeeping and independent comparison, without pretending to run Excel.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { cpSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
const hash = path => createHash('sha256').update(readFileSync(path)).digest('hex');
const run = (mode, path, features = false) => spawnSync(process.execPath, ['scripts/qa-excel.mjs', mode, path, ...(features ? ['--features'] : [])], { encoding: 'utf8', timeout: 30_000 });
const read = path => JSON.parse(readFileSync(path, 'utf8'));
const write = (path, value) => writeFileSync(path, JSON.stringify(value));
function session(check, features = false) {
  const root = mkdtempSync(join(tmpdir(), 'xlsx-excel-calibration-'));
  const path = join(root, 'session');
  try {
    const prepared = run('--prepare', path, features);
    assert.equal(prepared.status, 0, prepared.stderr);
    cpSync(join(path, 'input'), join(path, 'output'), { recursive: true });
    const file = join(path, 'observations.json');
    const observations = read(file);
    observations.version = 'synthetic calibration; Excel was not executed';
    for (const c of observations.cases) {
      c.openedWithoutRepair = true;
      c.observedAt = new Date().toISOString();
      c.outputSha256 = hash(join(path, 'output', `${c.id}.output.xlsx`));
    }
    write(file, observations);
    check(path, file, observations);
  } finally { rmSync(root, { recursive: true, force: true }); }
}
test('synthetic bookkeeping control passes independent comparison', () => session(path => {
  const checked = run('--check', path);
  assert.equal(checked.status, 0, checked.stderr);
  assert.equal(read(join(path, 'results.json')).semantic.cases.length, 9);
}));
const faults = [
  ['unobserved open', (_, __, o) => { o.cases[0].openedWithoutRepair = false; }, /Missing no-repair observation/],
  ['missing timestamp', (_, __, o) => { o.cases[0].observedAt = ''; }, /Missing no-repair observation/],
  ['missing version', (_, __, o) => { o.version = ''; }, /Missing Excel provenance/],
  ['duplicate observations', (_, __, o) => { o.cases[1] = o.cases[0]; }, /does not match profile/],
  ['wrong input hash', (_, __, o) => { o.cases[0].inputSha256 = '0'.repeat(64); }, /hash mismatch/],
  ['missing saved file', (p, _, o) => rmSync(join(p, 'output', `${o.cases[0].id}.output.xlsx`)), /Saved output set/],
  ['extra saved file', p => writeFileSync(join(p, 'output', 'unrelated.xlsx'), ''), /Saved output set/],
  ['changed expected meaning', p => { const f = join(p, 'input', 'manifest.json'); const m = read(f); m[0].officeValue = 999; write(f, m); }, /differs from normative corpus/],
  ['changed saved bytes', (p, _, o) => writeFileSync(join(p, 'output', `${o.cases[0].id}.output.xlsx`), 'changed'), /hash mismatch/],
];
for (const [name, mutate, error] of faults) test(`reject ${name} and remove stale success`, () => session((path, file, observations) => {
  write(join(path, 'results.json'), { status: 'pass' });
  mutate(path, file, observations);
  write(file, observations);
  const checked = run('--check', path);
  assert.equal(checked.status, 1, checked.stderr);
  assert.match(checked.stderr, error);
  assert.equal(existsSync(join(path, 'results.json')), false);
}));

test('correct hashes cannot hide a changed numeric value', () => session((path, file, observations) => {
  const c = observations.cases[0];
  const output = join(path, 'output', `${c.id}.output.xlsx`);
  const parts = unzipSync(readFileSync(output));
  const original = strFromU8(parts['xl/worksheets/sheet1.xml']);
  assert.match(original, /<v>3<\/v>/);
  parts['xl/worksheets/sheet1.xml'] = strToU8(original.replace('<v>3</v>', '<v>999</v>'));
  writeFileSync(output, zipSync(parts));
  c.outputSha256 = hash(output);
  write(file, observations);
  const checked = run('--check', path);
  assert.equal(checked.status, 1, checked.stderr);
  assert.match(checked.stderr, /Independent Excel comparison failed/);
  const failed = read(join(path, 'output', 'results.json'));
  assert.equal(failed.cases[0].status, 'fail');
  assert.equal(existsSync(join(path, 'results.json')), false);
}));

test('correct values and hashes cannot hide a lost date format', () => session((path, file, observations) => {
  const c = observations.cases.find(candidate => candidate.id.startsWith('date-'));
  assert.ok(c);
  const output = join(path, 'output', `${c.id}.output.xlsx`);
  const parts = unzipSync(readFileSync(output));
  const original = strFromU8(parts['xl/styles.xml']);
  assert.match(original, /numFmtId="14"/);
  parts['xl/styles.xml'] = strToU8(original.replaceAll('numFmtId="14"', 'numFmtId="0"'));
  writeFileSync(output, zipSync(parts));
  c.outputSha256 = hash(output);
  write(file, observations);
  const checked = run('--check', path);
  assert.equal(checked.status, 1, checked.stderr);
  const failed = read(join(path, 'output', 'results.json'));
  assert.ok(failed.unexpectedDifferences.some(([context]) => context === `${c.id}:numberFormat`));
  assert.equal(existsSync(join(path, 'results.json')), false);
}));

for (const format of ['mm/dd/yyyy', 'm/d/yyyy']) test(`date format survives a custom ${format} expansion`, () => session((path, file, observations) => {
  for (const c of observations.cases.filter(candidate => candidate.id.startsWith('date-'))) {
    const output = join(path, 'output', `${c.id}.output.xlsx`);
    const parts = unzipSync(readFileSync(output));
    const original = strFromU8(parts['xl/styles.xml']);
    parts['xl/styles.xml'] = strToU8(original.replace('<fonts', `<numFmts count="1"><numFmt numFmtId="165" formatCode="${format}"/></numFmts><fonts`).replaceAll('numFmtId="14"', 'numFmtId="165"'));
    writeFileSync(output, zipSync(parts));
    c.outputSha256 = hash(output);
  }
  write(file, observations);
  const checked = run('--check', path);
  assert.equal(checked.status, 0, checked.stderr);
}));


test('synthetic feature bookkeeping control passes relationship-aware comparison', () => session(path => {
  const checked = run('--check', path, true);
  assert.equal(checked.status, 0, checked.stderr);
  assert.equal(read(join(path, 'results.json')).semantic.cases.length, 1);
}, true));
const featureFaults = [
  ['detached table with orphan part retained', 'xl/worksheets/sheet1.xml', xml => xml.replace(/<tableParts\b[^>]*>[\s\S]*?<\/tableParts>/, '')],
  ['detached drawing with orphan chart/image retained', 'xl/worksheets/sheet1.xml', xml => xml.replace(/<drawing\b[^>]*\/>/, '')],
  ['changed chart reference', 'xl/charts/chart1.xml', xml => xml.replace('Audit!$A$2:$A$3', 'Audit!$A$2:$A$2')],
  ['changed table range', 'xl/tables/table1.xml', xml => xml.replaceAll('A1:A3', 'A1:A2')],
  ['changed fit width', 'xl/worksheets/sheet1.xml', xml => xml.replace('fitToWidth="1"', 'fitToWidth="2"')],
  ['lost page orientation', 'xl/worksheets/sheet1.xml', xml => xml.replace('orientation="landscape"', 'orientation="portrait"')],
  ['moved image anchor', 'xl/drawings/drawing1.xml', xml => xml.replace('<xdr:row>14</xdr:row>', '<xdr:row>15</xdr:row>')],
];
for (const [name, part, mutate] of featureFaults) test(`correct hashes cannot hide ${name}`, () => session((path, file, observations) => {
  const c = observations.cases[0];
  const output = join(path, 'output', `${c.id}.output.xlsx`);
  const parts = unzipSync(readFileSync(output));
  const before = strFromU8(parts[part]); const after = mutate(before);
  assert.notEqual(after, before, 'Calibration mutation must alter the package');
  parts[part] = strToU8(after); writeFileSync(output, zipSync(parts));
  c.outputSha256 = hash(output); write(file, observations);
  const checked = run('--check', path, true);
  assert.equal(checked.status, 1, checked.stderr);
  assert.match(checked.stderr, /Independent Excel comparison failed/);
  const report = read(join(path, 'output', 'results.json'));
  assert.ok(report.unexpectedDifferences.some(([context]) => context === `${c.id}:features`));
  assert.equal(existsSync(join(path, 'results.json')), false);
}, true));
test('correct hashes cannot hide changed attached image bytes', () => session((path, file, observations) => {
  const c = observations.cases[0]; const output = join(path, 'output', `${c.id}.output.xlsx`);
  const parts = unzipSync(readFileSync(output));
  const image = Object.keys(parts).find(part => part.startsWith('xl/media/'));
  assert.ok(image); parts[image] = new Uint8Array([1, 2, 3]); writeFileSync(output, zipSync(parts));
  c.outputSha256 = hash(output); write(file, observations);
  assert.equal(run('--check', path, true).status, 1);
  assert.equal(existsSync(join(path, 'results.json')), false);
}, true));

test('Excel may omit the schema-default fit width', () => session((path, file, observations) => {
  const c = observations.cases[0]; const output = join(path, 'output', `${c.id}.output.xlsx`);
  const parts = unzipSync(readFileSync(output));
  const original = strFromU8(parts['xl/worksheets/sheet1.xml']);
  assert.match(original, /fitToWidth="1"/);
  parts['xl/worksheets/sheet1.xml'] = strToU8(original.replace('fitToWidth="1"', ''));
  writeFileSync(output, zipSync(parts)); c.outputSha256 = hash(output); write(file, observations);
  const checked = run('--check', path, true);
  assert.equal(checked.status, 0, checked.stderr);
}, true));
