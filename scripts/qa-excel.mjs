// Native desktop Excel is a manual compatibility profile, never a hidden CI skip.
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
const mode = process.argv[2];
const directory = resolve(process.argv[3] ?? '.qa/excel-manual');
const input = join(directory, 'input');
const output = join(directory, 'output');
const profile = JSON.parse(readFileSync('tests/conformance/corpus/excel-profile.json', 'utf8'));
const hash = path => createHash('sha256').update(readFileSync(path)).digest('hex');
const sameSet = (actual, expected) => actual.length === expected.length && new Set(actual).size === actual.length && [...actual].sort().join('\n') === [...expected].sort().join('\n');
const normative = JSON.parse(readFileSync('tests/conformance/corpus/manifest.json', 'utf8')).cases;
if (!profile.cases.length || new Set(profile.cases).size !== profile.cases.length) throw new Error('Empty or duplicate Excel profile');
const expectedCases = profile.cases.map(id => {
  const c = normative.find(candidate => candidate.id === id);
  if (!c) throw new Error(`Missing normative case ${id}`);
  return c;
});
if (mode === '--prepare') {
  if (existsSync(directory)) throw new Error('Use a fresh session directory; refusing stale manual results');
  const corpus = JSON.parse(readFileSync('.qa/corpus/manifest.json', 'utf8'));
  const cases = profile.cases.map(id => {
    const c = corpus.find(candidate => candidate.id === id);
    if (!c) throw new Error(`Missing corpus case ${id}`);
    return c;
  });
  if (JSON.stringify(cases) !== JSON.stringify(expectedCases)) throw new Error('Regenerate qa:corpus after corpus changes');
  mkdirSync(input, { recursive: true }); mkdirSync(output);
  writeFileSync(join(input, 'manifest.json'), JSON.stringify(cases, null, 2));
  for (const c of cases) copyFileSync(`.qa/corpus/${c.id}.output.xlsx`, join(input, `${c.id}.output.xlsx`));
  const observations = { profile: profile.profile, application: 'Microsoft Excel', version: '', platform: process.platform, method: 'native-ui-manual', cases: cases.map(c => ({ id: c.id, openedWithoutRepair: false, observedAt: '', inputSha256: hash(join(input, `${c.id}.output.xlsx`)), outputSha256: '' })) };
  writeFileSync(join(directory, 'observations.json'), JSON.stringify(observations, null, 2));
  console.info(`Open each file in ${input} in desktop Excel. Do not accept repair. Save As the same basename into ${output}. Record version, timestamp, no-repair observation and saved SHA256 in observations.json; then run qa:excel --check.`);
} else if (mode === '--check') {
  // A failed rerun must not leave yesterday's passing report behind.
  for (const path of [join(directory, 'results.json'), join(output, 'results.json'), join(directory, 'check.log')]) rmSync(path, { force: true });
  const observations = JSON.parse(readFileSync(join(directory, 'observations.json'), 'utf8'));
  if (observations.profile !== profile.profile || observations.application !== 'Microsoft Excel' || !observations.version?.trim() || !observations.platform || observations.method !== 'native-ui-manual') throw new Error('Missing Excel provenance');
  const cases = JSON.parse(readFileSync(join(input, 'manifest.json'), 'utf8'));
  if (!sameSet(cases.map(c => c.id), profile.cases) || !sameSet(observations.cases.map(c => c.id), profile.cases)) throw new Error('Manual observation set does not match profile');
  if (JSON.stringify(cases) !== JSON.stringify(expectedCases)) throw new Error('Input manifest differs from normative corpus');
  const expected = cases.map(c => `${c.id}.output.xlsx`);
  if (!sameSet(readdirSync(input).filter(f => f.endsWith('.xlsx')), expected)) throw new Error('Input file set does not match profile');
  if (!sameSet(readdirSync(output).filter(f => f.endsWith('.xlsx')), expected)) throw new Error('Saved output set does not match profile');
  for (const c of observations.cases) {
    if (c.openedWithoutRepair !== true || !Number.isFinite(Date.parse(c.observedAt))) throw new Error(`Missing no-repair observation: ${c.id}`);
    if (c.inputSha256 !== hash(join(input, `${c.id}.output.xlsx`)) || c.outputSha256 !== hash(join(output, `${c.id}.output.xlsx`))) throw new Error(`Observation hash mismatch: ${c.id}`);
  }
  const check = spawnSync('python3', ['scripts/qa-office-check.py', input, output, 'excel'], { encoding: 'utf8', timeout: 30_000 });
  writeFileSync(join(directory, 'check.log'), `${check.stdout ?? ''}\n${check.stderr ?? ''}`);
  if (check.error || check.status !== 0) throw new Error(`Independent Excel comparison failed: ${check.error ?? check.stderr ?? check.status}`);
  writeFileSync(join(directory, 'results.json'), JSON.stringify({ profile: profile.profile, status: 'pass', provenance: observations, semantic: JSON.parse(readFileSync(join(output, 'results.json'), 'utf8')) }, null, 2));
  process.stdout.write(check.stdout);
} else throw new Error('Usage: qa:excel --prepare|--check [fresh-session-directory]');
