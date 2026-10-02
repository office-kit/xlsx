import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import { validateMatrix, renderMatrix } from './qa-matrix.mjs';
const read = path => readFileSync(path, 'utf8');
const normative = JSON.parse(read('tests/conformance/corpus/manifest.json'));
const fresh = () => JSON.parse(read('tests/conformance/coverage.json'));
test('real inventory resolves evidence and renders explicit Windows and production-rejection gaps', () => {
  const matrix = validateMatrix(fresh(), normative);
  const report = renderMatrix(matrix);
  const windows = report.split('## excel-windows')[1].split('\n## ')[0];
  assert.ok(windows);
  assert.doesNotMatch(windows, /\| T \|/);
  const node = report.split('## node-model')[1].split('## node-stream')[0];
  assert.match(node, /coordinate-uri-zip32 \| U \| U \| U \| U/);
  const oracle = report.split('## independent-oracles')[1].split('## libreoffice')[0];
  assert.match(oracle, /coordinate-uri-zip32 \| U \| U \| U \| T/);
});
const faults = [
  ['empty requirements', m => { m.requirements = []; }, /Empty/],
  ['duplicate requirements', m => m.requirements.push(m.requirements[0]), /Duplicate/],
  ['unlinked normative clause', m => { m.requirements[0].clauses = ['invented specification clause']; }, /Corpus clause missing/],
  ['unknown corpus case', m => m.requirements[0].caseIds.push('invented-case'), /Unknown corpus case/],
  ['new unassigned corpus case', (_, c) => c.cases.push({ id: 'new-unassigned' }), /missing from requirements/],
  ['duplicate case ownership', m => m.requirements[1].caseIds.push(m.requirements[0].caseIds[0]), /Duplicate/],
  ['unknown profile', m => { m.claims[0].profile = 'invented'; }, /Unknown claim/],
  ['unknown status', m => { m.claims[0].status = 'supported'; }, /Invalid claim status/],
  ['missing evidence', m => { m.claims[0].evidence = []; }, /Missing evidence/],
  ['missing evidence file', m => { m.claims[0].evidence[0].path = 'tests/nonexistent.test.ts'; }, /Missing evidence file/],
  ['missing test anchor', m => { m.claims[0].evidence[0].anchor = 'invented-test-title'; }, /Missing evidence anchor/],
  ['path traversal', m => { m.claims[0].evidence[0].path = '../private.txt'; }, /Unsafe evidence path/],
  ['duplicate operation cell', m => m.claims.push(m.claims[0]), /Duplicate claim cell/],
  ['partial tested corpus', m => { m.claims[0].caseIds.pop(); }, /Incomplete tested case set/],
  ['foreign case in claim', m => { m.claims[0].caseIds[0] = m.requirements[1].caseIds[0]; }, /Incomplete tested case set/],
  ['invalid operation', m => { m.claims[0].operations = ['render']; }, /Invalid operation/],
  ['unavailable Windows called tested', m => { m.claims[0].profile = 'excel-windows'; }, /Unavailable profile/],
  ['unrecorded Mac case', m => { const c = m.claims.find(candidate => candidate.profile === 'excel-mac'); c.requirement = 'font-booleans'; c.caseIds = m.requirements[0].caseIds; }, /Incomplete tested case set/],
];
for (const [name, mutate, error] of faults) test(`reject ${name}`, () => {
  const matrix = fresh(); const corpus = structuredClone(normative);
  mutate(matrix, corpus);
  assert.throws(() => validateMatrix(matrix, corpus), error);
});
for (const [name, mutate, error] of [
  ['missing Mac evidence', o => { o.cases.pop(); }, /Manual observation set mismatch/],
  ['unobserved Mac open', o => { o.cases[0].openedWithoutRepair = false; }, /Incomplete manual observation/],
]) test(`reject ${name}`, () => {
  const source = path => {
    const body = read(path);
    if (!path.endsWith('excel-mac-observation.json')) return body;
    const observation = JSON.parse(body); mutate(observation); return JSON.stringify(observation);
  };
  assert.throws(() => validateMatrix(fresh(), normative, source), error);
});

for (const [name, mutate, error] of [
  ['unknown external feature case', m => { m.profiles.find(p => p.id === 'excel-mac-features').caseManifest = 'tests/conformance/corpus/manifest.json'; }, /Unknown profile case/],
  ['unsafe feature manifest', m => { m.profiles.find(p => p.id === 'excel-mac-features').caseManifest = '../private.json'; }, /Unsafe evidence path/],
]) test(`reject ${name}`, () => {
  const matrix = fresh(); mutate(matrix);
  assert.throws(() => validateMatrix(matrix, normative), error);
});
test('reject duplicate external feature cases', () => {
  const source = path => {
    const body = read(path);
    if (!path.endsWith('excel-features-manifest.json')) return body;
    const manifest = JSON.parse(body); manifest.cases.push(manifest.cases[0]);
    return JSON.stringify(manifest);
  };
  assert.throws(() => validateMatrix(fresh(), normative, source), /Duplicate/);
});
