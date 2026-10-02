// Evidence inventory, not a conformance score. Missing cells are explicitly untested.
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const operations = ['read', 'write', 'preserve', 'reject'];
const statuses = ['tested', 'unsupported', 'untested'];
const matrixPath = 'tests/conformance/coverage.json';
const reportPath = 'tests/conformance/coverage.md';
const corpusPath = 'tests/conformance/corpus/manifest.json';
const read = path => readFileSync(path, 'utf8');
const fail = message => { throw new Error(message); };
const text = (value, context) => {
  if (typeof value !== 'string' || !value.trim()) fail(`Missing text: ${context}`);
};
const unique = (values, context, nonempty = true) => {
  if (!Array.isArray(values) || (nonempty && !values.length)) fail(`Empty or invalid list: ${context}`);
  for (const value of values) text(value, context);
  if (new Set(values).size !== values.length) fail(`Duplicate: ${context}`);
};
function evidence(refs, source, context, required = true) {
  if (!Array.isArray(refs) || (required && !refs.length)) fail(`Missing evidence: ${context}`);
  for (const ref of refs) {
    text(ref.path, context); text(ref.anchor, context);
    if (!/^[\w.[\] /+-]+$/.test(ref.path) || ref.path.startsWith('/') || ref.path.split('/').includes('..') || ref.path.includes('\\')) fail(`Unsafe evidence path: ${ref.path}`);
    let body;
    try { body = source(ref.path); } catch { fail(`Missing evidence file: ${ref.path}`); }
    if (!body.includes(ref.anchor)) fail(`Missing evidence anchor: ${ref.path}: ${ref.anchor}`);
  }
}
export function validateMatrix(matrix, corpus, source = read) {
  if (matrix.formatVersion !== 1 || matrix.defaultStatus !== 'untested' || JSON.stringify(matrix.operations) !== JSON.stringify(operations)) fail('Invalid matrix format/default/operations');
  text(matrix.scope, 'matrix scope');
  if (!Array.isArray(matrix.profiles) || !Array.isArray(matrix.requirements) || !Array.isArray(matrix.claims)) fail('Invalid matrix collections');
  unique(matrix.profiles.map(p => p.id), 'profile IDs');
  unique(matrix.requirements.map(r => r.id), 'requirement IDs');
  unique(corpus.cases.map(c => c.id), 'corpus IDs');
  const corpusIds = new Set(corpus.cases.map(c => c.id));
  const profiles = new Map(matrix.profiles.map(p => [p.id, p]));
  const requirements = new Map(matrix.requirements.map(r => [r.id, r]));
  const limits = new Map();
  for (const profile of matrix.profiles) {
    if (!['library', 'oracle', 'consumer'].includes(profile.subject) || !['ci', 'manual', 'unavailable'].includes(profile.execution)) fail(`Invalid profile: ${profile.id}`);
    unique(profile.runtimes, `${profile.id} runtimes`); unique(profile.adapters, `${profile.id} adapters`);
    evidence(profile.evidence, source, profile.id);
    if (profile.caseLimit) {
      const limit = JSON.parse(source(profile.caseLimit));
      unique(limit.cases, `${profile.id} case limit`);
      let allowed = corpusIds;
      if (profile.caseManifest) {
        evidence([{ path: profile.caseManifest, anchor: '"cases"' }], source, `${profile.id} case manifest`);
        const ids = JSON.parse(source(profile.caseManifest)).cases.map(c => c.id);
        unique(ids, `${profile.id} manifest cases`);
        allowed = new Set(ids);
      }
      if (limit.cases.some(id => !allowed.has(id))) fail(`Unknown profile case: ${profile.id}`);
      limits.set(profile.id, new Set(limit.cases));
      if (profile.execution === 'manual') {
        const observation = JSON.parse(source(profile.observation));
        if (observation.profile !== limit.profile || observation.method !== 'native-ui-manual') fail('Manual observation/profile mismatch');
        text(observation.application, 'manual application'); text(observation.version, 'manual version'); text(observation.platform, 'manual platform');
        unique(observation.cases.map(c => c.id), 'manual observation cases');
        if (JSON.stringify(observation.cases.map(c => c.id).sort()) !== JSON.stringify([...limit.cases].sort())) fail('Manual observation set mismatch');
        for (const c of observation.cases) if (c.openedWithoutRepair !== true || !Number.isFinite(Date.parse(c.observedAt)) || !/^[a-f0-9]{64}$/.test(c.inputSha256) || !/^[a-f0-9]{64}$/.test(c.outputSha256)) fail('Incomplete manual observation');
      }

    } else if (profile.execution === 'manual') fail('Manual profile requires a bounded case limit');
  }
  const assigned = [];
  for (const req of matrix.requirements) {
    text(req.contract, req.id); unique(req.clauses, `${req.id} clauses`); unique(req.caseIds, `${req.id} cases`, false);
    for (const id of req.caseIds) {
      if (!corpusIds.has(id)) fail(`Unknown corpus case: ${id}`);
      assigned.push(id);
    }
  }
  unique(assigned, 'requirement corpus assignment');
  if (assigned.length !== corpusIds.size) fail('Corpus cases missing from requirements');
  for (const req of matrix.requirements) {
    for (const id of req.caseIds) if (!req.clauses.includes(corpus.cases.find(c => c.id === id).clause)) fail(`Corpus clause missing from requirement: ${id}`);
  }
  const cells = new Set();
  for (const claim of matrix.claims) {
    const req = requirements.get(claim.requirement); const profile = profiles.get(claim.profile);
    if (!req || !profile) fail('Unknown claim requirement/profile');
    if (!statuses.includes(claim.status)) fail(`Invalid claim status: ${claim.status}`);
    text(claim.scope, `${claim.requirement}/${claim.profile} scope`);
    unique(claim.operations, 'claim operations'); unique(claim.caseIds, 'claim cases', false);
    evidence(claim.evidence, source, 'claim', claim.status !== 'untested');
    if (claim.status === 'tested') {
      if (profile.execution === 'unavailable') fail('Unavailable profile cannot be tested');
      const limit = limits.get(profile.id);
      const expected = req.caseIds.filter(id => !limit || limit.has(id));
      if (JSON.stringify([...claim.caseIds].sort()) !== JSON.stringify([...expected].sort()) || (req.caseIds.length && !expected.length)) fail(`Incomplete tested case set: ${claim.requirement}/${claim.profile}`);
    }
    for (const id of claim.caseIds) if (!req.caseIds.includes(id)) fail(`Claim case outside requirement: ${id}`);
    for (const operation of claim.operations) {
      if (!operations.includes(operation)) fail(`Invalid operation: ${operation}`);
      const key = `${claim.requirement}/${claim.profile}/${operation}`;
      if (cells.has(key)) fail(`Duplicate claim cell: ${key}`);
      cells.add(key);
    }
  }
  return matrix;
}
export function renderMatrix(matrix) {
  const clean = value => value.replaceAll('|', '\\|').replaceAll('\n', ' ');
  const lines = ['# Quality evidence matrix', '', '<!-- Generated by pnpm qa:matrix --write. Do not edit manually. -->', '', matrix.scope, '',
    'This inventory checks references and declared coverage; it does not execute the referenced tests or attest to their success. CI executes those gates separately. Counts are not an OOXML conformance percentage.', '',
    '`T` = tested bounded claim; `U` = untested; `O` = outside/unsupported in this QA profile. Every omitted cell expands to U. Consumer write means import/export of library-saved files; oracle read/write means validation of inputs/outputs. Reject describes the named subject only.', '',
    '## Requirements', '', '| Requirement | Normative reference / contract | Corpus cases |', '| --- | --- | --- |'];
  for (const req of matrix.requirements) lines.push(`| ${req.id} | ${clean(req.clauses.join('; '))}: ${clean(req.contract)} | ${req.caseIds.length} |`);
  for (const profile of matrix.profiles) {
    lines.push('', `## ${profile.id}`, '', `Subject: ${profile.subject}; execution: ${profile.execution}. Runtime: ${clean(profile.runtimes.join('; '))}. Adapter: ${clean(profile.adapters.join('; '))}.`, '', '| Requirement | Read | Write | Preserve | Reject |', '| --- | --- | --- | --- | --- |');
    for (const req of matrix.requirements) {
      const cells = operations.map(operation => {
        const claim = matrix.claims.find(c => c.profile === profile.id && c.requirement === req.id && c.operations.includes(operation));
        return ({ tested: 'T', unsupported: 'O', untested: 'U' })[claim?.status ?? matrix.defaultStatus];
      });
      lines.push(`| ${req.id} | ${cells.join(' | ')} |`);
    }
    lines.push('', `Profile evidence: ${profile.evidence.map(e => `[${e.path}](../../${e.path}) (anchor: ${clean(e.anchor)})`).join('; ')}`, '', 'Bounded claims and evidence:', '');
    const claims = matrix.claims.filter(c => c.profile === profile.id);
    if (!claims.length) lines.push('No tested claims. Every operation remains untested.');
    for (const claim of claims) {
      const links = claim.evidence.map(e => `[${e.path}](../../${e.path}) (anchor: ${clean(e.anchor)})`).join('; ');
      lines.push(`- **${claim.requirement} / ${claim.operations.join(', ')}: ${claim.status}**${claim.caseIds.length ? ` (${claim.caseIds.length} selected cases)` : ''}. ${clean(claim.scope)} ${links}`);
    }
  }
  return `${lines.join('\n')}\n`;
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const matrix = validateMatrix(JSON.parse(read(matrixPath)), JSON.parse(read(corpusPath)));
  const generated = renderMatrix(matrix);
  if (process.argv[2] === '--write') writeFileSync(reportPath, generated);
  else {
    if (process.argv[2]) fail('Usage: qa:matrix [--write]');
    if (read(reportPath) !== generated) fail('Coverage report drift; run pnpm qa:matrix --write');
  }
  console.info(`Quality matrix: ${matrix.requirements.length} requirements, ${matrix.profiles.length} profiles; references and report checked (not a conformance score)`);
}
