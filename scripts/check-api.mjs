// A conservative declaration-change gate, not an automatic compatibility proof.
import { createHash } from 'node:crypto';
import { execFileSync, spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';
const baselinePath = 'tests/consumer/api-baseline.json';
const approvalPath = 'tests/consumer/api-change.json';
const json = value => `${JSON.stringify(value, null, 2)}\n`;
export const digest = value => createHash('sha256').update(json(value)).digest('hex');
export function collectApi(root = resolve('.')) {
  const pkg = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));
  const entrypoints = Object.fromEntries(Object.entries(pkg.exports).sort(([a], [b]) => a.localeCompare(b)));
  const roots = Object.values(entrypoints).map(entry => resolve(root, entry.types));
  const program = ts.createProgram(roots, { module: ts.ModuleKind.NodeNext, target: ts.ScriptTarget.ES2022 });
  const printer = ts.createPrinter({ removeComments: true, newLine: ts.NewLineKind.LineFeed });
  const checker = program.getTypeChecker();
  const exports = {};
  for (const [name, entry] of Object.entries(entrypoints)) {
    const file = program.getSourceFile(resolve(root, entry.types));
    const symbol = file && checker.getSymbolAtLocation(file);
    if (!symbol) throw new Error(`Missing public declaration module: ${name}; build first`);
    exports[name] = checker.getExportsOfModule(symbol).map(s => s.name).sort();
    if (!exports[name].length) throw new Error(`Empty public module: ${name}`);
  }
  const declarations = Object.fromEntries(program.getSourceFiles()
    .filter(file => relative(root, file.fileName).replaceAll('\\', '/').startsWith('dist/'))
    .map(file => [relative(root, file.fileName).replaceAll('\\', '/'), printer.printFile(file)])
    .sort(([a], [b]) => a.localeCompare(b)));
  return { formatVersion: 1, entrypoints, exports, declarations };
}
export function validateApiChange(previous, current, approval, version, changeset) {
  if (digest(previous) === digest(current)) return;
  if (approval.baselineSha256 !== digest(current) || typeof approval.reason !== 'string' || !approval.reason.trim()) throw new Error('API change needs an explanation bound to the new baseline hash');
  const required = { additive: 'minor', correction: 'patch', breaking: version.startsWith('0.') ? 'minor' : 'major' }[approval.kind];
  if (!required) throw new Error('API change kind must be additive, correction or breaking');
  const bump = changeset.match(/^---\r?\n[\s\S]*?["']?@office-kit\/xlsx["']?\s*:\s*(patch|minor|major)\s*\r?\n[\s\S]*?---\r?\n([\s\S]+)$/);
  if (!bump || !bump[2].trim() || ['patch', 'minor', 'major'].indexOf(bump[1]) < ['patch', 'minor', 'major'].indexOf(required)) throw new Error(`API ${approval.kind} requires a ${required} or larger changeset`);
  if (approval.kind === 'breaking' && !/breaking/i.test(bump[2])) throw new Error('Breaking API changes must be flagged in the changeset');
}
function main() {
  const args = process.argv.slice(2);
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--write') continue;
    if (args[i] === '--base' && args[i + 1] && !args[i + 1].startsWith('-')) { i++; continue; }
    throw new Error('Usage: check:api [--write] [--base REF]');
  }
  const current = collectApi();
  if (args.includes('--write')) writeFileSync(baselinePath, json(current));
  const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
  if (json(current) !== json(baseline)) {
    const paths = new Set([...Object.keys(baseline.declarations), ...Object.keys(current.declarations)]);
    const changed = [...paths].filter(path => baseline.declarations[path] !== current.declarations[path]);
    throw new Error(`Published declaration drift: ${changed.join(', ') || 'exports map'}; review the API change and run pnpm check:api --write`);
  }
  const base = args.includes('--base') ? args[args.indexOf('--base') + 1] : process.env.QA_API_BASE_REF;
  if (args.includes('--base') && !base) throw new Error('Missing base ref');
  if (base) {
    execFileSync('git', ['rev-parse', '--verify', `${base}^{commit}`], { stdio: 'pipe' });
    const old = spawnSync('git', ['show', `${base}:${baselinePath}`], { encoding: 'utf8' });
    if (old.status !== 0) {
      // The initial gate PR bootstraps the baseline; later baseline updates require approval.
      const files = execFileSync('git', ['ls-tree', '-r', '--name-only', base], { encoding: 'utf8' });
      if (files.split('\n').includes(baselinePath)) throw new Error('Unable to read API baseline from base');
    } else if (digest(JSON.parse(old.stdout)) !== digest(current)) {
      const approval = JSON.parse(readFileSync(approvalPath, 'utf8'));
      if (!/^\.changeset\/[\w-]+\.md$/.test(approval.changeset)) throw new Error('API approval must name a changeset file');
      const added = execFileSync('git', ['diff', '--name-only', '--diff-filter=A', base, '--', '.changeset'], { encoding: 'utf8' }).split('\n');
      if (!added.includes(approval.changeset)) throw new Error('API approval requires a newly added changeset');
      validateApiChange(JSON.parse(old.stdout), current, approval, JSON.parse(readFileSync('package.json', 'utf8')).version, readFileSync(approval.changeset, 'utf8'));
    }
  }
  console.info(`Public API baseline: ${Object.keys(current.exports).length} subpaths, ${Object.keys(current.declarations).length} reachable declarations; hash ${digest(current)}`);
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
