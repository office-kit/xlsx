// Independent application check; process exit alone is not a semantic pass.
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { pathToFileURL } from 'node:url';
const corpus = resolve('.qa/corpus');
const output = resolve('.qa/libreoffice');
const profile = resolve('.qa/libreoffice-profile');
// Remove only this runner's output/profile; stale files cannot satisfy a new run.
rmSync(output, { recursive: true, force: true });
rmSync(profile, { recursive: true, force: true });
mkdirSync(output, { recursive: true });
const cases = JSON.parse(readFileSync(join(corpus, 'manifest.json'), 'utf8'));
if (!cases.length) throw new Error('Empty corpus');
const files = cases.map(c => join(corpus, `${c.id}.output.xlsx`));
const command = process.env['SOFFICE'] ?? 'soffice';
const result = spawnSync(command, [`-env:UserInstallation=${pathToFileURL(profile).href}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', output, ...files], { encoding: 'utf8', timeout: 180_000, maxBuffer: 8 * 1024 * 1024 });
if (result.error) throw result.error;
if (result.status !== 0) throw new Error(`LibreOffice failed: ${result.status}\n${result.stderr}`);
const expectedFiles = new Set(files.map(f => f.split('/').at(-1)));
const actual = new Set(readdirSync(output));
for (const file of expectedFiles) if (!actual.has(file)) throw new Error(`LibreOffice produced no output for ${file}`);
// Python's XML/ZIP stack reads application results without invoking this library.
const check = spawnSync('python3', ['scripts/qa-office-check.py', corpus, output], { encoding: 'utf8', timeout: 30_000 });
if (check.error) throw check.error;
writeFileSync(join(corpus, 'libreoffice-log.txt'), result.stdout + result.stderr + check.stdout + check.stderr);
process.stdout.write(check.stdout); process.stderr.write(check.stderr);
process.exitCode = check.status ?? 1;
