import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { expect, it } from 'vitest';
import { packageFor } from './corpus.js';
import { required } from './required.js';
import { validateXlsx } from './validate.js';

// Hand-authored numeric, ISO date and boolean cells. Namespace substitution
// constructs the independent input; production Strict normalization is unused.
function strictPackage(): Uint8Array {
  const parts = unzipSync(packageFor({ id: 'strict-oracle', clause: 'ECMA-376-1 ST_CellType', rowsXml: '<row r="1"><c r="A1"><v>42</v></c><c r="B1" t="d"><v>2024-02-29T12:00:00Z</v></c><c r="C1" t="b"><v>1</v></c></row>' }));
  for (const [path, bytes] of Object.entries(parts)) parts[path] = strToU8(strFromU8(bytes)
    .replaceAll('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main')
    .replaceAll('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'http://purl.oclc.org/ooxml/officeDocument/relationships'));
  return zipSync(parts);
}
it('independently validates Strict input with the normative Strict schemas', async () => {
  const bytes = strictPackage();
  expect(await validateXlsx(bytes, { conformance: 'strict' })).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  expect((await validateXlsx(bytes)).status).toBe('invalid');
});
it.each([
  ['invalid cell type', 't="b"', 't="unknown"', 'xsd'],
  ['missing shared string', 't="b"', 't="s"', 'semantic'],
  ['invalid style index', 'r="A1"', 'r="A1" s="99"', 'semantic'],
] as const)('Strict oracle rejects %s without normalization', async (_label, before, after, tier) => {
  const parts = unzipSync(strictPackage());
  const path = 'xl/worksheets/sheet1.xml';
  parts[path] = strToU8(strFromU8(required(parts[path])).replace(before, after));
  const result = await validateXlsx(zipSync(parts), { conformance: 'strict' });
  expect(result.status).toBe('invalid');
  expect(result.issues.some(i => i.tier === tier && i.part === path)).toBe(true);
});
it('pins unmodified official Strict schema bytes', () => {
  const manifest = JSON.parse(readFileSync(new URL('./schemas/strict/provenance.json', import.meta.url), 'utf8')) as { files: Record<string, string> };
  expect(Object.keys(manifest.files)).toHaveLength(21);
  for (const [file, hash] of Object.entries(manifest.files)) expect(createHash('sha256').update(readFileSync(new URL(`./schemas/strict/${file}`, import.meta.url))).digest('hex')).toBe(hash);
});
