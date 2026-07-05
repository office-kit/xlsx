// Round-trip genuine xlsx fixtures from third-party producers (Excel,
// LibreOffice, openpyxl) and assert the bytes we re-emit are still ECMA-376
// compliant. This catches a different class of bugs than the writer survey:
// it exposes the parser/serializer to markup variants @office-kit/xlsx didn't author
// — extension namespaces, alternate ordering, optional flags — and makes
// sure the round-trip preserves OPC validity.

import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { validateXlsx } from './validate';

const HERE = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(HERE, '../../reference/openpyxl/openpyxl/tests/data/genuine');

const dump = (issues: { tier: string; part: string; message: string }[]): string =>
  issues.map((i) => `[${i.tier}] ${i.part}: ${i.message}`).join('\n');

const fixtures: ReadonlyArray<{ file: string; note?: string }> = [
  { file: 'empty.xlsx' },
  { file: 'empty-with-styles.xlsx' },
  { file: 'sample.xlsx' },
  { file: 'mac_date.xlsx' },
  { file: 'libreoffice_nrt.xlsx', note: 'LibreOffice output' },
];

describe('conformance: genuine fixtures round-trip schema-clean', () => {
  for (const { file, note } of fixtures) {
    it(`${file}${note ? ` (${note})` : ''}`, async () => {
      const original = readFileSync(resolve(FIXTURES, file));
      const wb = await loadWorkbook(fromBuffer(original));
      const reEmitted = await workbookToBytes(wb);
      const result = await validateXlsx(reEmitted);
      expect(result.issues, dump(result.issues)).toEqual([]);
    });
  }

  // Belt-and-braces: the *original* bytes from real producers must also pass
  // our validator after MC stripping. If they don't, the validator is
  // overconstrained. (We expect Excel/LibreOffice output to be schema-clean.)
  for (const { file, note } of fixtures) {
    it(`${file}${note ? ` (${note})` : ''} — original bytes pass directly`, async () => {
      const original = readFileSync(resolve(FIXTURES, file));
      const result = await validateXlsx(original);
      expect(result.issues, dump(result.issues)).toEqual([]);
    });
  }
});
