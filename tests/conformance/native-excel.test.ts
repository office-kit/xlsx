import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { getSheet } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';
import { projectExcelFeatures, type Features } from './excel-feature-projection.js';
import { required } from './required.js';
import { validateXlsx } from './validate.js';
import { children, descendants, parseDocument } from './xml-tree.js';

it('preserves the pinned native Mac Excel feature output through another library edit', async () => {
  const provenance = JSON.parse(readFileSync(new URL('./fixtures/excel-mac/manifest.json', import.meta.url), 'utf8')) as { fixtureSha256: string };
  const expected = JSON.parse(readFileSync(new URL('./corpus/excel-features-manifest.json', import.meta.url), 'utf8')) as { cases: Array<{ expectedFeatures: Features }> };
  const bytes = readFileSync(new URL('./fixtures/excel-mac/feature-roundtrip.xlsx', import.meta.url));
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(provenance.fixtureSha256);
  expect(await validateXlsx(bytes)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  expect(projectExcelFeatures(bytes)).toEqual(required(expected.cases[0]).expectedFeatures);
  const wb = await loadWorkbook(fromBuffer(bytes));
  const sheet = required(getSheet(wb, 'Audit'));
  expect(getCell(sheet, 2, 2)?.value).toBe('edited');
  setCell(sheet, 2, 2, 'edited again');
  const saved = await workbookToBytes(wb);
  expect(await validateXlsx(saved)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  expect(projectExcelFeatures(saved)).toEqual(required(expected.cases[0]).expectedFeatures);
  const styles = parseDocument(strFromU8(required(unzipSync(saved)['xl/styles.xml'])));
  const fonts = required([...descendants(styles)].find(n => n.local === 'fonts'));
  expect(children(fonts, 'font').map(font => font.children.filter(child => typeof child !== 'string').map(child => child.local))).toEqual([['sz', 'name'], ['sz', 'name', 'family', 'charset']]);
  const reopened = await loadWorkbook(fromBuffer(saved));
  expect(getCell(required(getSheet(reopened, 'Audit')), 2, 2)?.value).toBe('edited again');
  const output = process.env['QA_NATIVE_EXCEL_OUTPUT'];
  if (output) {
    mkdirSync(output, { recursive: true });
    writeFileSync(join(output, 'manifest.json'), JSON.stringify([{ id: 'native-excel-feature-regression' }]));
    writeFileSync(join(output, 'native-excel-feature-regression.output.xlsx'), saved);
  }
});
