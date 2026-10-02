import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { getSheet } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';
import { excelFeaturePackage } from './excel-features.js';
import { required } from './required.js';
import { validateXlsx } from './validate.js';
import { projectExcelFeatures, type Features } from './excel-feature-projection.js';
const manifest = JSON.parse(readFileSync(new URL('./corpus/excel-features-manifest.json', import.meta.url), 'utf8')) as { cases: Array<{ id: string; expectedFeatures: Features }> };
it('independently preserves attached table, chart, image and print configuration after an unrelated edit', async () => {
  const c = required(manifest.cases[0]);
  const bytes = excelFeaturePackage();
  expect(await validateXlsx(bytes)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  expect(projectExcelFeatures(bytes)).toEqual(c.expectedFeatures);
  const wb = await loadWorkbook(fromBuffer(bytes));
  setCell(required(getSheet(wb, 'Audit')), 2, 2, 'edited');
  const saved = await workbookToBytes(wb);
  expect(await validateXlsx(saved)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  expect(projectExcelFeatures(saved)).toEqual(c.expectedFeatures);
  const output = process.env['QA_EXCEL_FEATURES_OUTPUT'];
  if (output) {
    mkdirSync(output, { recursive: true });
    writeFileSync(join(output, 'manifest.json'), JSON.stringify(manifest.cases, null, 2));
    writeFileSync(join(output, `${c.id}.input.xlsx`), bytes);
    writeFileSync(join(output, `${c.id}.output.xlsx`), saved);
  }
});
