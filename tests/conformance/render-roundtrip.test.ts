import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { excelFeaturePackage } from './excel-features.js';
import { required } from './required.js';
import { validateXlsx } from './validate.js';
it('prepares independently assembled print and drawing controls for real PDF comparison', async () => {
  const parts = unzipSync(excelFeaturePackage());
  const change = (path: string, before: string, after: string) => {
    const xml = strFromU8(required(parts[path]));
    expect(xml.split(before)).toHaveLength(2);
    parts[path] = strToU8(xml.replace(before, after));
  };
  change('xl/styles.xml', '<font></font>', '<font><sz val="11"/><name val="Liberation Sans"/></font>');
  change('xl/worksheets/sheet1.xml', '<sheetData>', '<sheetFormatPr defaultRowHeight="15"/><sheetData>');
  change('xl/worksheets/sheet1.xml', '<dimension', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension');
  change('xl/workbook.xml', '</sheets>', '</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Audit!$A$1:$K$18</definedName></definedNames>');
  change('xl/charts/chart1.xml', '<c:chart>', '<c:style val="2"/><c:chart>');
  change('xl/drawings/drawing1.xml', '<a:ext cx="0" cy="0"/>', '<a:ext cx="5486400" cy="2095500"/>');
  change('xl/charts/chart1.xml', '<c:cat>', '<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr></c:dPt><c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr></c:dPt><c:cat>');
  const input = zipSync(parts);
  expect(await validateXlsx(input)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  const output = await workbookToBytes(await loadWorkbook(fromBuffer(input)));
  expect(await validateXlsx(output)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  const target = process.env['QA_RENDER_OUTPUT'];
  if (!target) return; // Normal unit tests verify the preparation without launching desktop software.
  const save = (id: string, bytes: Uint8Array) => {
    mkdirSync(join(target, id), { recursive: true });
    writeFileSync(join(target, id, 'sample.xlsx'), bytes);
  };
  save('input', input); save('output', output);
  change('xl/charts/chart1.xml', '4472C4', 'FF0000');
  save('changed-chart-color', zipSync(parts));
  change('xl/charts/chart1.xml', 'FF0000', '4472C4');
  change('xl/worksheets/sheet1.xml', 'orientation="landscape"', 'orientation="portrait"');
  save('changed-orientation', zipSync(parts));
  change('xl/worksheets/sheet1.xml', 'orientation="portrait"', 'orientation="landscape"');
  change('xl/worksheets/sheet1.xml', '<drawing r:id="drawing"/>', '');
  save('lost-drawing', zipSync(parts));
});
