import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, posix } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { getSheet, moveSheet, swapSheets } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';
import { excelFeaturePackage } from './excel-features.js';
import { projectExcelFeatures, type Features } from './excel-feature-projection.js';
import { required } from './required.js';
import { validateXlsx } from './validate.js';
import { attribute, children, parseDocument, textContent } from './xml-tree.js';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const expected = (JSON.parse(readFileSync(new URL('./corpus/excel-features-manifest.json', import.meta.url), 'utf8')) as { cases: Array<{ expectedFeatures: Features }> }).cases[0]?.expectedFeatures;
// Both sheets are assembled without the production writer, with distinct rule/formula/print ownership.
function input(): Uint8Array {
  const parts = unzipSync(excelFeaturePackage());
  const change = (path: string, transform: (xml: string) => string) => { parts[path] = strToU8(transform(strFromU8(required(parts[path])))); };
  const rules = (name: string) => `<mergeCells count="1"><mergeCell ref="C5:D6"/></mergeCells><conditionalFormatting sqref="A2:A3"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>${name === 'Audit' ? 0 : 2}</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="1" sqref="B4:B5"><formula1>0</formula1><formula2>${name === 'Audit' ? 10 : 20}</formula2></dataValidation></dataValidations>`;
  change('xl/worksheets/sheet1.xml', xml => xml.replace('<printOptions', `${rules('Audit')}<printOptions`));
  for (const path of ['xl/worksheets/sheet1.xml', 'xl/worksheets/_rels/sheet1.xml.rels', 'xl/tables/table1.xml', 'xl/charts/chart1.xml', 'xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels']) {
    parts[path.replace('1.xml', '2.xml')] = strToU8(strFromU8(required(parts[path])).replaceAll('1.xml', '2.xml').replaceAll('Audit', 'Other').replace('id="1" name="OtherTable"', 'id="2" name="OtherTable"'));
  }
  change('xl/worksheets/sheet2.xml', xml => xml.replace('<formula>0</formula>', '<formula>2</formula>').replace('<formula2>10</formula2>', '<formula2>20</formula2>').replace('orientation="landscape"', 'orientation="portrait"'));
  change('xl/workbook.xml', xml => xml.replace('</sheets>', '<sheet name="Other" sheetId="2" r:id="other"/></sheets>'));
  change('xl/_rels/workbook.xml.rels', xml => xml.replace('</Relationships>', `<Relationship Id="other" Type="${R}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`));
  change('[Content_Types].xml', xml => xml.replace('</Types>', [...xml.matchAll(/<Override[^>]+\/(?:sheet|table|chart|drawing)1\.xml[^>]+\/>/g)].map(match => match[0].replace('1.xml', '2.xml')).join('') + '</Types>'));
  return zipSync(parts);
}
function projectRules(bytes: Uint8Array, title: string) {
  const parts = unzipSync(bytes);
  const xml = (path: string) => parseDocument(strFromU8(required(parts[path])));
  const sheetId = attribute(required(children(required(children(xml('xl/workbook.xml'), 'sheets')[0]), 'sheet').find(n => attribute(n, 'name') === title)), 'id', R);
  const link = required(children(xml('xl/_rels/workbook.xml.rels'), 'Relationship').find(n => attribute(n, 'Id') === sheetId));
  const sheet = xml(posix.normalize(posix.join('xl', required(attribute(link, 'Target')))));
  return {
    merges: children(required(children(sheet, 'mergeCells')[0]), 'mergeCell').map(n => attribute(n, 'ref')),
    conditional: children(sheet, 'conditionalFormatting').map(n => ({ range: attribute(n, 'sqref'), rules: children(n, 'cfRule').map(rule => ({ type: attribute(rule, 'type'), operator: attribute(rule, 'operator'), priority: attribute(rule, 'priority'), formulas: children(rule, 'formula').map(textContent) })) })),
    validations: children(required(children(sheet, 'dataValidations')[0]), 'dataValidation').map(n => ({ type: attribute(n, 'type'), operator: attribute(n, 'operator'), range: attribute(n, 'sqref'), allowBlank: attribute(n, 'allowBlank'), minimum: textContent(required(children(n, 'formula1')[0])), maximum: textContent(required(children(n, 'formula2')[0])) })),
  };
}
function assertFeatures(bytes: Uint8Array) {
  for (const title of ['Audit', 'Other']) {
    const feature = structuredClone(required(expected));
    if (title === 'Other') {
      required(feature.tables[0]).name = 'OtherTable';
      required(feature.charts[0]).valueRef = 'Other!$A$2:$A$3';
      required(feature.charts[0]).categoryRef = 'Other!$C$2:$C$3';
      feature.pageSetup['orientation'] = 'portrait';
    }
    expect(projectExcelFeatures(bytes, title)).toEqual(feature);
    expect(projectRules(bytes, title)).toEqual({ merges: ['C5:D6'], conditional: [{ range: 'A2:A3', rules: [{ type: 'cellIs', operator: 'greaterThan', priority: '1', formulas: [title === 'Audit' ? '0' : '2'] }] }], validations: [{ type: 'whole', operator: 'between', range: 'B4:B5', allowBlank: '1', minimum: '0', maximum: title === 'Audit' ? '10' : '20' }] });
  }
}
it('preserves per-sheet tables, drawings, rules and print setup through edits, moves and swaps', async () => {
  const original = input();
  expect(await validateXlsx(original)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  assertFeatures(original);
  const wb = await loadWorkbook(fromBuffer(original));
  setCell(required(getSheet(wb, 'Audit')), 2, 2, 'audit-edit');
  setCell(required(getSheet(wb, 'Other')), 2, 2, 'other-edit');
  const output = process.env['QA_FEATURE_INTERACTIONS_OUTPUT'];
  if (output) {
    mkdirSync(output, { recursive: true });
    writeFileSync(join(output, 'manifest.json'), JSON.stringify(['edit', 'move', 'swap'].map(id => ({ id }))));
  }
  for (const [id, operation] of [
    ['edit', () => {}], ['move', () => moveSheet(wb, 'Other', 0)], ['swap', () => swapSheets(wb, 'Other', 'Audit')],
  ] as const) {
    operation();
    const saved = await workbookToBytes(wb);
    expect(await validateXlsx(saved)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
    assertFeatures(saved);
    const reopened = await loadWorkbook(fromBuffer(saved));
    for (const title of ['Audit', 'Other']) expect(getCell(required(getSheet(reopened, title)), 2, 2)?.value).toBe(`${title.toLowerCase()}-edit`);
    if (output) writeFileSync(join(output, `${id}.output.xlsx`), saved);
  }
});
it.each(['formula', 'orientation', 'relationship'])('independent per-sheet oracle detects a changed %s', fault => {
  const parts = unzipSync(input());
  const path = fault === 'relationship' ? 'xl/worksheets/_rels/sheet2.xml.rels' : 'xl/worksheets/sheet2.xml';
  const xml = strFromU8(required(parts[path]));
  parts[path] = strToU8(fault === 'formula' ? xml.replace('<formula>2</formula>', '<formula>0</formula>') : fault === 'orientation' ? xml.replace('orientation="portrait"', 'orientation="landscape"') : xml.replace('drawing2.xml', 'drawing1.xml'));
  expect(() => assertFeatures(zipSync(parts))).toThrow();
});
