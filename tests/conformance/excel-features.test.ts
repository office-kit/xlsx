import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, posix } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { getSheet } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';
import { excelFeaturePackage } from './excel-features.js';
import { required } from './required.js';
import { validateXlsx } from './validate.js';
import { attribute, children, descendants, parseDocument, textContent, type XmlElement } from './xml-tree.js';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
interface Features {
  tables: Array<{ name: string; ref: string; columns: string[]; autoFilter: string }>;
  charts: Array<{ type: string; valueRef: string; categoryRef: string; values: number[]; categories: string[] }>;
  images: string[]; anchors: Array<[string, number, number, number, number]>;
  pageSetup: Record<string, string>; pageMargins: Record<string, number>; printOptions: Record<string, string>;
}
const manifest = JSON.parse(readFileSync(new URL('./corpus/excel-features-manifest.json', import.meta.url), 'utf8')) as { cases: Array<{ id: string; expectedFeatures: Features }> };
/** Independent XML/ZIP projection follows attached relationships, not filenames. */
function project(bytes: Uint8Array): Features {
  const parts = unzipSync(bytes);
  const xml = (path: string) => parseDocument(strFromU8(required(parts[path])));
  const target = (path: string, id: string, kind: string): string => {
    const rels = xml(posix.join(posix.dirname(path), '_rels', `${posix.basename(path)}.rels`));
    const rel = required(children(rels, 'Relationship').find(n => attribute(n, 'Id') === id));
    expect(attribute(rel, 'Type')).toBe(`${R}/${kind}`);
    expect(attribute(rel, 'TargetMode')).not.toBe('External');
    const to = required(attribute(rel, 'Target'));
    return posix.normalize(to.startsWith('/') ? to.slice(1) : posix.join(posix.dirname(path), to));
  };
  const sheetPath = target('xl/workbook.xml', required(attribute(required(children(required(children(xml('xl/workbook.xml'), 'sheets')[0]), 'sheet')[0]), 'id', R)), 'worksheet');
  const sheet = xml(sheetPath);
  const tables = children(required(children(sheet, 'tableParts')[0]), 'tablePart').map(n => {
    const table = xml(target(sheetPath, required(attribute(n, 'id', R)), 'table'));
    return { name: required(attribute(table, 'name')), ref: required(attribute(table, 'ref')), columns: children(required(children(table, 'tableColumns')[0]), 'tableColumn').map(column => required(attribute(column, 'name'))), autoFilter: required(attribute(required(children(table, 'autoFilter')[0]), 'ref')) };
  });
  const drawingPath = target(sheetPath, required(attribute(required(children(sheet, 'drawing')[0]), 'id', R)), 'drawing');
  const drawing = xml(drawingPath);
  const charts: Features['charts'] = []; const images: string[] = []; const anchors: Features['anchors'] = [];
  for (const anchor of children(drawing, 'twoCellAnchor')) {
    const nodes = [...descendants(anchor)];
    const chartLink = nodes.find(n => n.uri === C && n.local === 'chart');
    const imageLink = nodes.find(n => n.uri === A && n.local === 'blip');
    const marker = (tag: string, coordinate: string) => Number(textContent(required(children(required(children(anchor, tag)[0]), coordinate)[0])));
    anchors.push([chartLink ? 'chart' : 'image', marker('from', 'col'), marker('from', 'row'), marker('to', 'col'), marker('to', 'row')]);
    if (chartLink) {
      const chart = xml(target(drawingPath, required(attribute(chartLink, 'id', R)), 'chart'));
      const pie = required([...descendants(chart)].find(n => n.uri === C && n.local === 'pieChart'));
      const series = required(children(pie, 'ser')[0]);
      const values = required(children(required(children(series, 'val')[0]), 'numRef')[0]);
      const categories = required(children(required(children(series, 'cat')[0]), 'strRef')[0]);
      const cache = (node: XmlElement, name: string) => children(required(children(node, name)[0]), 'pt').map(n => textContent(required(children(n, 'v')[0])));
      charts.push({ type: pie.local, valueRef: textContent(required(children(values, 'f')[0])), categoryRef: textContent(required(children(categories, 'f')[0])), values: cache(values, 'numCache').map(Number), categories: cache(categories, 'strCache') });
    }
    if (imageLink) images.push(createHash('sha256').update(required(parts[target(drawingPath, required(attribute(imageLink, 'embed', R)), 'image')])).digest('hex'));
  }
  const attributes = (tag: string, keys: string[]) => Object.fromEntries(keys.map(key => [key, required(attribute(required(children(sheet, tag)[0]), key))]));
  return { tables, charts, images, anchors, pageSetup: attributes('pageSetup', ['paperSize', 'orientation', 'fitToWidth', 'fitToHeight']), pageMargins: Object.fromEntries(Object.entries(attributes('pageMargins', ['left', 'right', 'top', 'bottom', 'header', 'footer'])).map(([k, v]) => [k, Number(v)])), printOptions: attributes('printOptions', ['horizontalCentered']) };
}
it('independently preserves attached table, chart, image and print configuration after an unrelated edit', async () => {
  const c = required(manifest.cases[0]);
  const bytes = excelFeaturePackage();
  expect(await validateXlsx(bytes)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  expect(project(bytes)).toEqual(c.expectedFeatures);
  const wb = await loadWorkbook(fromBuffer(bytes));
  setCell(required(getSheet(wb, 'Audit')), 2, 2, 'edited');
  const saved = await workbookToBytes(wb);
  expect(await validateXlsx(saved)).toMatchObject({ status: 'valid', issues: [], skipped: [] });
  expect(project(saved)).toEqual(c.expectedFeatures);
  const output = process.env['QA_EXCEL_FEATURES_OUTPUT'];
  if (output) {
    mkdirSync(output, { recursive: true });
    writeFileSync(join(output, 'manifest.json'), JSON.stringify(manifest.cases, null, 2));
    writeFileSync(join(output, `${c.id}.input.xlsx`), bytes);
    writeFileSync(join(output, `${c.id}.output.xlsx`), saved);
  }
});
