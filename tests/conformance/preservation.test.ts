import { readFileSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { setCell } from '../../src/worksheet/worksheet.js';
import { packageFor, SML } from './corpus.js';
import { preservationProjection, type PreservationContract } from './preservation.js';
import { required } from './required.js';
const base = '../../reference/openpyxl/openpyxl/';
const fixtures = [
  { file: 'tests/data/reader/vba+comments.xlsm', prefixes: ['xl/vbaProject.bin', 'xl/ctrlProps/', 'xl/printerSettings/'], types: ['vbaProject', 'ctrlProp', 'printerSettings', 'drawing', 'vmlDrawing'] },
  { file: 'reader/tests/data/legacy_drawing.xlsm', prefixes: ['xl/ctrlProps/', 'xl/drawings/vmlDrawing2.vml'], types: ['ctrlProp', 'drawing', 'vmlDrawing'], sources: ['xl/worksheets/sheet2.xml'] },
  { file: 'workbook/external_link/tests/data/book1.xlsx', prefixes: ['xl/externalLinks/'], types: ['externalLink', 'externalLinkPath'] },
];
function extensionPackage(): Uint8Array {
  const parts = unzipSync(packageFor({ id: 'preservation', clause: 'ECMA-376 parts 2/3' }));
  const path = 'xl/worksheets/sheet1.xml';
  parts[path] = strToU8(strFromU8(required(parts[path])).replace('</worksheet>', '<extLst><ext uri="urn:qa:extension"><qa:payload xmlns:qa="urn:qa:extension" token="keep">unknown data</qa:payload></ext></extLst></worksheet>'));
  parts['xl/styles.xml'] = strToU8(strFromU8(required(parts['xl/styles.xml'])).replace('</styleSheet>', '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="urn:qa:style"><qa:payload xmlns:qa="urn:qa:extension">style payload</qa:payload></ext></extLst></styleSheet>'));
  return zipSync(parts);
}
async function edit(input: Uint8Array): Promise<Uint8Array> {
  const wb = await loadWorkbook(fromBuffer(input));
  const ws = required(wb.sheets.find(s => s.kind === 'worksheet'));
  if (!('rows' in ws.sheet)) throw new Error('Expected worksheet');
  setCell(ws.sheet, 100, 20, 'unrelated QA edit');
  return workbookToBytes(wb);
}
describe('independent package preservation contracts', () => {
  it.each(fixtures)('$file preserves bytes, content types, graph and XML bindings', async fixture => {
    const input = readFileSync(new URL(base + fixture.file, import.meta.url));
    const paths = Object.keys(unzipSync(input)).filter(p => !p.endsWith('/') && fixture.prefixes.some(prefix => p.startsWith(prefix)));
    expect(paths.length).toBeGreaterThan(0);
    const contract: PreservationContract = { parts: paths, relationshipTypes: fixture.types, ...('sources' in fixture ? { relationshipSources: fixture.sources } : {}) };
    expect(preservationProjection(await edit(input), contract)).toEqual(preservationProjection(input, contract));
  });
  it('preserves unknown extension subtrees and explicit empty style pools', async () => {
    const input = extensionPackage();
    const contract: PreservationContract = { parts: [], relationshipTypes: [], fragments: [
      { part: 'xl/worksheets/sheet1.xml', local: 'extLst', uri: SML },
      ...['dxfs', 'tableStyles', 'extLst'].map(local => ({ part: 'xl/styles.xml', local, uri: SML })),
    ] };
    expect(preservationProjection(await edit(input), contract)).toEqual(preservationProjection(input, contract));
    const damaged = unzipSync(await edit(input));
    damaged['xl/styles.xml'] = strToU8(strFromU8(required(damaged['xl/styles.xml'])).replace(/<extLst>[\s\S]*?<\/extLst>/, ''));
    expect(preservationProjection(zipSync(damaged), contract)).not.toEqual(preservationProjection(input, contract));
  });
  it('calibrates missing binary, changed content type, graph target', () => {
    const input = readFileSync(new URL(base + required(fixtures[0]).file, import.meta.url));
    const contract: PreservationContract = { parts: ['xl/vbaProject.bin'], relationshipTypes: ['vbaProject'] };
    const expected = preservationProjection(input, contract);
    for (const fault of ['binary', 'type', 'target']) {
      const parts = unzipSync(input);
      if (fault === 'binary') { const data = required(parts['xl/vbaProject.bin']); data[0] = required(data[0]) ^ 1; }
      if (fault === 'type') parts['[Content_Types].xml'] = strToU8(strFromU8(required(parts['[Content_Types].xml'])).replace('application/vnd.ms-office.vbaProject', 'application/octet-stream'));
      if (fault === 'target') parts['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(required(parts['xl/_rels/workbook.xml.rels'])).replace('vbaProject.bin', 'missing.bin'));
      if (fault === 'target') expect(() => preservationProjection(zipSync(parts), contract)).toThrow(/Missing relationship target/);
      else expect(preservationProjection(zipSync(parts), contract)).not.toEqual(expected);
    }
  });
});
