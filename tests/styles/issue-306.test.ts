import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { assert, describe, expect, it } from 'vitest';
import { makeRichText } from '../../src/cell/rich-text.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { fromTree, toTree } from '../../src/schema/serialize.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { getCellFont, setBold, setCellFont } from '../../src/styles/cell-style.js';
import { addDxf, getDxfs } from '../../src/styles/differential.js';
import { makeFont } from '../../src/styles/fonts.js';
import { FontSchema } from '../../src/styles/fonts.schema.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { parseRichString, serializeRichString } from '../../src/workbook/shared-strings.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';
import { SHEET_MAIN_NS } from '../../src/xml/namespaces.js';
import { parseXml } from '../../src/xml/parser.js';
import { serializeXml } from '../../src/xml/serializer.js';
import { validateXlsx } from '../conformance/validate.js';

const parseFont = (scheme: string) =>
  fromTree(parseXml(`<font xmlns="${SHEET_MAIN_NS}"><name val="Arial"/>${scheme}</font>`), FontSchema);

async function fixture(): Promise<Uint8Array> {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Report');
  setCellFont(wb, setCell(ws, 1, 1, 42), makeFont({ name: 'Arial', scheme: 'major' }));
  setCell(ws, 2, 1, { kind: 'rich-text', runs: makeRichText([{ text: 'note', font: { name: 'Arial', scheme: 'major' } }]) });
  addDxf(wb.styles, { font: makeFont({ scheme: 'major' }) });
  const parts = unzipSync(await workbookToBytes(wb));
  for (const name of ['xl/styles.xml', 'xl/sharedStrings.xml']) {
    const part = parts[name];
    assert(part, `expected ${name}`);
    const xml = strFromU8(part);
    expect(xml).toContain('<scheme val="major"/>');
    parts[name] = strToU8(xml.replaceAll('<scheme val="major"/>', '<scheme val="none"/>'));
  }
  return zipSync(parts);
}

describe('issue #306 — fonts opted out of the theme with scheme none', () => {
  it.each([
    ['', undefined],
    ['<scheme val="none"/>', 'none'],
    ['<scheme val="major"/>', 'major'],
    ['<scheme val="minor"/>', 'minor'],
  ] as const)('preserves the meaning of %s', (xml, expected) => {
    const font = parseFont(xml);
    expect(font).toEqual({ name: 'Arial', ...(expected === undefined ? {} : { scheme: expected }) });
    const saved = serializeXml(toTree(makeFont(font), FontSchema));
    expect(fromTree(parseXml(saved), FontSchema)).toEqual(font);
  });

  it.each(['None', 'body', ''])('still rejects invalid scheme value %j', (value) => {
    expect(() => parseFont(`<scheme val="${value}"/>`)).toThrow(OpenXmlSchemaError);
  });

  it('validates a rich-text run scheme the way the stylesheet does', () => {
    const run = (scheme: string) => parseXml(`<si xmlns="${SHEET_MAIN_NS}"><r><rPr><scheme val="${scheme}"/></rPr><t>x</t></r></si>`);
    expect(parseRichString(run('none'))).toEqual({ kind: 'rich-text', runs: [{ text: 'x', font: { scheme: 'none' } }] });
    expect(() => parseRichString(run('body'))).toThrow(OpenXmlSchemaError);
  });

  it('escapes run attributes that reach the writer unvalidated', () => {
    const si = parseXml(`<si xmlns="${SHEET_MAIN_NS}"><r><rPr><u val="a&quot;/&gt;&lt;b/&gt;"/></rPr><t>x</t></r></si>`);
    const xml = serializeRichString(parseRichString(si));
    expect(xml).toContain('<u val="a&quot;/&gt;&lt;b/&gt;"/>');
    expect(parseRichString(parseXml(`<si xmlns="${SHEET_MAIN_NS}">${xml}</si>`))).toEqual(parseRichString(si));
  });

  it('loads, edits and saves cell, rich-text and differential fonts without losing none', async () => {
    const bytes = await fixture();
    expect((await validateXlsx(bytes)).issues).toEqual([]);
    const wb = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb.sheets[0];
    assert(sheet?.kind === 'worksheet', 'expected report worksheet');
    const cell = getCell(sheet.sheet, 1, 1);
    assert(cell, 'expected report cell');
    expect(getCellFont(wb, cell).scheme).toBe('none');
    expect(getDxfs(wb.styles)[0]?.font?.scheme).toBe('none');
    expect(getCell(sheet.sheet, 2, 1)?.value).toEqual({ kind: 'rich-text', runs: [{ text: 'note', font: { name: 'Arial', scheme: 'none' } }] });
    setBold(wb, cell);
    expect(getCellFont(wb, cell)).toMatchObject({ name: 'Arial', bold: true, scheme: 'none' });
    const saved = await workbookToBytes(wb);
    expect((await validateXlsx(saved)).issues).toEqual([]);
    const savedStrings = unzipSync(saved)['xl/sharedStrings.xml'];
    assert(savedStrings, 'expected saved shared strings');
    expect(strFromU8(savedStrings)).toContain('<scheme val="none"/>');
    const back = await loadWorkbook(fromBuffer(saved));
    expect(back.styles.fonts).toEqual(wb.styles.fonts);
    expect(getDxfs(back.styles)).toEqual(getDxfs(wb.styles));
  });

  it('loads the same fonts through the streaming reader', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await fixture()));
    try {
      expect(wb.styles.fonts.some((font) => font.scheme === 'none')).toBe(true);
      expect(getDxfs(wb.styles)[0]?.font?.scheme).toBe('none');
      const rows: unknown[][] = [];
      for await (const row of wb.openWorksheet('Report').iterValues()) rows.push(row);
      expect(rows[0]).toEqual([42]);
    } finally {
      await wb.close();
    }
  });
});
