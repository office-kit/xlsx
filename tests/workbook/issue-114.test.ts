// Regression for https://github.com/office-kit/xlsx/issues/114 — a shared
// string built from `<r>` runs was written back as a single plain `<t>`,
// losing every run and its `<rPr>` formatting. The loader flattened rich-text
// `<si>` entries into their concatenated text before the worksheet reader ever
// saw them, so a cell with a bold first half came back uniformly unformatted.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load';
import { fromBuffer } from '../../src/io/node';
import { workbookToBytes } from '../../src/io/save';
import { getCell } from '../../src/worksheet/worksheet';
import { openZip } from '../../src/zip/reader';
import { validateXlsx } from '../conformance/validate';

const enc = new TextEncoder();
const dec = new TextDecoder();

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

// The reporter's fixture: one `<si>` of two runs, the first bold. B1 points at
// the same entry so the writer's dedup is exercised too.
const SHARED_STRINGS =
  `${XML_DECL}<sst xmlns="${MAIN_NS}" count="2" uniqueCount="1"><si>` +
  '<r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t>fett</t></r>' +
  '<r><t xml:space="preserve"> normal</t></r>' +
  '</si></sst>';

const buildFixture = (): Uint8Array =>
  zipSync({
    '[Content_Types].xml': enc.encode(
      `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>` +
        '</Relationships>',
    ),
    'xl/workbook.xml': enc.encode(
      `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
        '<sheets><sheet name="Tabelle1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId8" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>` +
        `<Relationship Id="rId9" Type="${REL_NS}/styles" Target="styles.xml"/>` +
        '</Relationships>',
    ),
    'xl/styles.xml': enc.encode(
      `${XML_DECL}<styleSheet xmlns="${MAIN_NS}">` +
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
        '</styleSheet>',
    ),
    'xl/sharedStrings.xml': enc.encode(SHARED_STRINGS),
    'xl/worksheets/sheet1.xml': enc.encode(
      `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData><row r="1">` +
        '<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c>' +
        '</row></sheetData></worksheet>',
    ),
  });

const roundTrip = async (bytes: Uint8Array = buildFixture()): Promise<Uint8Array> =>
  workbookToBytes(await loadWorkbook(fromBuffer(bytes)));

const partOf = async (bytes: Uint8Array, path: string): Promise<string> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    return dec.decode(archive.read(path));
  } finally {
    archive.close();
  }
};

describe('issue #114 — rich-text shared strings keep their runs', () => {
  it('loads a rich <si> as a rich-text cell value', async () => {
    const wb = await loadWorkbook(fromBuffer(buildFixture()));
    const sheet = wb.sheets[0]?.sheet;
    if (sheet === undefined || !('rows' in sheet)) throw new Error('expected a worksheet');
    expect(getCell(sheet, 1, 1)?.value).toEqual({
      kind: 'rich-text',
      runs: [
        { text: 'fett', font: { b: true, sz: 11, name: 'Calibri' } },
        { text: ' normal' },
      ],
    });
  });

  it('writes the runs back into sharedStrings.xml instead of flattening them', async () => {
    const sst = await partOf(await roundTrip(), 'xl/sharedStrings.xml');
    expect(sst).toContain('<r><rPr><rFont val="Calibri"/><b/><sz val="11"/></rPr><t>fett</t></r>');
    expect(sst).toContain('<r><t xml:space="preserve"> normal</t></r>');
    expect(sst).not.toContain('<t>fett normal</t>');
  });

  it('dedupes a rich string shared by two cells into one <si>', async () => {
    const out = await roundTrip();
    expect(await partOf(out, 'xl/sharedStrings.xml')).toContain('uniqueCount="1"');
    const sheet = await partOf(out, 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c>');
  });

  it('is byte-stable across a second round-trip and validates', async () => {
    const once = await roundTrip();
    const twice = await roundTrip(once);
    expect(await partOf(twice, 'xl/sharedStrings.xml')).toEqual(await partOf(once, 'xl/sharedStrings.xml'));
    expect((await validateXlsx(twice)).issues).toEqual([]);
  });
});
