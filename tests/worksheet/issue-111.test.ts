// Regression for https://github.com/office-kit/xlsx/issues/111 — a cell with
// no value and the default style was dropped on save. Dropping a genuinely
// unreferenced empty cell would be harmless, but an empty cell can still be
// the target of a `definedName`, a form control's `fmlaLink` or a drawing
// anchor; the reporter's file kept `<c r="B1"/>` as the target of a defined
// name, which dangled after a round-trip.
//
// `deleteCell` — not `setCell(…, null)` — is how a caller says a cell is gone.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load';
import { fromBuffer } from '../../src/io/node';
import { workbookToBytes } from '../../src/io/save';
import { openZip } from '../../src/zip/reader';
import { validateXlsx } from '../conformance/validate';

const enc = new TextEncoder();
const dec = new TextDecoder();

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

// The reporter's fixture: A1 carries a value, B1 is bare, C1 is bare with the
// default style, and a defined name points at B1.
const buildFixture = (): Uint8Array =>
  zipSync({
    '[Content_Types].xml': enc.encode(
      `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>` +
        '</Relationships>',
    ),
    'xl/workbook.xml': enc.encode(
      `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
        '<sheets><sheet name="Tabelle1" sheetId="1" r:id="rId1"/></sheets>' +
        '<definedNames><definedName name="Kontrollkaestchen2">Tabelle1!$B$1</definedName></definedNames>' +
        '</workbook>',
    ),
    'xl/_rels/workbook.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
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
    'xl/worksheets/sheet1.xml': enc.encode(
      `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData>` +
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"/><c r="C1" s="0"/></row>' +
        '</sheetData></worksheet>',
    ),
  });

const roundTrip = async (bytes: Uint8Array = buildFixture()): Promise<Uint8Array> =>
  workbookToBytes(await loadWorkbook(fromBuffer(bytes)));

const sheetXml = async (bytes: Uint8Array): Promise<string> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    return dec.decode(archive.read('xl/worksheets/sheet1.xml'));
  } finally {
    archive.close();
  }
};

describe('issue #111 — valueless cells survive a round-trip', () => {
  it('re-emits <c r="B1"/> so the defined name still resolves', async () => {
    const xml = await sheetXml(await roundTrip());
    expect(xml).toContain('<c r="A1"><v>1</v></c><c r="B1"/><c r="C1"/>');
  });

  it('does not invent cells that were never in the source', async () => {
    const xml = await sheetXml(await roundTrip());
    expect(xml).not.toContain('r="D1"');
    expect(xml).not.toContain('r="A2"');
  });

  it('is stable across a second round-trip and validates', async () => {
    const once = await roundTrip();
    const twice = await roundTrip(once);
    expect(await sheetXml(twice)).toEqual(await sheetXml(once));
    expect((await validateXlsx(twice)).issues).toEqual([]);
  });
});
