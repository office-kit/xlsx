// Regression for https://github.com/office-kit/xlsx/issues/116 — the
// markup-compatibility header Excel writes on `<workbook>` was destroyed on
// save: `mc:Ignorable` and the namespace declarations it names disappeared
// from the root, the captured children came back on generated prefixes
// (`x15ac:` → `ns0:`, `xr:` → `x16:`), `<workbookPr>` moved behind the
// `mc:AlternateContent` block and the `xr2:uid` on `<workbookView>` was
// dropped. Excel then reported the file as corrupt.
//
// This is the workbook-level sibling of #105, which fixed the same rewrite on
// the worksheet.

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
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const X15_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const X15AC_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac';
const XR_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';
const XR2_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2';
const XR6_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6';
const XR10_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10';

// The block Excel writes into essentially every workbook, reduced to one sheet.
const WORKBOOK_XML =
  `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}" xmlns:mc="${MC_NS}" ` +
  `mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="${X15_NS}" xmlns:xr="${XR_NS}" ` +
  `xmlns:xr6="${XR6_NS}" xmlns:xr10="${XR10_NS}" xmlns:xr2="${XR2_NS}">` +
  '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="20415"/>' +
  '<workbookPr defaultThemeVersion="166925"/>' +
  `<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Choice Requires="x15">` +
  `<x15ac:absPath url="C:\\Vorlagen\\" xmlns:x15ac="${X15AC_NS}"/>` +
  '</mc:Choice></mc:AlternateContent>' +
  '<xr:revisionPtr revIDLastSave="0" documentId="13_ncr:1_{00000000-0000-0000-0000-000000000001}" ' +
  'xr6:coauthVersionLast="36" xr6:coauthVersionMax="36" ' +
  'xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>' +
  '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="12000" ' +
  'xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/></bookViews>' +
  '<sheets><sheet name="Tabelle1" sheetId="1" r:id="rId1"/></sheets>' +
  '<calcPr calcId="191029"/>' +
  `<extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="${X15_NS}">` +
  '<x15:workbookPr chartTrackingRefBase="1"/></ext></extLst>' +
  '</workbook>';

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
    'xl/workbook.xml': enc.encode(WORKBOOK_XML),
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
      `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData/></worksheet>`,
    ),
  });

const roundTrip = async (bytes: Uint8Array = buildFixture()): Promise<Uint8Array> =>
  workbookToBytes(await loadWorkbook(fromBuffer(bytes)));

const workbookXml = async (bytes: Uint8Array): Promise<string> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    return dec.decode(archive.read('xl/workbook.xml'));
  } finally {
    archive.close();
  }
};

describe("issue #116 — workbook.xml keeps Excel's markup-compatibility header", () => {
  it('keeps mc:Ignorable and every prefix it names, declared on the root', async () => {
    const xml = await workbookXml(await roundTrip());
    const start = xml.indexOf('<workbook ');
    const root = xml.slice(start, xml.indexOf('>', start) + 1);
    expect(root).toContain('mc:Ignorable="x15 xr xr6 xr10 xr2"');
    for (const [prefix, ns] of [
      ['mc', MC_NS],
      ['x15', X15_NS],
      ['xr', XR_NS],
      ['xr6', XR6_NS],
      ['xr10', XR10_NS],
      ['xr2', XR2_NS],
    ]) {
      expect(root).toContain(`xmlns:${prefix}="${ns}"`);
    }
  });

  it('keeps the child prefixes rather than generating new ones', async () => {
    const xml = await workbookXml(await roundTrip());
    expect(xml).toContain('<xr:revisionPtr');
    expect(xml).toContain('xr6:coauthVersionLast="36"');
    expect(xml).toContain('xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"');
    expect(xml).toContain('<x15ac:absPath url="C:\\Vorlagen\\"/>');
    expect(xml).not.toContain('ns0:');
    expect(xml).not.toContain('<x16:revisionPtr');
  });

  it('keeps <workbookPr> ahead of the mc:AlternateContent block', async () => {
    const xml = await workbookXml(await roundTrip());
    expect(xml.indexOf('<workbookPr')).toBeGreaterThan(-1);
    expect(xml.indexOf('<workbookPr')).toBeLessThan(xml.indexOf('<mc:AlternateContent'));
    expect(xml.indexOf('<mc:AlternateContent')).toBeLessThan(xml.indexOf('<bookViews>'));
  });

  it('keeps the xr2:uid on <workbookView>', async () => {
    const xml = await workbookXml(await roundTrip());
    expect(xml).toContain('xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"');
  });

  it('is byte-stable across a second round-trip', async () => {
    const once = await roundTrip();
    const twice = await roundTrip(once);
    expect(await workbookXml(twice)).toEqual(await workbookXml(once));
    // XSD is skipped: the conformance harness strips mc-ignorable content
    // before validating, which empties this fixture's `<ext>` — the same
    // complaint it raises about the input file, so it says nothing about the
    // round-trip. OPC + semantic checks still run.
    expect((await validateXlsx(twice, { skipXsd: true })).issues).toEqual([]);
  });
});
