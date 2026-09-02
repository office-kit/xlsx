// Regression for https://github.com/office-kit/xlsx/issues/113 — everything
// after `<cellStyles>` in xl/styles.xml was dropped on save. `<dxfs>` was only
// written when it held entries, and `<tableStyles>`, `<colors>` and `<extLst>`
// (where Excel keeps its x14 slicer / timeline styles) were never read at all,
// so they disappeared from every workbook that went through load → save.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { openZip } from '../../src/zip/reader';
import { validateXlsx } from '../conformance/validate';

const enc = new TextEncoder();
const dec = new TextDecoder();

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

// Excel's own styles.xml tail: a dxf a table style points at, the table-style
// defaults, an MRU colour and the x14 extension that carries slicer/timeline
// styles.
const STYLES_TAIL =
  '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>' +
  '<tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">' +
  '<tableStyle name="Firmenstil" pivot="0" count="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle>' +
  '</tableStyles>' +
  '<colors><mruColors><color rgb="FF3366CC"/></mruColors></colors>' +
  `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="${X14_NS}">` +
  '<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst>';

const buildFixtureWithTail = (tail: string): Uint8Array =>
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
        '<sheets><sheet name="Tabelle1" sheetId="1" r:id="rId1"/></sheets></workbook>',
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
        '<cellStyles count="1"><cellStyle name="Standard" xfId="0" builtinId="0"/></cellStyles>' +
        tail +
        '</styleSheet>',
    ),
    'xl/worksheets/sheet1.xml': enc.encode(
      `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
        '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
    ),
  });

const buildFixture = (): Uint8Array => buildFixtureWithTail(STYLES_TAIL);

const roundTrip = async (bytes: Uint8Array = buildFixture()): Promise<Uint8Array> =>
  workbookToBytes(await loadWorkbook(fromBuffer(bytes)));

const stylesXml = async (bytes: Uint8Array): Promise<string> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    return dec.decode(archive.read('xl/styles.xml'));
  } finally {
    archive.close();
  }
};

describe('issue #113 — styles.xml keeps everything after <cellStyles>', () => {
  it('keeps tableStyles, colors and extLst across a round-trip', async () => {
    const xml = await stylesXml(await roundTrip());
    expect(xml).toContain('<tableStyle name="Firmenstil"');
    expect(xml).toContain('<tableStyleElement type="wholeTable" dxfId="0"/>');
    expect(xml).toContain('defaultTableStyle="TableStyleMedium2"');
    expect(xml).toContain('defaultPivotStyle="PivotStyleLight16"');
    expect(xml).toContain('<color rgb="FF3366CC"/>');
    expect(xml).toContain('{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}');
    expect(xml).toContain('slicerStyles defaultSlicerStyle="SlicerStyleLight1"');
  });

  it('keeps them in CT_Stylesheet order, after <dxfs>', async () => {
    const xml = await stylesXml(await roundTrip());
    const order = ['<dxfs', '<tableStyles', '<colors', 'extLst'].map((tag) => xml.indexOf(tag));
    expect(order.every((i) => i > -1)).toBe(true);
    expect([...order].sort((a, b) => a - b)).toEqual(order);
  });

  it('writes <dxfs count="0"/> even when the pool is empty', async () => {
    // Excel puts the element in every workbook it writes, and a rule's dxfId
    // is an index into it, so the empty anchor has to survive too.
    const emptyDxfs = buildFixtureWithTail('<dxfs count="0"/>');
    const xml = await stylesXml(await roundTrip(emptyDxfs));
    expect(xml).toContain('<dxfs count="0"/>');
  });

  it('is stable across a second round-trip and validates', async () => {
    const once = await roundTrip();
    const twice = await roundTrip(once);
    expect(await stylesXml(twice)).toEqual(await stylesXml(once));
    const result = await validateXlsx(twice);
    expect(result.issues).toEqual([]);
  });
});
