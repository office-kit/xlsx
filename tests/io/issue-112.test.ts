// Regression for https://github.com/office-kit/xlsx/issues/112 — CT_Workbook
// (ECMA-376 §18.2.27) is an xsd:sequence, so `<workbook>`'s children have a
// normative order. The writer used to emit `<definedNames>` straight after
// `<sheets>` — ahead of `<functionGroups>` and `<externalReferences>` — and
// `<pivotCaches>` ahead of `<calcPr>`, producing a package Excel refuses.

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

// The reporter's fixture reduced to what the ordering needs: a workbook that
// carries both `<externalReferences>` and `<definedNames>`, plus a `<calcPr>`
// and a `<pivotCaches>` so the second swap shows too. The pivot cache part is
// carried through as a passthrough — only the workbook-root link matters here.
const buildFixture = (): Uint8Array =>
  zipSync({
    '[Content_Types].xml': enc.encode(
      `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>' +
        '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>' +
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
        '<externalReferences><externalReference r:id="rId2"/></externalReferences>' +
        '<definedNames><definedName name="Marke">Tabelle1!$A$1</definedName></definedNames>' +
        '<calcPr calcId="191029"/>' +
        '<pivotCaches><pivotCache cacheId="1" r:id="rId3"/></pivotCaches>' +
        '</workbook>',
    ),
    'xl/_rels/workbook.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL_NS}/externalLink" Target="externalLinks/externalLink1.xml"/>` +
        `<Relationship Id="rId3" Type="${REL_NS}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>` +
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
      `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
        '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
    ),
    'xl/externalLinks/externalLink1.xml': enc.encode(
      `${XML_DECL}<externalLink xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
        '<externalBook r:id="rId1"><sheetNames><sheetName val="Extern"/></sheetNames></externalBook>' +
        '</externalLink>',
    ),
    'xl/externalLinks/_rels/externalLink1.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/externalLinkPath" Target="Preise.xlsx" TargetMode="External"/>` +
        '</Relationships>',
    ),
    'xl/pivotCache/pivotCacheDefinition1.xml': enc.encode(
      `${XML_DECL}<pivotCacheDefinition xmlns="${MAIN_NS}" xmlns:r="${REL_NS}" recordCount="0">` +
        '<cacheSource type="worksheet"><worksheetSource ref="A1:A1" sheet="Tabelle1"/></cacheSource>' +
        '<cacheFields count="0"/></pivotCacheDefinition>',
    ),
  });

const roundTrip = async (): Promise<Uint8Array> =>
  workbookToBytes(await loadWorkbook(fromBuffer(buildFixture())));

const workbookXml = async (bytes: Uint8Array): Promise<string> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    return dec.decode(archive.read('xl/workbook.xml'));
  } finally {
    archive.close();
  }
};

describe('issue #112 — workbook.xml keeps the CT_Workbook element sequence', () => {
  it('writes externalReferences before definedNames', async () => {
    const xml = await workbookXml(await roundTrip());
    expect(xml.indexOf('<externalReferences>')).toBeGreaterThan(-1);
    expect(xml.indexOf('<externalReferences>')).toBeLessThan(xml.indexOf('<definedNames>'));
  });

  it('writes calcPr before pivotCaches', async () => {
    const xml = await workbookXml(await roundTrip());
    expect(xml.indexOf('<calcPr')).toBeGreaterThan(-1);
    expect(xml.indexOf('<calcPr')).toBeLessThan(xml.indexOf('<pivotCaches>'));
  });

  it('validates against the CT_Workbook schema', async () => {
    const result = await validateXlsx(await roundTrip());
    expect(result.issues).toEqual([]);
  });
});
