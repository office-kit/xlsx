// Regression for https://github.com/office-kit/xlsx/issues/104 — a sheet with a
// header/footer picture (`<legacyDrawingHF r:id>` → VML → image) must survive
// load → save with every relationship still resolvable. Before the fix the
// writer re-emitted the r:id but dropped the sheet rel, the VML's own rels and
// the image, and Excel refused the output as corrupt.

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

// Minimal package modelled on the reporter's fixture: one empty sheet whose
// header is `&G` (picture placeholder), backed by a VML drawing that carries
// no comment shapes and references one JPEG. The VML is deliberately numbered
// 2 so a comments writer starting at vmlDrawing1 would not mask a dropped part.
const IMAGE_BYTES = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00]);

const buildFixture = (): Uint8Array =>
  zipSync({
    '[Content_Types].xml': enc.encode(
      `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
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
      `${XML_DECL}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL_NS}">` +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>` +
        '</Relationships>',
    ),
    'xl/styles.xml': enc.encode(
      `${XML_DECL}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
        '</styleSheet>',
    ),
    'xl/worksheets/sheet1.xml': enc.encode(
      `${XML_DECL}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL_NS}">` +
        '<sheetData/>' +
        '<headerFooter><oddHeader>&amp;L&amp;G</oddHeader></headerFooter>' +
        '<legacyDrawingHF r:id="rId4"/>' +
        '</worksheet>',
    ),
    'xl/worksheets/_rels/sheet1.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId4" Type="${REL_NS}/vmlDrawing" Target="../drawings/vmlDrawing2.vml"/>` +
        '</Relationships>',
    ),
    'xl/drawings/vmlDrawing2.vml': enc.encode(
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
        '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
        '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">' +
        '<v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
        '<v:shape id="LH" o:spid="_x0000_s1025" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:100pt;height:40pt;z-index:1">' +
        '<v:imagedata o:relid="rId1" o:title="letterhead"/><o:lock v:ext="edit" rotation="t"/></v:shape>' +
        '</xml>',
    ),
    'xl/drawings/_rels/vmlDrawing2.vml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/image" Target="../media/image1.jpeg"/>` +
        '</Relationships>',
    ),
    'xl/media/image1.jpeg': IMAGE_BYTES,
  });

const roundTrip = async (): Promise<Uint8Array> =>
  workbookToBytes(await loadWorkbook(fromBuffer(buildFixture())));

describe('issue #104 — header/footer VML picture survives a round-trip', () => {
  it('re-emits <legacyDrawingHF> together with a resolvable vmlDrawing rel', async () => {
    const bytes = await roundTrip();
    const archive = await openZip(fromBuffer(bytes));
    try {
      const sheetXml = dec.decode(archive.read('xl/worksheets/sheet1.xml'));
      const hfRId = /<legacyDrawingHF r:id="([^"]+)"\/>/.exec(sheetXml)?.[1];
      expect(hfRId).toBeDefined();

      expect(archive.has('xl/worksheets/_rels/sheet1.xml.rels')).toBe(true);
      const sheetRels = dec.decode(archive.read('xl/worksheets/_rels/sheet1.xml.rels'));
      expect(sheetRels).toMatch(
        new RegExp(
          `<Relationship[^>]*Id="${hfRId}"[^>]*Type="${REL_NS}/vmlDrawing"[^>]*Target="../drawings/vmlDrawing2.vml"`,
        ),
      );
    } finally {
      archive.close();
    }
  });

  it('carries the VML part, its rels and the referenced image', async () => {
    const bytes = await roundTrip();
    const archive = await openZip(fromBuffer(bytes));
    try {
      expect(archive.has('xl/drawings/vmlDrawing2.vml')).toBe(true);
      expect(archive.has('xl/drawings/_rels/vmlDrawing2.vml.rels')).toBe(true);
      expect(archive.read('xl/media/image1.jpeg')).toEqual(IMAGE_BYTES);
    } finally {
      archive.close();
    }
  });

  it('produces a package where every part has a content type and every rel resolves', async () => {
    const result = await validateXlsx(await roundTrip(), { skipXsd: true });
    expect(result.issues).toEqual([]);
  });

  it('is stable across a second round-trip', async () => {
    const once = await roundTrip();
    const twice = await workbookToBytes(await loadWorkbook(fromBuffer(once)));
    const result = await validateXlsx(twice, { skipXsd: true });
    expect(result.issues).toEqual([]);
  });
});
