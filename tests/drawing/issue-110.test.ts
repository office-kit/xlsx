// Regression for https://github.com/office-kit/xlsx/issues/110 — an anchor
// holding anything the model doesn't cover was rewritten as a chart
// `graphicFrame` with an empty `r:id`, so a rectangle came back as "Chart 1"
// with a dangling chart reference and no drawing rels part. Anchors sitting
// inside an `<mc:AlternateContent>` wrapper were dropped outright, leaving an
// empty `<xdr:wsDr/>` — which is what still broke worksheets carrying form
// controls after #105: the worksheet side was fixed, but the shapes that draw
// the controls live in drawing1.xml.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load';
import { fromBuffer } from '../../src/io/node';
import { workbookToBytes } from '../../src/io/save';
import { openZip } from '../../src/zip/reader';

const enc = new TextEncoder();
const dec = new TextDecoder();

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main';

// 1×1 PNG-shaped bytes: signature + an IHDR big enough for the dimension read.
const png = (w: number): Uint8Array =>
  new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52,
    0, 0, 0, w, 0, 0, 0, 1, 8, 6, 0, 0, 0,
  ]);

const MARKER = (tag: string, col: number, row: number): string =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;

// The reporter's shape, verbatim.
const SHAPE_ANCHOR =
  `<xdr:twoCellAnchor>${MARKER('from', 1, 1)}${MARKER('to', 3, 3)}` +
  '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Rechteck 1"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
  '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>' +
  '<xdr:clientData fLocksWithSheet="0"/></xdr:twoCellAnchor>';

// What Excel writes when a shape needs a feature older readers can't render:
// the anchor itself sits inside the wrapper.
const ALTERNATE_CONTENT =
  `<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Choice xmlns:a14="${A14_NS}" Requires="a14">` +
  `<xdr:twoCellAnchor>${MARKER('from', 5, 5)}${MARKER('to', 7, 7)}` +
  '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Ellipse 2"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
  '<xdr:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>' +
  '<xdr:clientData/></xdr:twoCellAnchor></mc:Choice></mc:AlternateContent>';

// A group the model doesn't cover, holding a picture — so the raw node carries
// an r:embed that has to keep resolving.
const GROUP_ANCHOR =
  `<xdr:twoCellAnchor>${MARKER('from', 9, 9)}${MARKER('to', 11, 11)}` +
  '<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="4" name="Gruppe 3"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>' +
  '<xdr:grpSpPr/>' +
  '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="5" name="Bild im Verband"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
  `<xdr:blipFill><a:blip xmlns:r="${REL_NS}" r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
  '<xdr:spPr/></xdr:pic></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>';

// A picture the model does cover — its rId is reallocated on save, and must
// not land on the id the raw group still uses.
const PICTURE_ANCHOR =
  `<xdr:oneCellAnchor>${MARKER('from', 13, 13)}<xdr:ext cx="500000" cy="500000"/>` +
  '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="6" name="Logo"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
  `<xdr:blipFill><a:blip xmlns:r="${REL_NS}" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
  '<xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>';

const DRAWING_XML =
  `${XML_DECL}<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">` +
  SHAPE_ANCHOR +
  ALTERNATE_CONTENT +
  GROUP_ANCHOR +
  PICTURE_ANCHOR +
  '</xdr:wsDr>';

const buildFixture = (): Uint8Array =>
  zipSync({
    '[Content_Types].xml': enc.encode(
      `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Default Extension="png" ContentType="image/png"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
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
        '</styleSheet>',
    ),
    'xl/worksheets/sheet1.xml': enc.encode(
      `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData/><drawing r:id="rId1"/></worksheet>`,
    ),
    'xl/worksheets/_rels/sheet1.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/drawing" Target="../drawings/drawing1.xml"/>` +
        '</Relationships>',
    ),
    'xl/drawings/drawing1.xml': enc.encode(DRAWING_XML),
    'xl/drawings/_rels/drawing1.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="${REL_NS}/image" Target="../media/image1.png"/>` +
        `<Relationship Id="rId2" Type="${REL_NS}/image" Target="../media/image2.png"/>` +
        '</Relationships>',
    ),
    'xl/media/image1.png': png(1),
    'xl/media/image2.png': png(2),
  });

const roundTrip = async (bytes: Uint8Array = buildFixture()): Promise<Uint8Array> =>
  workbookToBytes(await loadWorkbook(fromBuffer(bytes)));

interface Package {
  part(path: string): string;
  has(path: string): boolean;
  paths: string[];
}

const openPackage = async (bytes: Uint8Array): Promise<Package> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    const paths = archive.list();
    const parts = new Map<string, string>();
    for (const p of paths) parts.set(p, dec.decode(archive.read(p)));
    return {
      paths,
      has: (p) => parts.has(p),
      part: (p) => parts.get(p) ?? '',
    };
  } finally {
    archive.close();
  }
};

describe('issue #110 — unmodeled drawing content survives a round-trip', () => {
  it('keeps the shape instead of rewriting it as an empty chart frame', async () => {
    const pkg = await openPackage(await roundTrip());
    const drawing = pkg.part('xl/drawings/drawing1.xml');
    expect(drawing).toContain('<xdr:cNvPr id="2" name="Rechteck 1"/>');
    expect(drawing).toContain('<a:prstGeom prst="rect">');
    expect(drawing).toContain('<a:ext cx="1000000" cy="500000"/>');
    // The anchor's own clientData flags come back too.
    expect(drawing).toContain('<xdr:clientData fLocksWithSheet="0"/>');
    expect(drawing).not.toContain('name="Chart 1"');
    expect(drawing).not.toContain('r:id=""');
  });

  it('keeps an mc:AlternateContent-wrapped anchor, in place', async () => {
    const drawing = (await openPackage(await roundTrip())).part('xl/drawings/drawing1.xml');
    expect(drawing).toContain('Requires="a14"');
    expect(drawing).toContain('<xdr:cNvPr id="3" name="Ellipse 2"/>');
    // Document order is z-order: the wrapper sat between the rectangle and the
    // group, and has to stay there.
    expect(drawing.indexOf('Rechteck 1')).toBeLessThan(drawing.indexOf('Ellipse 2'));
    expect(drawing.indexOf('Ellipse 2')).toBeLessThan(drawing.indexOf('Gruppe 3'));
  });

  it('keeps the rel a verbatim node references, and its image part', async () => {
    const pkg = await openPackage(await roundTrip());
    const drawing = pkg.part('xl/drawings/drawing1.xml');
    const rels = pkg.part('xl/drawings/_rels/drawing1.xml.rels');
    // The group still points at rId2, so rId2 has to still be an image rel...
    expect(drawing).toContain('r:embed="rId2"');
    const target = /<Relationship Id="rId2"[^>]*Target="([^"]+)"/.exec(rels)?.[1];
    expect(target).toBeDefined();
    // ...pointing at a part that exists.
    const resolved = `xl/${(target as string).replace('../', '')}`;
    expect(pkg.has(resolved)).toBe(true);
  });

  it('does not reuse a carried rId for the modeled picture', async () => {
    const pkg = await openPackage(await roundTrip());
    const drawing = pkg.part('xl/drawings/drawing1.xml');
    // The modeled picture is the last anchor; read the r:embed that follows
    // its cNvPr name.
    const picRId = /r:embed="(rId\d+)"/.exec(drawing.slice(drawing.indexOf('name="Logo"')))?.[1];
    expect(picRId).toBeDefined();
    expect(picRId).not.toBe('rId2');
  });

  it('is byte-stable across a second round-trip', async () => {
    const once = await roundTrip();
    const twice = await roundTrip(once);
    expect((await openPackage(twice)).part('xl/drawings/drawing1.xml')).toEqual(
      (await openPackage(once)).part('xl/drawings/drawing1.xml'),
    );
  });
});
