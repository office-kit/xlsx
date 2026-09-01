// Round-tripping a form control (issue #105): `<legacyDrawing>` + Excel's
// x14-gated `<controls>` block. The saved package used to drop the VML
// relationship while re-emitting the element, and to stitch the `<controls>`
// wrapper in before `<drawing>` with its `Requires="x14"` prefix undeclared —
// Excel refused the file. The header/footer-picture half of the report is
// covered by issue-104.test.ts.
//
// The packages below are in-memory reductions of Excel 365 output — the
// same shape as the issue's fixtures, minus everything unrelated.

import { describe, expect, it } from 'vitest';
import { unzipSync } from 'fflate';
import { fromBuffer, toBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { manifestFromBytes } from '../../src/packaging/manifest';
import { relsFromBytes } from '../../src/packaging/relationships';
import { REL_NS } from '../../src/xml/namespaces';
import { createZipWriter } from '../../src/zip/writer';
import type { Worksheet } from '../../src/worksheet/worksheet';
import { validateXlsx } from '../conformance/validate';

const td = new TextDecoder();
const te = new TextEncoder();

const VML_DRAWING_REL = `${REL_NS}/vmlDrawing`;
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

const CONTENT_TYPES_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
const ROOT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
const WORKSHEET_OVERRIDE = (n: number): string =>
  `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
const WORKSHEET_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">';
const RELS_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
const VML_REL = (id: string, n: number): string =>
  `<Relationship Id="${id}" Type="${VML_DRAWING_REL}" Target="../drawings/vmlDrawing${n}.vml"/>`;

const workbookParts = (sheetNames: ReadonlyArray<string>): Record<string, string> => ({
  '_rels/.rels': ROOT_RELS,
  'xl/workbook.xml':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
    sheetNames.map((name, i) => `<sheet name="${name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
    '</sheets></workbook>',
  'xl/_rels/workbook.xml.rels':
    RELS_HEAD +
    sheetNames
      .map(
        (_, i) =>
          `<Relationship Id="rId${i + 1}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
      )
      .join('') +
    '</Relationships>',
});

// Excel's VML for a single checkbox (`x:ClientData ObjectType="Checkbox"`) —
// no `ObjectType="Note"`, so the loader must treat it as a passthrough part.
const CHECKBOX_VML = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
 <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>
 <v:shape id="_x0000_s1027" type="#_x0000_t201" style='position:absolute;margin-left:208.5pt;margin-top:392.25pt;width:30.75pt;height:15pt;z-index:1;mso-wrap-style:tight' filled="f" fillcolor="window [65]" stroked="f" strokecolor="windowText [64]" o:insetmode="auto">
  <v:path shadowok="t" strokeok="t" fillok="t"/><o:lock v:ext="edit" rotation="t"/>
  <v:textbox style='mso-direction-alt:auto' o:singleclick="f"><div style='text-align:left'></div></v:textbox>
  <x:ClientData ObjectType="Checkbox"><x:SizeWithCells/><x:Anchor>2, 6, 22, 4, 2, 47, 22, 24</x:Anchor><x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine><x:TextVAlign>Center</x:TextVAlign><x:NoThreeD/></x:ClientData>
 </v:shape>
</xml>`;

// Header/footer picture VML: `&G` in the header resolves to this shape, whose
// image lives behind the VML part's own rels file.
const HEADER_PICTURE_VML = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
 <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>
 <v:shape id="CH" o:spid="_x0000_s2049" type="#_x0000_t75" style='position:absolute;margin-left:0;margin-top:0;width:60pt;height:20pt;z-index:1'><v:imagedata o:relid="rId1" o:title="logo"/><o:lock v:ext="edit" rotation="t"/></v:shape>
</xml>`;

// A comment overlay VML — the marker Excel puts on note shapes is what makes
// the loader regenerate this part from `Worksheet.legacyComments`.
const COMMENT_VML = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
 <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
 <v:shape id="_x0000_s1025" type="#_x0000_t202" style='position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto">
  <v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>
  <x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1, 15, 0, 2, 3, 15, 3, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData>
 </v:shape>
</xml>`;

const COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Author</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>hello</t></r></text></comment></commentList></comments>';

const PNG_BYTES = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x00]);

async function packageFrom(parts: Record<string, string | Uint8Array>): Promise<Uint8Array> {
  const sink = toBuffer();
  const w = createZipWriter(sink);
  for (const [path, body] of Object.entries(parts)) {
    await w.addEntry(path, typeof body === 'string' ? te.encode(body) : body);
  }
  await w.finalize();
  return new Uint8Array(sink.result());
}

const roundTrip = async (input: Uint8Array): Promise<Uint8Array> =>
  workbookToBytes(await loadWorkbook(fromBuffer(input)));

const expectWorksheet = (wb: Awaited<ReturnType<typeof loadWorkbook>>, index: number): Worksheet => {
  const sheet = wb.sheets[index]?.sheet;
  if (!sheet || !('rows' in sheet)) throw new Error(`expected worksheet at index ${index}`);
  return sheet;
};

/** rId on `<tag r:id="…"/>` in a worksheet part; `undefined` when absent. */
const rIdOf = (sheetXml: string, tag: string): string | undefined =>
  new RegExp(`<${tag} r:id="([^"]+)"/>`).exec(sheetXml)?.[1];

/** Content type the manifest resolves for `path`, via Override or Default. */
const contentTypeOf = (zip: Record<string, Uint8Array>, path: string): string | undefined => {
  const manifest = manifestFromBytes(zip['[Content_Types].xml'] ?? new Uint8Array());
  const override = manifest.overrides.find((o) => o.partName === `/${path}`);
  if (override) return override.contentType;
  const ext = path.slice(path.lastIndexOf('.') + 1);
  return manifest.defaults.find((d) => d.ext === ext)?.contentType;
};

const expectConformant = async (bytes: Uint8Array, options: { skipXsd?: boolean } = {}): Promise<void> => {
  const result = await validateXlsx(bytes, options);
  expect(result.issues, result.issues.map((i) => `${i.tier} ${i.part}: ${i.message}`).join('\n')).toEqual([]);
};

describe('issue #105 — form control (<legacyDrawing> + <controls>)', () => {
  const build = (): Promise<Uint8Array> =>
    packageFrom({
      '[Content_Types].xml':
        CONTENT_TYPES_HEAD +
        WORKSHEET_OVERRIDE(1) +
        '<Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/></Types>',
      ...workbookParts(['Sheet1']),
      'xl/worksheets/sheet1.xml':
        WORKSHEET_HEAD +
        '<dimension ref="A1"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait"/>' +
        '<legacyDrawing r:id="rId1"/>' +
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><controls><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><control shapeId="1027" r:id="rId2" name="Check Box 3"><controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0"><anchor moveWithCells="1"><from><xdr:col>2</xdr:col><xdr:colOff>57150</xdr:colOff><xdr:row>22</xdr:row><xdr:rowOff>38100</xdr:rowOff></from><to><xdr:col>2</xdr:col><xdr:colOff>447675</xdr:colOff><xdr:row>22</xdr:row><xdr:rowOff>228600</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent></controls></mc:Choice></mc:AlternateContent>' +
        '</worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels':
        RELS_HEAD +
        VML_REL('rId1', 1) +
        `<Relationship Id="rId2" Type="${REL_NS}/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/></Relationships>`,
      'xl/drawings/vmlDrawing1.vml': CHECKBOX_VML,
      'xl/ctrlProps/ctrlProp1.xml':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" lockText="1" noThreeD="1"/>',
    });

  it('keeps <legacyDrawing> together with its vmlDrawing relationship and part', async () => {
    const out = await roundTrip(await build());
    const zip = unzipSync(out);
    const sheetXml = td.decode(zip['xl/worksheets/sheet1.xml']);
    const rels = relsFromBytes(zip['xl/worksheets/_rels/sheet1.xml.rels'] ?? new Uint8Array());

    const rId = rIdOf(sheetXml, 'legacyDrawing');
    expect(rId).toBeDefined();
    const rel = rels.rels.find((r) => r.id === rId);
    expect(rel?.type).toBe(VML_DRAWING_REL);
    expect(rel?.target).toBe('../drawings/vmlDrawing1.vml');
    expect(td.decode(zip['xl/drawings/vmlDrawing1.vml'])).toBe(CHECKBOX_VML);
    expect(contentTypeOf(zip, 'xl/drawings/vmlDrawing1.vml')).toBe(
      'application/vnd.openxmlformats-officedocument.vmlDrawing',
    );
  });

  it('emits <controls> after <legacyDrawing>, gated by a declared x14 prefix', async () => {
    const out = await roundTrip(await build());
    const zip = unzipSync(out);
    const sheetXml = td.decode(zip['xl/worksheets/sheet1.xml']);

    expect(sheetXml).toContain('<controls>');
    expect(sheetXml.indexOf('<legacyDrawing ')).toBeLessThan(sheetXml.indexOf('<controls>'));
    expect(sheetXml).toContain('Requires="x14"');
    expect(sheetXml).toContain(`xmlns:x14="${X14_NS}"`);
    // The block is re-serialised from the typed model, not stitched in as an
    // opaque node with auto-allocated prefixes.
    expect(sheetXml).not.toContain('ns0:');
    // ctrlProp rel survives alongside the re-emitted control.
    expect(sheetXml).toContain('<control shapeId="1027" r:id="rId2" name="Check Box 3">');
    expect(sheetXml).toContain('<anchor moveWithCells="1">');
    // Excel writes `<controlPr><anchor><from>` in the main namespace although
    // CT_ObjectAnchor's XSD wants `xdr:from`; the input fails the same XSD
    // check, so only the OPC + semantic tiers are meaningful here.
    await expectConformant(out, { skipXsd: true });
  });

  it('reloads the saved package with the control and its VML link intact', async () => {
    const wb = await loadWorkbook(fromBuffer(await roundTrip(await build())));
    const ws = expectWorksheet(wb, 0);
    expect(ws.legacyDrawingRId).toBeDefined();
    expect(ws.controls).toHaveLength(1);
    expect(ws.controls[0]?.name).toBe('Check Box 3');
    expect(ws.controls[0]?.controlPr).toBeDefined();
  });
});

describe('issue #105 — comment VML next to passthrough VML', () => {
  // Sheet1: header picture at vmlDrawing1 (the name the comments writer would
  // pick first). Sheet2: a comment overlay at vmlDrawing2 that is regenerated.
  const build = (): Promise<Uint8Array> =>
    packageFrom({
      '[Content_Types].xml':
        CONTENT_TYPES_HEAD +
        '<Default Extension="png" ContentType="image/png"/>' +
        WORKSHEET_OVERRIDE(1) +
        WORKSHEET_OVERRIDE(2) +
        '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>',
      ...workbookParts(['Sheet1', 'Sheet2']),
      'xl/worksheets/sheet1.xml':
        WORKSHEET_HEAD +
        '<dimension ref="A1"/><sheetData/><headerFooter><oddHeader>&amp;L&amp;G</oddHeader></headerFooter><legacyDrawingHF r:id="rId1"/></worksheet>',
      'xl/worksheets/_rels/sheet1.xml.rels': RELS_HEAD + VML_REL('rId1', 1) + '</Relationships>',
      'xl/worksheets/sheet2.xml':
        WORKSHEET_HEAD +
        '<dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><legacyDrawing r:id="rId1"/></worksheet>',
      'xl/worksheets/_rels/sheet2.xml.rels':
        RELS_HEAD +
        VML_REL('rId1', 2) +
        `<Relationship Id="rId2" Type="${REL_NS}/comments" Target="../comments1.xml"/></Relationships>`,
      'xl/drawings/vmlDrawing1.vml': HEADER_PICTURE_VML,
      'xl/drawings/_rels/vmlDrawing1.vml.rels':
        RELS_HEAD + `<Relationship Id="rId1" Type="${REL_NS}/image" Target="../media/image1.png"/></Relationships>`,
      'xl/media/image1.png': PNG_BYTES,
      'xl/drawings/vmlDrawing2.vml': COMMENT_VML,
      'xl/comments1.xml': COMMENTS_XML,
    });

  it('regenerates the comment VML without clobbering the passthrough VML', async () => {
    const out = await roundTrip(await build());
    const zip = unzipSync(out);

    const sheet1 = td.decode(zip['xl/worksheets/sheet1.xml']);
    const rels1 = relsFromBytes(zip['xl/worksheets/_rels/sheet1.xml.rels'] ?? new Uint8Array());
    const hfRel = rels1.rels.find((r) => r.id === rIdOf(sheet1, 'legacyDrawingHF'));
    expect(hfRel?.type).toBe(VML_DRAWING_REL);
    expect(td.decode(zip[`xl/drawings/${hfRel?.target.replace('../drawings/', '')}`])).toBe(
      HEADER_PICTURE_VML,
    );

    const sheet2 = td.decode(zip['xl/worksheets/sheet2.xml']);
    const rels2 = relsFromBytes(zip['xl/worksheets/_rels/sheet2.xml.rels'] ?? new Uint8Array());
    expect(rels2.rels.filter((r) => r.type === VML_DRAWING_REL)).toHaveLength(1);
    const commentRel = rels2.rels.find((r) => r.id === rIdOf(sheet2, 'legacyDrawing'));
    expect(commentRel?.type).toBe(VML_DRAWING_REL);
    const commentVml = td.decode(zip[`xl/drawings/${commentRel?.target.replace('../drawings/', '')}`]);
    expect(commentVml).toContain('ObjectType="Note"');
    expect(commentVml).not.toBe(COMMENT_VML);

    expect(Object.keys(zip).filter((p) => p.endsWith('.vml'))).toHaveLength(2);
    await expectConformant(out);
  });
});
