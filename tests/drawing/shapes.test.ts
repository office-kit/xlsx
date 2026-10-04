import { describe, expect, it } from 'vitest';
import { makeOneCellAnchor, makeTwoCellAnchor } from '../../src/drawing/anchor.js';
import { makeSchemeColor, makeSrgbColor, makeColor } from '../../src/drawing/dml/colors.js';
import { makeSolidFill } from '../../src/drawing/dml/fill.js';
import { makePresetGeometry } from '../../src/drawing/dml/geometry.js';
import { makeLine } from '../../src/drawing/dml/line.js';
import { makeParagraph, makeRun, makeTextBody } from '../../src/drawing/dml/text.js';
import { drawingToBytes, parseDrawingXml } from '../../src/drawing/drawing-xml.js';
import { type DrawingItem, makeDrawing, makeShapeDrawingItem, type ShapeStyle } from '../../src/drawing/drawing.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import type { Worksheet } from '../../src/worksheet/worksheet.js';
import type { Chartsheet } from '../../src/chartsheet/chartsheet.js';
import { validateXlsx } from '../conformance/validate.js';

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const EXCEL_STYLE: ShapeStyle = {
  lnRef: { idx: 2, color: makeColor(makeSchemeColor('accent1'), [{ kind: 'shade', val: 50000 }]) },
  fillRef: { idx: 1, color: makeColor(makeSchemeColor('accent1')) },
  effectRef: { idx: 0, color: makeColor(makeSchemeColor('accent1')) },
  fontRef: { idx: 'minor', color: makeColor(makeSchemeColor('lt1')) },
};

const wsDr = (anchors: string): string =>
  `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">${anchors}</xdr:wsDr>`;

const MARKERS =
  '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>';

const anchorWith = (body: string, clientData = '<xdr:clientData/>'): string =>
  `<xdr:twoCellAnchor>${MARKERS}${body}${clientData}</xdr:twoCellAnchor>`;

// What Excel for Mac writes for Insert ▸ Shapes ▸ Rounded Rectangle with typed text.
const EXCEL_SHAPE =
  '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Rounded Rectangle 1"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
  '<xdr:spPr><a:xfrm><a:off x="825500" y="203200"/><a:ext cx="2032000" cy="1016000"/></a:xfrm>' +
  '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></xdr:spPr>' +
  '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>' +
  '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
  '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>' +
  '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/>' +
  '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1100" b="1"/><a:t>Hello</a:t></a:r></a:p></xdr:txBody></xdr:sp>';

const onlyItem = (xml: string): DrawingItem => {
  const item = parseDrawingXml(wsDr(xml)).items[0];
  if (!item) throw new Error('expected one drawing item');
  return item;
};

const expectSheet = (ws: Worksheet | Chartsheet | undefined): Worksheet => {
  if (!ws || !('rows' in ws)) throw new Error('expected worksheet');
  return ws;
};

describe('drawing shapes — reader', () => {
  it('models an Excel-authored shape: geometry, theme style and text', () => {
    const item = onlyItem(anchorWith(EXCEL_SHAPE));
    expect(item.raw).toBeUndefined();
    expect(item.content.kind).toBe('shape');
    if (item.content.kind !== 'shape') return;
    const { shape } = item.content;
    expect(shape.name).toBe('Rounded Rectangle 1');
    expect(shape.spPr.geometry).toEqual({ kind: 'preset', prst: 'roundRect' });
    expect(shape.style).toEqual(EXCEL_STYLE);
    expect(shape.txBody?.bodyPr.anchor).toBe('ctr');
    expect(shape.txBody?.paragraphs[0]?.pPr?.algn).toBe('ctr');
    expect(shape.txBody?.paragraphs[0]?.runs[0]).toEqual({ kind: 'r', rPr: { lang: 'en-US', sz: 1100, b: true }, t: 'Hello' });
  });

  it('models a text box and a connector', () => {
    const textBox = onlyItem(
      anchorWith(
        '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="TextBox 2"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>' +
          '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></xdr:spPr>' +
          '<xdr:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/><a:p><a:r><a:t>Note</a:t></a:r></a:p></xdr:txBody></xdr:sp>',
      ),
    );
    expect(textBox.content.kind === 'shape' && textBox.content.shape.textBox).toBe(true);

    const line = onlyItem(
      anchorWith(
        '<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="4" name="Straight Connector 3"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>' +
          '<xdr:spPr><a:xfrm flipV="1"/><a:prstGeom prst="line"><a:avLst/></a:prstGeom></xdr:spPr></xdr:cxnSp>',
      ),
    );
    expect(line.content.kind).toBe('shape');
    if (line.content.kind !== 'shape') return;
    expect(line.content.shape.connector).toBe(true);
    expect(line.content.shape.spPr.xfrm?.flipV).toBe(true);
  });

  it.each([
    ['a macro', EXCEL_SHAPE.replace('macro=""', 'macro="[0]!Run"')],
    [
      'a hyperlink through the drawing rels',
      EXCEL_SHAPE.replace(
        '<xdr:cNvPr id="2" name="Rounded Rectangle 1"/>',
        `<xdr:cNvPr id="2" name="Rounded Rectangle 1"><a:hlinkClick xmlns:r="${REL_NS}" r:id="rId1"/></xdr:cNvPr>`,
      ),
    ],
    ['shape locks', EXCEL_SHAPE.replace('<xdr:cNvSpPr/>', '<xdr:cNvSpPr><a:spLocks noGrp="1"/></xdr:cNvSpPr>')],
  ])('keeps a shape with %s verbatim', (_label, sp) => {
    const item = onlyItem(anchorWith(sp));
    expect(item.content.kind).toBe('unsupported');
    expect(item.raw).toBeDefined();
  });

  it('keeps a shape verbatim when its anchor carries clientData flags', () => {
    const item = onlyItem(anchorWith(EXCEL_SHAPE, '<xdr:clientData fLocksWithSheet="0"/>'));
    expect(item.content.kind).toBe('unsupported');
  });
});

describe('drawing shapes — writer', () => {
  it('writes a shape back with the same model', () => {
    const drawing = parseDrawingXml(wsDr(anchorWith(EXCEL_SHAPE)));
    const again = parseDrawingXml(drawingToBytes(drawing));
    expect(again.items).toEqual(drawing.items);
  });

  it('round-trips shapes, a text box and a line through a saved workbook that passes the conformance validator', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Shapes');
    const items: DrawingItem[] = [
      makeShapeDrawingItem(makeTwoCellAnchor({ from: 'B2', to: 'E6' }), {
        name: 'Oval 1',
        spPr: { geometry: makePresetGeometry('ellipse') },
        style: EXCEL_STYLE,
        txBody: makeTextBody([makeParagraph([makeRun('Centre', { sz: 1400, b: true })], { algn: 'ctr' })], { anchor: 'ctr' }),
      }),
      makeShapeDrawingItem(makeTwoCellAnchor({ from: 'G2', to: 'J4' }), {
        textBox: true,
        spPr: {
          geometry: makePresetGeometry('rect'),
          fill: makeSolidFill(makeColor(makeSrgbColor('FFF2CC'))),
          ln: makeLine({ w: 12700, fill: makeSolidFill(makeColor(makeSrgbColor('BF9000'))) }),
        },
        txBody: makeTextBody([makeParagraph([makeRun('Text box')])], { wrap: 'square', anchor: 't' }),
      }),
      makeShapeDrawingItem(makeOneCellAnchor({ from: 'B9', widthPx: 200, heightPx: 0 }), {
        connector: true,
        spPr: { geometry: makePresetGeometry('straightConnector1'), ln: makeLine({ w: 19050, tailEnd: { type: 'triangle' } }) },
        style: EXCEL_STYLE,
      }),
      makeShapeDrawingItem(makeTwoCellAnchor({ from: 'G7', to: 'I12' }), {
        spPr: { geometry: makePresetGeometry('star5'), fill: makeSolidFill(makeColor(makeSchemeColor('accent2'))) },
      }),
    ];
    ws.drawing = makeDrawing(items);
    const bytes = await workbookToBytes(wb);
    expect(await validateXlsx(bytes)).toMatchObject({ status: 'valid' });

    const ws2 = expectSheet((await loadWorkbook(fromBuffer(bytes))).sheets[0]?.sheet);
    expect(ws2.drawing?.items.map((item) => item.content)).toEqual(
      items.map((item, i) =>
        // The writer names an unnamed shape after its kind and position.
        item.content.kind === 'shape' && item.content.shape.name === undefined
          ? { kind: 'shape', shape: { ...item.content.shape, name: `${item.content.shape.textBox ? 'TextBox' : item.content.shape.connector ? 'Straight Connector' : 'Shape'} ${i + 1}` } }
          : item.content,
      ),
    );
  });
});
