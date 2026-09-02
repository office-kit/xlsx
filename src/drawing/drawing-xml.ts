// `xl/drawings/drawingN.xml` reader/writer. Charts and pictures are modeled;
// shapes, connectors and groups — and Excel's `<mc:AlternateContent>` wrappers
// around an anchor — are kept as the verbatim source XML and written back
// untouched, because the model has no slot for their geometry or text.

import { escapeXmlAttr } from '../utils/escape';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { DRAWING_NS, REL_NS, SHEET_DRAWING_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { serializeXml } from '../xml/serializer';
import { findChild, type XmlNode } from '../xml/tree';
import type { AnchorMarker, DrawingAnchor, Point2D, PositiveSize2D } from './anchor';
import { parseShapeProperties, serializeShapeProperties } from './dml/dml-xml';
import { type ChartReference, type Drawing, type DrawingItem, makeDrawing, type PictureReference } from './drawing';

const WS_DRAWING_TAG = `{${SHEET_DRAWING_NS}}wsDr`;
const ABSOLUTE_ANCHOR_TAG = `{${SHEET_DRAWING_NS}}absoluteAnchor`;
const ONE_CELL_ANCHOR_TAG = `{${SHEET_DRAWING_NS}}oneCellAnchor`;
const TWO_CELL_ANCHOR_TAG = `{${SHEET_DRAWING_NS}}twoCellAnchor`;
const FROM_TAG = `{${SHEET_DRAWING_NS}}from`;
const TO_TAG = `{${SHEET_DRAWING_NS}}to`;
const POS_TAG = `{${SHEET_DRAWING_NS}}pos`;
const EXT_TAG = `{${SHEET_DRAWING_NS}}ext`;
const COL_TAG = `{${SHEET_DRAWING_NS}}col`;
const COLOFF_TAG = `{${SHEET_DRAWING_NS}}colOff`;
const ROW_TAG = `{${SHEET_DRAWING_NS}}row`;
const ROWOFF_TAG = `{${SHEET_DRAWING_NS}}rowOff`;
const GRAPHIC_FRAME_TAG = `{${SHEET_DRAWING_NS}}graphicFrame`;
const CLIENT_DATA_TAG = `{${SHEET_DRAWING_NS}}clientData`;

const A_GRAPHIC_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/main}graphic';
const A_GRAPHIC_DATA_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData';
const C_CHART_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/chart}chart';
const CX_CHART_TAG = '{http://schemas.microsoft.com/office/drawing/2014/chartex}chart';
const PIC_TAG = `{${SHEET_DRAWING_NS}}pic`;
const PIC_NV_PR_TAG = `{${SHEET_DRAWING_NS}}nvPicPr`;
const C_NV_PR_TAG = `{${SHEET_DRAWING_NS}}cNvPr`;
const BLIP_FILL_TAG = `{${SHEET_DRAWING_NS}}blipFill`;
const PIC_SP_PR_TAG = `{${SHEET_DRAWING_NS}}spPr`;
const A_BLIP_TAG = `{${DRAWING_NS}}blip`;

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const escapeAttr = escapeXmlAttr;

const parseIntChild = (parent: XmlNode, tag: string): number | undefined => {
  const child = findChild(parent, tag);
  if (!child) return undefined;
  const n = Number.parseInt(child.text ?? '', 10);
  return Number.isInteger(n) ? n : undefined;
};

const parseAnchorMarker = (node: XmlNode): AnchorMarker | undefined => {
  const col = parseIntChild(node, COL_TAG);
  const colOff = parseIntChild(node, COLOFF_TAG);
  const row = parseIntChild(node, ROW_TAG);
  const rowOff = parseIntChild(node, ROWOFF_TAG);
  if (col === undefined || row === undefined) return undefined;
  return { col, colOff: colOff ?? 0, row, rowOff: rowOff ?? 0 };
};

const parsePoint2D = (node: XmlNode): Point2D | undefined => {
  const x = node.attrs['x'];
  const y = node.attrs['y'];
  if (!x || !y) return undefined;
  const xn = Number.parseInt(x, 10);
  const yn = Number.parseInt(y, 10);
  if (!Number.isFinite(xn) || !Number.isFinite(yn)) return undefined;
  return { x: xn, y: yn };
};

const parsePositiveSize2D = (node: XmlNode): PositiveSize2D | undefined => {
  const cx = node.attrs['cx'];
  const cy = node.attrs['cy'];
  if (!cx || !cy) return undefined;
  const cxn = Number.parseInt(cx, 10);
  const cyn = Number.parseInt(cy, 10);
  if (!Number.isFinite(cxn) || !Number.isFinite(cyn)) return undefined;
  return { cx: cxn, cy: cyn };
};

const parsePictureReference = (node: XmlNode): PictureReference | undefined => {
  const pic = findChild(node, PIC_TAG);
  if (!pic) return undefined;
  const out: PictureReference = {};
  const nvPicPr = findChild(pic, PIC_NV_PR_TAG);
  if (nvPicPr) {
    const cNvPr = findChild(nvPicPr, C_NV_PR_TAG);
    if (cNvPr) {
      if (cNvPr.attrs['name'] !== undefined) out.name = cNvPr.attrs['name'];
      if (cNvPr.attrs['descr'] !== undefined) out.descr = cNvPr.attrs['descr'];
      const hiddenRaw = cNvPr.attrs['hidden'];
      if (hiddenRaw === '1' || hiddenRaw === 'true') out.hidden = true;
    }
  }
  const blipFill = findChild(pic, BLIP_FILL_TAG);
  if (blipFill) {
    const blip = findChild(blipFill, A_BLIP_TAG);
    const embed = blip?.attrs[`{${REL_NS}}embed`];
    if (embed) out.rId = embed;
  }
  const spPrEl = findChild(pic, PIC_SP_PR_TAG);
  if (spPrEl) out.spPr = parseShapeProperties(spPrEl);
  return out;
};

const parseChartReference = (node: XmlNode): ChartReference | undefined => {
  const graphic = findChild(node, GRAPHIC_FRAME_TAG);
  if (!graphic) return undefined;
  const aGraphic = findChild(graphic, A_GRAPHIC_TAG);
  const aGraphicData = aGraphic ? findChild(aGraphic, A_GRAPHIC_DATA_TAG) : undefined;
  if (!aGraphicData) return undefined;
  // Chart references use either `<c:chart>` (legacy) or `<cx:chart>`
  // (chartex). Both carry the rels rId on the element.
  const cChart = findChild(aGraphicData, C_CHART_TAG);
  if (cChart) {
    const rId = cChart.attrs[`{${REL_NS}}id`];
    return rId !== undefined ? { rId } : {};
  }
  const cxChart = findChild(aGraphicData, CX_CHART_TAG);
  if (cxChart) {
    const rId = cxChart.attrs[`{${REL_NS}}id`];
    return rId !== undefined ? { rId, isCx: true } : { isCx: true };
  }
  return undefined;
};

const parseAnchor = (node: XmlNode): DrawingItem | undefined => {
  let anchor: DrawingAnchor | undefined;
  if (node.name === ABSOLUTE_ANCHOR_TAG) {
    const pos = findChild(node, POS_TAG);
    const ext = findChild(node, EXT_TAG);
    if (!pos || !ext) return undefined;
    const p = parsePoint2D(pos);
    const e = parsePositiveSize2D(ext);
    if (!p || !e) return undefined;
    anchor = { kind: 'absolute', pos: p, ext: e };
  } else if (node.name === ONE_CELL_ANCHOR_TAG) {
    const fromEl = findChild(node, FROM_TAG);
    const ext = findChild(node, EXT_TAG);
    const from = fromEl ? parseAnchorMarker(fromEl) : undefined;
    const e = ext ? parsePositiveSize2D(ext) : undefined;
    if (!from || !e) return undefined;
    anchor = { kind: 'oneCell', from, ext: e };
  } else if (node.name === TWO_CELL_ANCHOR_TAG) {
    const fromEl = findChild(node, FROM_TAG);
    const toEl = findChild(node, TO_TAG);
    const from = fromEl ? parseAnchorMarker(fromEl) : undefined;
    const to = toEl ? parseAnchorMarker(toEl) : undefined;
    if (!from || !to) return undefined;
    const editAsRaw = node.attrs['editAs'];
    const validEditAs = editAsRaw === 'twoCell' || editAsRaw === 'oneCell' || editAsRaw === 'absolute';
    anchor = {
      kind: 'twoCell',
      from,
      to,
      ...(validEditAs ? { editAs: editAsRaw as 'twoCell' | 'oneCell' | 'absolute' } : {}),
    };
  } else {
    return undefined;
  }

  // Detect content kind in priority order: chart graphic frame first,
  // then picture; everything else is tagged "unsupported".
  const chart = parseChartReference(node);
  if (chart) {
    return { anchor, content: { kind: 'chart', chart } };
  }
  const picture = parsePictureReference(node);
  if (picture) {
    return { anchor, content: { kind: 'picture', picture } };
  }
  // Find the first child that isn't a marker/pos/ext/clientData. Anything we
  // don't model keeps the whole anchor element verbatim — rebuilding it would
  // drop the shape body along with the anchor's own `editAs` and the
  // `<xdr:clientData fLocksWithSheet="0">` flags form controls rely on.
  const skip = new Set([FROM_TAG, TO_TAG, POS_TAG, EXT_TAG, CLIENT_DATA_TAG]);
  for (const child of node.children) {
    if (skip.has(child.name)) continue;
    return { anchor, content: { kind: 'unsupported', rawTag: child.name }, raw: node };
  }
  // Bare anchor with no content — record as unsupported with empty tag.
  return { anchor, content: { kind: 'unsupported', rawTag: '' }, raw: node };
};

/** Parse a `xl/drawings/drawingN.xml` payload into a Drawing object. */
export function parseDrawingXml(bytes: Uint8Array | string): Drawing {
  const root = parseXml(bytes);
  if (root.name !== WS_DRAWING_TAG) {
    throw new OpenXmlSchemaError(`parseDrawingXml: root is "${root.name}", expected wsDr`);
  }
  const items: DrawingItem[] = [];
  const rawChildren: Array<{ beforeItem: number; node: XmlNode }> = [];
  // Document order matters — anchors must round-trip in their source
  // sequence regardless of their kind, and that order is the z-order.
  for (const child of root.children) {
    if (child.name === ABSOLUTE_ANCHOR_TAG || child.name === ONE_CELL_ANCHOR_TAG || child.name === TWO_CELL_ANCHOR_TAG) {
      const item = parseAnchor(child);
      if (item) items.push(item);
      continue;
    }
    // Not an anchor: Excel's `<mc:AlternateContent>` wrapper around one, or
    // anything else a future schema puts here. Keep it verbatim at its place
    // in the sequence.
    rawChildren.push({ beforeItem: items.length, node: child });
  }
  const drawing = makeDrawing(items);
  if (rawChildren.length > 0) drawing.rawChildren = rawChildren;
  return drawing;
}

// Declared on the `<xdr:wsDr>` root the fragments below are spliced into, so
// captured nodes reuse these prefixes instead of redeclaring them per anchor.
const WS_DR_PREFIXES: Readonly<Record<string, string>> = {
  [SHEET_DRAWING_NS]: 'xdr',
  [DRAWING_NS]: 'a',
};

/** Re-emit a captured node inline, verbatim. */
const serializeRawNode = (node: XmlNode): string =>
  new TextDecoder().decode(serializeXml(node, { xmlDeclaration: false, inScopePrefixes: WS_DR_PREFIXES }));

/**
 * Every relationship id referenced inside the drawing's verbatim nodes —
 * `r:embed` on a picture nested in a group, `r:id` on a shape's hyperlink, and
 * whatever else sits under an `<mc:AlternateContent>` wrapper. The loader uses
 * this to carry the matching rels (and their target parts) across a re-save.
 */
export function collectRawRelIds(drawing: Drawing): Set<string> {
  const ids = new Set<string>();
  const walk = (node: XmlNode): void => {
    for (const [name, value] of Object.entries(node.attrs)) {
      if (name.startsWith(`{${REL_NS}}`) && value.length > 0) ids.add(value);
    }
    for (const child of node.children) walk(child);
  };
  for (const item of drawing.items) {
    if (item.raw) walk(item.raw);
  }
  for (const rc of drawing.rawChildren ?? []) walk(rc.node);
  return ids;
}

const serializeMarker = (tag: string, m: AnchorMarker): string =>
  `<xdr:${tag}><xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff}</xdr:colOff><xdr:row>${m.row}</xdr:row><xdr:rowOff>${m.rowOff}</xdr:rowOff></xdr:${tag}>`;

const serializePictureFrame = (picture: PictureReference, anchorIdx: number): string => {
  const rId = picture.rId ?? '';
  const idAttr = `id="${anchorIdx + 2}"`;
  const nameAttr = `name="${escapeAttr(picture.name ?? `Picture ${anchorIdx + 1}`)}"`;
  const descrAttr = picture.descr !== undefined ? ` descr="${escapeAttr(picture.descr)}"` : '';
  const hiddenAttr = picture.hidden ? ' hidden="1"' : '';
  const spPr = picture.spPr
    ? serializeShapeProperties(picture.spPr, 'xdr:spPr')
    : '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>';
  return [
    '<xdr:pic>',
    '<xdr:nvPicPr>',
    `<xdr:cNvPr ${idAttr} ${nameAttr}${descrAttr}${hiddenAttr}/>`,
    '<xdr:cNvPicPr/>',
    '</xdr:nvPicPr>',
    '<xdr:blipFill>',
    `<a:blip xmlns:r="${REL_NS}" r:embed="${escapeAttr(rId)}"/>`,
    '<a:stretch><a:fillRect/></a:stretch>',
    '</xdr:blipFill>',
    spPr,
    '</xdr:pic>',
  ].join('');
};

const serializeChartGraphicFrame = (
  chart: ChartReference,
  anchorIdx: number,
  ext: { cx: number; cy: number },
): string => {
  const rId = chart.rId ?? '';
  const isCx = chart.isCx === true;
  // chartex (`cx:`) and the legacy ECMA-376 (`c:`) chart kinds use
  // different `<a:graphicData uri>` values *and* different inline
  // chart-reference element namespaces. Mismatched URIs make Excel
  // reject the workbook on open.
  const uri = isCx
    ? 'http://schemas.microsoft.com/office/drawing/2014/chartex'
    : 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const chartEl = isCx
    ? `<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="${REL_NS}" r:id="${escapeAttr(rId)}"/>`
    : `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${REL_NS}" r:id="${escapeAttr(rId)}"/>`;
  return [
    '<xdr:graphicFrame>',
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="${anchorIdx + 2}" name="Chart ${anchorIdx + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`,
    // The graphicFrame's xfrm carries the chart's actual on-sheet size.
    // Some Excel versions render a zero-size frame as a hidden / 0-pixel
    // element, which is why earlier outputs showed a chart-less workbook.
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${ext.cx}" cy="${ext.cy}"/></xdr:xfrm>`,
    `<a:graphic><a:graphicData uri="${uri}">`,
    chartEl,
    '</a:graphicData></a:graphic>',
    '</xdr:graphicFrame>',
  ].join('');
};

const serializeAnchor = (item: DrawingItem, idx: number): string => {
  if (item.raw) return serializeRawNode(item.raw);
  const a = item.anchor;
  let body = '';
  // Default size for twoCellAnchor where the from→to delta is unknown at this
  // layer: Excel auto-resizes graphicFrames to the cell range so any non-zero
  // value is fine. 6"×4" matches openpyxl's default chart frame.
  let chartExt: { cx: number; cy: number } = { cx: 5486400, cy: 3200400 };
  if (a.kind === 'absolute') {
    body = `<xdr:pos x="${a.pos.x}" y="${a.pos.y}"/><xdr:ext cx="${a.ext.cx}" cy="${a.ext.cy}"/>`;
    chartExt = a.ext;
  } else if (a.kind === 'oneCell') {
    body = `${serializeMarker('from', a.from)}<xdr:ext cx="${a.ext.cx}" cy="${a.ext.cy}"/>`;
    chartExt = a.ext;
  } else {
    body = `${serializeMarker('from', a.from)}${serializeMarker('to', a.to)}`;
  }
  let content = '';
  if (item.content.kind === 'chart') {
    content = serializeChartGraphicFrame(item.content.chart, idx, chartExt);
  } else if (item.content.kind === 'picture') {
    content = serializePictureFrame(item.content.picture, idx);
  } else {
    // Unsupported content with no captured source — only reachable for an item
    // built in code, since the reader always records `raw`. An empty chart
    // frame keeps the anchor well-formed.
    content = serializeChartGraphicFrame({}, idx, chartExt);
  }
  const editAs = a.kind === 'twoCell' && a.editAs ? ` editAs="${a.editAs}"` : '';
  const tag = a.kind === 'absolute' ? 'absoluteAnchor' : a.kind === 'oneCell' ? 'oneCellAnchor' : 'twoCellAnchor';
  return `<xdr:${tag}${editAs}>${body}${content}<xdr:clientData/></xdr:${tag}>`;
};

/** Serialise a Drawing to its `xl/drawings/drawingN.xml` payload. */
export function drawingToBytes(drawing: Drawing): Uint8Array {
  return new TextEncoder().encode(serializeDrawing(drawing));
}

function serializeDrawing(drawing: Drawing): string {
  const parts: string[] = [
    XML_HEADER,
    `<xdr:wsDr xmlns:xdr="${SHEET_DRAWING_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`,
  ];
  const rawChildren = drawing.rawChildren ?? [];
  const emitRawBefore = (index: number): void => {
    for (const rc of rawChildren) {
      if (rc.beforeItem === index) parts.push(serializeRawNode(rc.node));
    }
  };
  for (let i = 0; i < drawing.items.length; i++) {
    emitRawBefore(i);
    const item = drawing.items[i];
    if (item) parts.push(serializeAnchor(item, i));
  }
  emitRawBefore(drawing.items.length);
  parts.push('</xdr:wsDr>');
  return parts.join('');
}
