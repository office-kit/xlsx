// Insert ▸ Shapes / Text Box and the Shape Format tab: building shapes the
// way Excel writes them, and the edits the tab applies. Every edit is one
// undo step on the sheet's `drawing`.

import {
  makeColor,
  makeLine,
  makeNoFill,
  makePresetGeometry,
  makeSchemeColor,
  makeShapeDrawingItem,
  makeSolidFill,
  makeSrgbColor,
  makeDrawing,
  type ShapeReference,
  type ShapeStyle,
  type TextBody,
  type TextParagraph,
} from '@office-kit/xlsx/drawing';
import type { Axis } from './axis.ts';
import { anchorFor, type Rect } from './drawing-anchor.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';

/** What Insert ▸ Shapes / Text Box puts the editor into: the next drag on the grid draws this. */
export type ShapeTool = { readonly kind: 'shape'; readonly prst: string } | { readonly kind: 'textBox' };

/** The Shapes gallery, grouped as in Excel's menu. */
export const SHAPE_GALLERY: ReadonlyArray<{ readonly group: 'shpLines' | 'shpRectangles' | 'shpBasic' | 'shpArrows' | 'shpStars' | 'shpCallouts'; readonly shapes: readonly string[] }> = [
  { group: 'shpLines', shapes: ['line', 'straightConnector1'] },
  { group: 'shpRectangles', shapes: ['rect', 'roundRect'] },
  { group: 'shpBasic', shapes: ['ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid', 'pentagon', 'hexagon', 'octagon', 'plus', 'flowChartTerminator'] },
  { group: 'shpArrows', shapes: ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'chevron', 'homePlate'] },
  { group: 'shpStars', shapes: ['star4', 'star5', 'star6', 'star8'] },
  { group: 'shpCallouts', shapes: ['wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout'] },
];

/** Lines are `<xdr:cxnSp>` connectors; `straightConnector1` is the arrow, with a triangle tail. */
const CONNECTOR_PRESETS: ReadonlySet<string> = new Set(['line', 'straightConnector1']);

/** A click without a drag inserts Excel's default 1-inch shape. */
export const DEFAULT_SHAPE_PX = 96;

const accent1 = () => makeColor(makeSchemeColor('accent1'));

/** Excel's theme style for a new shape: accent1 fill, darker accent1 outline, light text. */
function shapeStyle(connector: boolean): ShapeStyle {
  return connector
    ? {
        lnRef: { idx: 1, color: accent1() },
        fillRef: { idx: 0, color: accent1() },
        effectRef: { idx: 0, color: accent1() },
        fontRef: { idx: 'minor', color: makeColor(makeSchemeColor('tx1')) },
      }
    : {
        lnRef: { idx: 2, color: makeColor(makeSchemeColor('accent1'), [{ kind: 'shade', val: 50000 }]) },
        fillRef: { idx: 1, color: accent1() },
        effectRef: { idx: 0, color: accent1() },
        fontRef: { idx: 'minor', color: makeColor(makeSchemeColor('lt1')) },
      };
}

function newShape(tool: ShapeTool, flipH: boolean, flipV: boolean): ShapeReference {
  if (tool.kind === 'textBox') {
    // Excel's text box: white fill, thin grey outline, top-anchored wrapping text.
    return {
      textBox: true,
      spPr: {
        geometry: makePresetGeometry('rect'),
        fill: makeSolidFill(makeColor(makeSchemeColor('lt1'))),
        ln: makeLine({ w: 9525, fill: makeSolidFill(makeColor(makeSchemeColor('tx1'), [{ kind: 'lumMod', val: 25000 }, { kind: 'lumOff', val: 75000 }])) }),
      },
      style: { ...shapeStyle(false), fontRef: { idx: 'minor', color: makeColor(makeSchemeColor('dk1')) } },
      txBody: { bodyPr: { vertOverflow: 'clip', horzOverflow: 'clip', wrap: 'square', rtlCol: false, anchor: 't' }, paragraphs: [{ runs: [] }] },
    };
  }
  const connector = CONNECTOR_PRESETS.has(tool.prst);
  const flips = flipH || flipV ? { xfrm: { ...(flipH ? { flipH: true } : {}), ...(flipV ? { flipV: true } : {}) } } : {};
  if (connector) {
    return {
      connector: true,
      spPr: {
        ...flips,
        geometry: makePresetGeometry(tool.prst),
        ...(tool.prst === 'straightConnector1' ? { ln: makeLine({ tailEnd: { type: 'triangle' } }) } : {}),
      },
      style: shapeStyle(true),
    };
  }
  return {
    spPr: { geometry: makePresetGeometry(tool.prst) },
    style: shapeStyle(false),
    txBody: { bodyPr: { rtlCol: false, anchor: 'ctr' }, paragraphs: [{ pPr: { algn: 'ctr' }, runs: [] }] },
  };
}

/**
 * Draw a shape from `start` to `end`, in content px measured on `axes`. Pass
 * the on-screen geometry axes rather than the sheet's: in Page Layout view
 * they include the paper margins and gaps, so the sheet's axes would misplace
 * the shape. A drag that runs up or to the left still yields a positive box; a
 * line remembers the direction as a flip. A click (no drag) inserts the
 * default size. Returns the new item's index, or undefined when the sheet's
 * protection refused the edit.
 */
export function insertShape(doc: SpreadsheetEditor, axes: { readonly cols: Axis; readonly rows: Axis }, tool: ShapeTool, start: { x: number; y: number }, end: { x: number; y: number }): number | undefined {
  const dragged = Math.abs(end.x - start.x) >= 3 || Math.abs(end.y - start.y) >= 3;
  const isLine = tool.kind === 'shape' && CONNECTOR_PRESETS.has(tool.prst);
  const rect: Rect = dragged
    ? { x: Math.min(start.x, end.x), y: Math.min(start.y, end.y), w: Math.abs(end.x - start.x), h: Math.abs(end.y - start.y) }
    : { x: start.x, y: start.y, w: DEFAULT_SHAPE_PX, h: isLine ? 0 : DEFAULT_SHAPE_PX };
  const shape = newShape(tool, dragged && end.x < start.x, dragged && end.y < start.y);
  const ws = doc.ws;
  const seed = { kind: 'twoCell' as const, from: { col: 0, colOff: 0, row: 0, rowOff: 0 }, to: { col: 0, colOff: 0, row: 0, rowOff: 0 } };
  const anchor = anchorFor(seed, rect, axes.cols, axes.rows);
  return doc.transact(tool.kind === 'textBox' ? 'Insert Text Box' : 'Insert Shape', (tx) => {
    tx.sheet(ws, 'drawing');
    ws.drawing ??= makeDrawing([]);
    ws.drawing.items.push(makeShapeDrawingItem(anchor, shape));
    return ws.drawing.items.length - 1;
  });
}

/** The shape at `index` on the active sheet, if that item is one. */
export function shapeAt(doc: SpreadsheetEditor, index: number | null): ShapeReference | undefined {
  if (index === null) return undefined;
  const item = doc.ws.drawing?.items[index];
  return item?.content.kind === 'shape' ? item.content.shape : undefined;
}

/** Apply `fn` to the shape at `index` as one undo step. */
export function editShape(doc: SpreadsheetEditor, index: number, label: string, fn: (shape: ShapeReference) => void): void {
  const ws = doc.ws;
  if (!shapeAt(doc, index)) return;
  doc.transact(label, (tx) => {
    tx.sheet(ws, 'drawing');
    const item = ws.drawing?.items[index];
    if (item?.content.kind === 'shape') fn(item.content.shape);
  });
}

/** The shape's text as typed: paragraphs joined by newlines. */
export function shapeText(body: TextBody | undefined): string {
  if (!body) return '';
  return body.paragraphs
    .map((p) => p.runs.map((r) => (r.kind === 'br' ? '\n' : r.kind === 'r' ? r.t : (r.t ?? ''))).join(''))
    .join('\n');
}

/**
 * Replace the text, one paragraph per line. Each new paragraph reuses the
 * first paragraph's alignment and the first run's font, which is what typing
 * over a shape's text keeps in Excel.
 */
export function withText(body: TextBody | undefined, text: string): TextBody {
  const base: TextBody = body ?? { bodyPr: { rtlCol: false, anchor: 'ctr' }, paragraphs: [] };
  const first = base.paragraphs[0];
  const firstRun = first?.runs.find((r) => r.kind === 'r');
  const rPr = firstRun?.rPr ?? first?.endParaRPr;
  const paragraphs: TextParagraph[] = text.split('\n').map((line) => {
    const p: TextParagraph = { runs: line === '' ? [] : [rPr ? { kind: 'r', rPr, t: line } : { kind: 'r', t: line }] };
    if (first?.pPr) p.pPr = first.pPr;
    if (rPr) p.endParaRPr = rPr;
    return p;
  });
  return { ...base, paragraphs };
}

export function setShapeText(doc: SpreadsheetEditor, index: number, text: string): void {
  const shape = shapeAt(doc, index);
  if (!shape || shape.connector || shapeText(shape.txBody) === text) return;
  editShape(doc, index, 'Edit Text', (s) => {
    s.txBody = withText(s.txBody, text);
  });
}

/** Shape Fill: a `#RRGGBB` colour, or null for No Fill. */
export function setShapeFill(doc: SpreadsheetEditor, index: number, color: string | null): void {
  editShape(doc, index, 'Shape Fill', (s) => {
    s.spPr.fill = color === null ? makeNoFill() : makeSolidFill(makeColor(makeSrgbColor(color.replace('#', ''))));
  });
}

/** Shape Outline: a `#RRGGBB` colour, or null for No Outline. */
export function setShapeOutline(doc: SpreadsheetEditor, index: number, color: string | null): void {
  editShape(doc, index, 'Shape Outline', (s) => {
    s.spPr.ln = { ...s.spPr.ln, fill: color === null ? makeNoFill() : makeSolidFill(makeColor(makeSrgbColor(color.replace('#', '')))) };
  });
}

/** Shape Outline ▸ Weight, in points. */
export function setShapeOutlineWeight(doc: SpreadsheetEditor, index: number, pt: number): void {
  editShape(doc, index, 'Shape Outline', (s) => {
    s.spPr.ln = { ...s.spPr.ln, w: Math.round(pt * 12700) };
  });
}

/**
 * Bring Forward / Send Backward: swap with the neighbour in z-order (document
 * order). Returns the item's new index.
 */
export function reorderDrawing(doc: SpreadsheetEditor, index: number, step: 1 | -1): number {
  const ws = doc.ws;
  const items = ws.drawing?.items;
  const target = index + step;
  if (!items || target < 0 || target >= items.length) return index;
  const moved = doc.transact(step > 0 ? 'Bring Forward' : 'Send Backward', (tx) => {
    tx.sheet(ws, 'drawing');
    const list = ws.drawing?.items;
    const a = list?.[index];
    const b = list?.[target];
    if (!list || !a || !b) return false;
    list[index] = b;
    list[target] = a;
    return true;
  });
  return moved ? target : index;
}
