// Spreadsheet drawing data model.
//
// A `Drawing` is the per-worksheet `xl/drawings/drawingN.xml` part — a list of
// anchor entries, each carrying a content variant (chart, picture, shape,
// connector, group). Charts and pictures are modeled; everything else is kept
// as the verbatim source XML so it survives a round-trip untouched.

import type { ChartSpace } from '../chart/chart';
import type { CxChartSpace } from '../chart/cx/chartex';
import type { DrawingAnchor } from './anchor';
import type { ShapeProperties } from './dml/shape-properties';
import type { XlsxImage } from './image';

/** Reference to a chart part — the chart's drawing-rels rId resolves to xl/charts/chartN.xml. */
export interface ChartReference {
  /** Drawing-rels rId pointing at the chart part. Populated on read; the writer assigns its own. */
  rId?: string;
  /**
   * Legacy ECMA-376 chart payload (`c:chartSpace`). Stage-1 supports BarChart
   * end-to-end; other chart kinds populate this field as their parsers /
   * writers land.
   */
  space?: ChartSpace;
  /**
   * Excel-2016 chartex payload (`cx:chartSpace`). Mutually exclusive with
   * {@link space} for any given drawing item; the parser sniffs the root
   * element and populates whichever is appropriate.
   */
  cxSpace?: CxChartSpace;
  /**
   * `true` when the resolved chart part is a chartex (`cx:`) chart. Set by the
   * package writer so the drawing emitter knows to use the chartex
   * `<a:graphicData uri>` instead of the legacy chart URI — Excel rejects the
   * workbook when the URI doesn't match the chart's actual root namespace.
   */
  isCx?: boolean;
}

/** Reference to an embedded picture inside a worksheet drawing. */
export interface PictureReference {
  /** Drawing-rels rId pointing at the embedded image. Populated on read; the writer assigns its own. */
  rId?: string;
  /** Resolved image bytes + metadata. Populated on read; the writer reads it back. */
  image?: XlsxImage;
  /** Picture display name (`<xdr:cNvPr name="...">`). */
  name?: string;
  /** Optional alt-text description. */
  descr?: string;
  /** Hidden flag (`<xdr:cNvPr hidden="1"/>`). */
  hidden?: boolean;
  /** Per-picture shape properties (extra fill / line / rotation). */
  spPr?: ShapeProperties;
}

export interface DrawingItem {
  anchor: DrawingAnchor;
  content:
    | { kind: 'chart'; chart: ChartReference }
    | { kind: 'picture'; picture: PictureReference }
    | { kind: 'unsupported'; rawTag: string };
  /**
   * The whole source anchor element, captured verbatim. Set by the reader for
   * `unsupported` content — a plain shape, a group, a connector — because the
   * model has no slot for its geometry, text or `<xdr:clientData>` attributes.
   * When present the writer emits it as-is instead of rebuilding the anchor
   * from {@link anchor} + {@link content}.
   */
  raw?: import('../xml/tree').XmlNode;
}

export interface Drawing {
  items: DrawingItem[];
  /**
   * `<xdr:wsDr>` children that aren't anchors — in practice Excel's
   * `<mc:AlternateContent>` wrapper around an anchor that needs a newer
   * feature. Kept verbatim together with the `items` index they sat before, so
   * document order (which is z-order) survives a round-trip.
   */
  rawChildren?: Array<{ beforeItem: number; node: import('../xml/tree').XmlNode }>;
  /**
   * Drawing-rels entries the verbatim nodes above reference (`r:embed` on a
   * picture inside a group, `r:id` on a shape's hyperlink, …). Carried over
   * from the source part under their original ids so those references still
   * resolve after a re-save; the writer allocates its own ids around them.
   */
  rawRels?: import('../packaging/relationships').Relationships;
}

export function makeDrawing(items: DrawingItem[] = []): Drawing {
  return { items };
}

export function makeChartDrawingItem(anchor: DrawingAnchor, chart: ChartReference = {}): DrawingItem {
  return { anchor, content: { kind: 'chart', chart } };
}

export function makePictureDrawingItem(anchor: DrawingAnchor, picture: PictureReference | XlsxImage): DrawingItem {
  // Distinguish raw image bytes from a full PictureReference by checking the
  // discriminator: XlsxImage carries `format`, PictureReference doesn't.
  const ref: PictureReference = 'format' in picture ? { image: picture as XlsxImage } : (picture as PictureReference);
  return { anchor, content: { kind: 'picture', picture: ref } };
}

// ---- Worksheet ergonomic helpers ----------------------------------------

import { loadImage } from './image';
import { makeOneCellAnchor } from './anchor';
import type { Worksheet } from '../worksheet/worksheet';

/**
 * Drop an image onto a worksheet at a single-cell anchor. Lazy-allocates
 * `ws.drawing` (as `makeDrawing([])`) on first call and appends a picture
 * DrawingItem.
 *
 * `image` accepts either an `XlsxImage` (already loaded via `loadImage`) or raw
 * image bytes — in the bytes case, this helper sniffs the format with
 * `loadImage` itself.
 *
 * `at` is a cell ref like `"C3"`. Override `widthPx` / `heightPx` to scale;
 * otherwise the helper uses 96×96 defaults that look fine for typical icons.
 */
export const addImageAt = (
  ws: Worksheet,
  at: string,
  image: XlsxImage | Uint8Array,
  opts: { widthPx?: number; heightPx?: number } = {},
): DrawingItem => {
  const xlsxImage: XlsxImage = image instanceof Uint8Array ? loadImage(image) : image;
  const anchor = makeOneCellAnchor({
    from: at,
    widthPx: opts.widthPx ?? 96,
    heightPx: opts.heightPx ?? 96,
  });
  const item = makePictureDrawingItem(anchor, xlsxImage);
  if (!ws.drawing) ws.drawing = makeDrawing([]);
  ws.drawing.items.push(item);
  return item;
};

/**
 * Anchor a chart to a worksheet at a single-cell ref. Lazy-allocates
 * `ws.drawing`. `chart` is the same `ChartReference` shape `makeChart
 * DrawingItem` accepts (`{ space }` for legacy chart, `{ cxSpace }` for
 * chartex).
 */
export const addChartAt = (
  ws: Worksheet,
  at: string,
  chart: ChartReference,
  opts: { widthPx?: number; heightPx?: number } = {},
): DrawingItem => {
  const anchor = makeOneCellAnchor({
    from: at,
    widthPx: opts.widthPx ?? 480,
    heightPx: opts.heightPx ?? 320,
  });
  const item = makeChartDrawingItem(anchor, chart);
  if (!ws.drawing) ws.drawing = makeDrawing([]);
  ws.drawing.items.push(item);
  return item;
};

/**
 * Read-only snapshot of every picture DrawingItem on the sheet. Returns the
 * matching items (each with its anchor + picture reference). Empty array when
 * the sheet has no drawing or only non-picture items.
 */
export const listImagesOnSheet = (ws: Worksheet): ReadonlyArray<DrawingItem> => {
  if (!ws.drawing) return [];
  return ws.drawing.items.filter((it) => it.content.kind === 'picture');
};

/**
 * Read-only snapshot of every chart DrawingItem on the sheet. Each item has its
 * anchor + chart reference.
 */
export const listChartsOnSheet = (ws: Worksheet): ReadonlyArray<DrawingItem> => {
  if (!ws.drawing) return [];
  return ws.drawing.items.filter((it) => it.content.kind === 'chart');
};

/**
 * Drop every DrawingItem from the worksheet. Returns the count removed. The
 * `ws.drawing` field itself is left in place (empty) so subsequent `addImageAt`
 * / `addChartAt` calls don't have to re-allocate.
 */
export const removeAllDrawingItems = (ws: Worksheet): number => {
  if (!ws.drawing) return 0;
  const n = ws.drawing.items.length;
  ws.drawing.items = [];
  return n;
};

/**
 * Drop every picture DrawingItem from the worksheet, leaving charts and any
 * other content kinds untouched. Returns the count removed.
 */
export const removeAllImages = (ws: Worksheet): number => {
  if (!ws.drawing) return 0;
  const before = ws.drawing.items.length;
  ws.drawing.items = ws.drawing.items.filter((it) => it.content.kind !== 'picture');
  return before - ws.drawing.items.length;
};

/**
 * Drop every chart DrawingItem from the worksheet, leaving pictures and any
 * other content kinds untouched. Returns the count removed.
 */
export const removeAllCharts = (ws: Worksheet): number => {
  if (!ws.drawing) return 0;
  const before = ws.drawing.items.length;
  ws.drawing.items = ws.drawing.items.filter((it) => it.content.kind !== 'chart');
  return before - ws.drawing.items.length;
};
