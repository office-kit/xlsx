// xl/charts/chartN.xml read/write.
//
// Stage-1 covers BarChart end-to-end: parse + serialize with title / legend /
// catAx / valAx / series (cat + val refs + numCache / strCache). Other chart
// kinds slot in alongside as their own `<c:lineChart>` / `<c:pieChart>` / etc.
// parsers.

import {
  parseShapeProperties,
  parseTextBody,
  serializeShapeProperties,
  serializeTextBody,
} from '../drawing/dml/dml-xml.js';
import type { ShapeProperties } from '../drawing/dml/shape-properties.js';
import type { TextBody } from '../drawing/dml/text.js';
import { escapeXmlText } from '../utils/escape.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { CHART_NS, REL_NS, SHEET_DRAWING_NS } from '../xml/namespaces.js';
import { parseXml } from '../xml/parser.js';
import { findChild, findChildren, type XmlNode } from '../xml/tree.js';
import {
  type Area3DChart,
  type AreaChart,
  type Bar3DChart,
  type BarChart,
  type BarDirection,
  type BarSeries,
  type BubbleChart,
  type BubbleSeries,
  type BubbleSizeRepresents,
  type AxisCrossBetween,
  type AxisCrosses,
  type AxisOrientation,
  type AxisScaling,
  type CategoryAxis,
  type CategoryLabelAlignment,
  type DateAxis,
  type SeriesAxis,
  type TimeUnit,
  type CategoryRef,
  type ChartKind,
  type ChartSpace,
  type ChartTitle,
  type DataLabel,
  type DataLabelList,
  type DataLabelPosition,
  type DoughnutChart,
  type ErrorBarDirection,
  type ErrorBars,
  type ErrorBarType,
  type ErrorValType,
  type Gridlines,
  type GroupingType,
  type Legend,
  type LegendPosition,
  type Line3DChart,
  type LineChart,
  type LineSeries,
  type Marker,
  type MarkerSymbol,
  type DataPoint,
  type HiLowLines,
  type Layout,
  type LayoutMode,
  type LayoutTarget,
  type ManualLayout,
  type UpDownBars,
  makeArea3DChart,
  makeAreaChart,
  makeBar3DChart,
  makeBarChart,
  makeBarSeries,
  makeBubbleChart,
  makeBubbleSeries,
  makeChartSpace,
  makeDoughnutChart,
  makeLine3DChart,
  makeLineChart,
  makeOfPieChart,
  makePie3DChart,
  makePieChart,
  makeRadarChart,
  makeScatterChart,
  makeScatterSeries,
  makeStockChart,
  makeSurface3DChart,
  makeSurfaceChart,
  type ChartNumberFormat,
  type NumericRef,
  type OfPieChart,
  type OfPieType,
  type Pie3DChart,
  type PieChart,
  type PlotArea,
  type RadarChart,
  type RadarStyle,
  type ScatterChart,
  type ScatterSeries,
  type ScatterStyle,
  type SplitType,
  type StockChart,
  type Surface3DChart,
  type SurfaceChart,
  type SurfaceFrame,
  type TickLabelPosition,
  type TickMark,
  type Trendline,
  type TrendlineType,
  type ValueAxis,
  type View3D,
} from './chart.js';

const CHART_SPACE_TAG = `{${CHART_NS}}chartSpace`;
const CHART_TAG = `{${CHART_NS}}chart`;
const TITLE_TAG = `{${CHART_NS}}title`;
const TX_TAG = `{${CHART_NS}}tx`;
const RICH_TAG = `{${CHART_NS}}rich`;
const PLOT_AREA_TAG = `{${CHART_NS}}plotArea`;
const BAR_CHART_TAG = `{${CHART_NS}}barChart`;
const CAT_AX_TAG = `{${CHART_NS}}catAx`;
const VAL_AX_TAG = `{${CHART_NS}}valAx`;
const DATE_AX_TAG = `{${CHART_NS}}dateAx`;
const SER_AX_TAG = `{${CHART_NS}}serAx`;
const SER_TAG = `{${CHART_NS}}ser`;
const IDX_TAG = `{${CHART_NS}}idx`;
const ORDER_TAG = `{${CHART_NS}}order`;
const CAT_TAG = `{${CHART_NS}}cat`;
const VAL_TAG = `{${CHART_NS}}val`;
const NUM_REF_TAG = `{${CHART_NS}}numRef`;
const STR_REF_TAG = `{${CHART_NS}}strRef`;
const NUM_CACHE_TAG = `{${CHART_NS}}numCache`;
const STR_CACHE_TAG = `{${CHART_NS}}strCache`;
const F_TAG = `{${CHART_NS}}f`;
const PT_TAG = `{${CHART_NS}}pt`;
const V_TAG = `{${CHART_NS}}v`;
const FORMAT_CODE_TAG = `{${CHART_NS}}formatCode`;
const BAR_DIR_TAG = `{${CHART_NS}}barDir`;
const GROUPING_TAG = `{${CHART_NS}}grouping`;
const VARY_COLORS_TAG = `{${CHART_NS}}varyColors`;
const GAP_WIDTH_TAG = `{${CHART_NS}}gapWidth`;
const OVERLAP_TAG = `{${CHART_NS}}overlap`;
const STYLE_TAG = `{${CHART_NS}}style`;
const AX_ID_TAG = `{${CHART_NS}}axId`;
const DELETE_TAG = `{${CHART_NS}}delete`;
const LINE_CHART_TAG = `{${CHART_NS}}lineChart`;
const AREA_CHART_TAG = `{${CHART_NS}}areaChart`;
const PIE_CHART_TAG = `{${CHART_NS}}pieChart`;
const DOUGHNUT_CHART_TAG = `{${CHART_NS}}doughnutChart`;
const SCATTER_CHART_TAG = `{${CHART_NS}}scatterChart`;
const RADAR_CHART_TAG = `{${CHART_NS}}radarChart`;
const SMOOTH_TAG = `{${CHART_NS}}smooth`;
const MARKER_TAG = `{${CHART_NS}}marker`;
const SYMBOL_TAG = `{${CHART_NS}}symbol`;
const SIZE_TAG = `{${CHART_NS}}size`;
const D_PT_TAG = `{${CHART_NS}}dPt`;
const INVERT_IF_NEGATIVE_TAG = `{${CHART_NS}}invertIfNegative`;
const EXPLOSION_TAG = `{${CHART_NS}}explosion`;
const HOLE_SIZE_TAG = `{${CHART_NS}}holeSize`;
const FIRST_SLICE_ANG_TAG = `{${CHART_NS}}firstSliceAng`;
const SCATTER_STYLE_TAG = `{${CHART_NS}}scatterStyle`;
const RADAR_STYLE_TAG = `{${CHART_NS}}radarStyle`;
const X_VAL_TAG = `{${CHART_NS}}xVal`;
const Y_VAL_TAG = `{${CHART_NS}}yVal`;
const BUBBLE_CHART_TAG = `{${CHART_NS}}bubbleChart`;
const STOCK_CHART_TAG = `{${CHART_NS}}stockChart`;
const SURFACE_CHART_TAG = `{${CHART_NS}}surfaceChart`;
const BUBBLE_SIZE_TAG = `{${CHART_NS}}bubbleSize`;
const BUBBLE_3D_TAG = `{${CHART_NS}}bubble3D`;
const BUBBLE_SCALE_TAG = `{${CHART_NS}}bubbleScale`;
const SHOW_NEG_BUBBLES_TAG = `{${CHART_NS}}showNegBubbles`;
const SIZE_REPRESENTS_TAG = `{${CHART_NS}}sizeRepresents`;
const HI_LOW_LINES_TAG = `{${CHART_NS}}hiLowLines`;
const UP_DOWN_BARS_TAG = `{${CHART_NS}}upDownBars`;
const WIREFRAME_TAG = `{${CHART_NS}}wireframe`;
const OF_PIE_CHART_TAG = `{${CHART_NS}}ofPieChart`;
const OF_PIE_TYPE_TAG = `{${CHART_NS}}ofPieType`;
const SPLIT_TYPE_TAG = `{${CHART_NS}}splitType`;
const SPLIT_POS_TAG = `{${CHART_NS}}splitPos`;
const CUST_SPLIT_TAG = `{${CHART_NS}}custSplit`;
const SECOND_PIE_SIZE_TAG = `{${CHART_NS}}secondPieSize`;
const SEC_BLOCK_PT_TAG = `{${CHART_NS}}secondaryPt`;
const BAR3D_CHART_TAG = `{${CHART_NS}}bar3DChart`;
const LINE3D_CHART_TAG = `{${CHART_NS}}line3DChart`;
const PIE3D_CHART_TAG = `{${CHART_NS}}pie3DChart`;
const AREA3D_CHART_TAG = `{${CHART_NS}}area3DChart`;
const SURFACE3D_CHART_TAG = `{${CHART_NS}}surface3DChart`;
const GAP_DEPTH_TAG = `{${CHART_NS}}gapDepth`;
const SHAPE_TAG = `{${CHART_NS}}shape`;
const AX_POS_TAG = `{${CHART_NS}}axPos`;
const CROSS_AX_TAG = `{${CHART_NS}}crossAx`;
const MAJOR_GRIDLINES_TAG = `{${CHART_NS}}majorGridlines`;
const MINOR_GRIDLINES_TAG = `{${CHART_NS}}minorGridlines`;
const SCALING_TAG = `{${CHART_NS}}scaling`;
const ORIENTATION_TAG = `{${CHART_NS}}orientation`;
const MIN_TAG = `{${CHART_NS}}min`;
const MAX_TAG = `{${CHART_NS}}max`;
const LOG_BASE_TAG = `{${CHART_NS}}logBase`;
const CROSSES_TAG = `{${CHART_NS}}crosses`;
const CROSSES_AT_TAG = `{${CHART_NS}}crossesAt`;
const CROSS_BETWEEN_TAG = `{${CHART_NS}}crossBetween`;
const MAJOR_UNIT_TAG = `{${CHART_NS}}majorUnit`;
const MINOR_UNIT_TAG = `{${CHART_NS}}minorUnit`;
const MAJOR_TICK_MARK_TAG = `{${CHART_NS}}majorTickMark`;
const MINOR_TICK_MARK_TAG = `{${CHART_NS}}minorTickMark`;
const TICK_LBL_POS_TAG = `{${CHART_NS}}tickLblPos`;
const AUTO_TAG = `{${CHART_NS}}auto`;
const LBL_ALGN_TAG = `{${CHART_NS}}lblAlgn`;
const LBL_OFFSET_TAG = `{${CHART_NS}}lblOffset`;
const NO_MULTI_LVL_LBL_TAG = `{${CHART_NS}}noMultiLvlLbl`;
const VIEW3D_TAG = `{${CHART_NS}}view3D`;
const ROT_X_TAG = `{${CHART_NS}}rotX`;
const ROT_Y_TAG = `{${CHART_NS}}rotY`;
const H_PERCENT_TAG = `{${CHART_NS}}hPercent`;
const DEPTH_PERCENT_TAG = `{${CHART_NS}}depthPercent`;
const R_ANG_AX_TAG = `{${CHART_NS}}rAngAx`;
const PERSPECTIVE_TAG = `{${CHART_NS}}perspective`;
const FLOOR_TAG = `{${CHART_NS}}floor`;
const SIDE_WALL_TAG = `{${CHART_NS}}sideWall`;
const BACK_WALL_TAG = `{${CHART_NS}}backWall`;
const THICKNESS_TAG = `{${CHART_NS}}thickness`;
const LAYOUT_TAG = `{${CHART_NS}}layout`;
const MANUAL_LAYOUT_TAG = `{${CHART_NS}}manualLayout`;
const LAYOUT_TARGET_TAG = `{${CHART_NS}}layoutTarget`;
const X_MODE_TAG = `{${CHART_NS}}xMode`;
const Y_MODE_TAG = `{${CHART_NS}}yMode`;
const W_MODE_TAG = `{${CHART_NS}}wMode`;
const H_MODE_TAG = `{${CHART_NS}}hMode`;
const X_TAG = `{${CHART_NS}}x`;
const Y_TAG = `{${CHART_NS}}y`;
const W_TAG = `{${CHART_NS}}w`;
const H_TAG = `{${CHART_NS}}h`;
const LEGEND_TAG = `{${CHART_NS}}legend`;
const LEGEND_POS_TAG = `{${CHART_NS}}legendPos`;
const PLOT_VIS_ONLY_TAG = `{${CHART_NS}}plotVisOnly`;
const DISP_BLANKS_AS_TAG = `{${CHART_NS}}dispBlanksAs`;
const SP_PR_TAG = `{${CHART_NS}}spPr`;
const TX_PR_TAG = `{${CHART_NS}}txPr`;
const OVERLAY_TAG = `{${CHART_NS}}overlay`;
const USER_SHAPES_TAG = `{${CHART_NS}}userShapes`;
const D_LBLS_TAG = `{${CHART_NS}}dLbls`;
const D_LBL_TAG = `{${CHART_NS}}dLbl`;
const NUM_FMT_TAG = `{${CHART_NS}}numFmt`;
const D_LBL_POS_TAG = `{${CHART_NS}}dLblPos`;
const SHOW_LEGEND_KEY_TAG = `{${CHART_NS}}showLegendKey`;
const SHOW_VAL_TAG = `{${CHART_NS}}showVal`;
const SHOW_CAT_NAME_TAG = `{${CHART_NS}}showCatName`;
const SHOW_SER_NAME_TAG = `{${CHART_NS}}showSerName`;
const SHOW_PERCENT_TAG = `{${CHART_NS}}showPercent`;
const SHOW_BUBBLE_SIZE_TAG = `{${CHART_NS}}showBubbleSize`;
const SHOW_LEADER_LINES_TAG = `{${CHART_NS}}showLeaderLines`;
const SEPARATOR_TAG = `{${CHART_NS}}separator`;
const DELETE_LBL_TAG = `{${CHART_NS}}delete`;
const TRENDLINE_TAG = `{${CHART_NS}}trendline`;
const TRENDLINE_TYPE_TAG = `{${CHART_NS}}trendlineType`;
const TRENDLINE_NAME_TAG = `{${CHART_NS}}name`;
const TRENDLINE_ORDER_TAG = `{${CHART_NS}}order`;
const TRENDLINE_PERIOD_TAG = `{${CHART_NS}}period`;
const TRENDLINE_FORWARD_TAG = `{${CHART_NS}}forward`;
const TRENDLINE_BACKWARD_TAG = `{${CHART_NS}}backward`;
const TRENDLINE_INTERCEPT_TAG = `{${CHART_NS}}intercept`;
const TRENDLINE_DISP_RSQR_TAG = `{${CHART_NS}}dispRSqr`;
const TRENDLINE_DISP_EQ_TAG = `{${CHART_NS}}dispEq`;
const ERR_BARS_TAG = `{${CHART_NS}}errBars`;
const ERR_DIR_TAG = `{${CHART_NS}}errDir`;
const ERR_BAR_TYPE_TAG = `{${CHART_NS}}errBarType`;
const ERR_VAL_TYPE_TAG = `{${CHART_NS}}errValType`;
const NO_END_CAP_TAG = `{${CHART_NS}}noEndCap`;
const PLUS_TAG = `{${CHART_NS}}plus`;
const MINUS_TAG = `{${CHART_NS}}minus`;
const A_R_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/main}r';
const A_T_TAG = '{http://schemas.openxmlformats.org/drawingml/2006/main}t';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeText = escapeXmlText;

const valAttr = (n: XmlNode | undefined): string | undefined => n?.attrs['val'];
const intVal = (n: XmlNode | undefined): number | undefined => {
  const v = valAttr(n);
  if (v === undefined) return undefined;
  const x = Number.parseInt(v, 10);
  return Number.isInteger(x) ? x : undefined;
};
const boolVal = (n: XmlNode | undefined): boolean | undefined => {
  const v = valAttr(n);
  if (v === undefined) return undefined;
  if (v === '1' || v === 'true') return true;
  if (v === '0' || v === 'false') return false;
  return undefined;
};

const parseRichTextString = (richEl: XmlNode): string | undefined => {
  let out = '';
  for (const p of richEl.children) {
    if (typeof p === 'string') continue;
    for (const r of findChildren(p, A_R_TAG)) {
      const t = findChild(r, A_T_TAG);
      if (t?.text) out += t.text;
    }
  }
  return out.length > 0 ? out : undefined;
};

const parseChartTitle = (titleEl: XmlNode): ChartTitle => {
  const out: ChartTitle = {};
  const tx = findChild(titleEl, TX_TAG);
  if (tx) {
    const rich = findChild(tx, RICH_TAG);
    if (rich) {
      // Parse the rich text body fully so callers see the formatted content.
      // For the convenience plain-text shortcut we also run a simple scrape
      // when the body has only literal runs.
      const body = parseTextBody(rich);
      out.tx = body;
      const flat = parseRichTextString(rich);
      if (flat !== undefined) out.text = flat;
    }
  }
  const overlay = boolVal(findChild(titleEl, OVERLAY_TAG));
  if (overlay !== undefined) out.overlay = overlay;
  const layoutEl = findChild(titleEl, LAYOUT_TAG);
  if (layoutEl) {
    const layout = parseLayout(layoutEl);
    if (layout) out.layout = layout;
  }
  const spPrEl = findChild(titleEl, SP_PR_TAG);
  if (spPrEl) out.spPr = parseShapeProperties(spPrEl);
  const txPrEl = findChild(titleEl, TX_PR_TAG);
  if (txPrEl) out.txPr = parseTextBody(txPrEl);
  return out;
};

const parseSpPrSlot = (parent: XmlNode): ShapeProperties | undefined => {
  const el = findChild(parent, SP_PR_TAG);
  return el ? parseShapeProperties(el) : undefined;
};

const parseTxPrSlot = (parent: XmlNode): TextBody | undefined => {
  const el = findChild(parent, TX_PR_TAG);
  return el ? parseTextBody(el) : undefined;
};

// ---- Series decorations (dLbls, trendlines, error bars) -------------------

const VALID_D_LBL_POS: ReadonlyArray<string> = ['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't'];

const parseNumberFormat = (el: XmlNode | undefined): ChartNumberFormat | undefined => {
  if (!el) return undefined;
  const formatCode = el.attrs['formatCode'];
  if (formatCode === undefined) return undefined;
  const sourceLinkedRaw = el.attrs['sourceLinked'];
  const sourceLinked =
    sourceLinkedRaw === '1' || sourceLinkedRaw === 'true'
      ? true
      : sourceLinkedRaw === '0' || sourceLinkedRaw === 'false'
        ? false
        : undefined;
  return { formatCode, ...(sourceLinked !== undefined ? { sourceLinked } : {}) };
};

const serializeNumberFormat = (n: ChartNumberFormat): string => {
  const sl = n.sourceLinked !== undefined ? ` sourceLinked="${n.sourceLinked ? '1' : '0'}"` : '';
  return `<c:numFmt formatCode="${escapeText(n.formatCode)}"${sl}/>`;
};

interface DataLabelCommon {
  delete?: boolean;
  numFmt?: ChartNumberFormat;
  spPr?: ShapeProperties;
  txPr?: TextBody;
  dLblPos?: DataLabelPosition;
  showLegendKey?: boolean;
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  separator?: string;
}

const parseDataLabelCommon = (el: XmlNode): DataLabelCommon => {
  const out: DataLabelCommon = {};
  const del = boolVal(findChild(el, DELETE_LBL_TAG));
  if (del !== undefined) out.delete = del;
  const numFmt = parseNumberFormat(findChild(el, NUM_FMT_TAG));
  if (numFmt) out.numFmt = numFmt;
  const spPr = parseSpPrSlot(el);
  if (spPr) out.spPr = spPr;
  const txPr = parseTxPrSlot(el);
  if (txPr) out.txPr = txPr;
  const posRaw = valAttr(findChild(el, D_LBL_POS_TAG));
  if (posRaw && VALID_D_LBL_POS.includes(posRaw)) out.dLblPos = posRaw as DataLabelPosition;
  const showLegendKey = boolVal(findChild(el, SHOW_LEGEND_KEY_TAG));
  if (showLegendKey !== undefined) out.showLegendKey = showLegendKey;
  const showVal = boolVal(findChild(el, SHOW_VAL_TAG));
  if (showVal !== undefined) out.showVal = showVal;
  const showCatName = boolVal(findChild(el, SHOW_CAT_NAME_TAG));
  if (showCatName !== undefined) out.showCatName = showCatName;
  const showSerName = boolVal(findChild(el, SHOW_SER_NAME_TAG));
  if (showSerName !== undefined) out.showSerName = showSerName;
  const showPercent = boolVal(findChild(el, SHOW_PERCENT_TAG));
  if (showPercent !== undefined) out.showPercent = showPercent;
  const showBubbleSize = boolVal(findChild(el, SHOW_BUBBLE_SIZE_TAG));
  if (showBubbleSize !== undefined) out.showBubbleSize = showBubbleSize;
  const sep = findChild(el, SEPARATOR_TAG);
  if (sep?.text !== undefined) out.separator = sep.text;
  return out;
};

const serializeDataLabelCommon = (d: DataLabelCommon): string => {
  // ECMA-376 element ordering inside <c:dLbl>/<c:dLbls>: numFmt → spPr → txPr →
  // dLblPos → showLegendKey → showVal → showCatName → showSerName → showPercent
  // → showBubbleSize → separator → showLeaderLines (delete is exclusive — when
  // true the label has no other children).
  if (d.delete) return '<c:delete val="1"/>';
  const parts: string[] = [];
  if (d.numFmt) parts.push(serializeNumberFormat(d.numFmt));
  if (d.spPr) parts.push(serializeShapeProperties(d.spPr));
  if (d.txPr) parts.push(serializeTextBody(d.txPr, 'c:txPr'));
  if (d.dLblPos) parts.push(`<c:dLblPos val="${d.dLblPos}"/>`);
  if (d.showLegendKey !== undefined) parts.push(`<c:showLegendKey val="${d.showLegendKey ? '1' : '0'}"/>`);
  if (d.showVal !== undefined) parts.push(`<c:showVal val="${d.showVal ? '1' : '0'}"/>`);
  if (d.showCatName !== undefined) parts.push(`<c:showCatName val="${d.showCatName ? '1' : '0'}"/>`);
  if (d.showSerName !== undefined) parts.push(`<c:showSerName val="${d.showSerName ? '1' : '0'}"/>`);
  if (d.showPercent !== undefined) parts.push(`<c:showPercent val="${d.showPercent ? '1' : '0'}"/>`);
  if (d.showBubbleSize !== undefined) parts.push(`<c:showBubbleSize val="${d.showBubbleSize ? '1' : '0'}"/>`);
  if (d.separator !== undefined) parts.push(`<c:separator>${escapeText(d.separator)}</c:separator>`);
  return parts.join('');
};

const parseDataLabel = (el: XmlNode): DataLabel | undefined => {
  const idx = intVal(findChild(el, IDX_TAG));
  if (idx === undefined) return undefined;
  const out: DataLabel = { idx, ...parseDataLabelCommon(el) };
  // <c:tx>: either <c:rich> (TextBody) or <c:strRef><c:f>.
  const txEl = findChild(el, TX_TAG);
  if (txEl) {
    const richEl = findChild(txEl, RICH_TAG);
    if (richEl) {
      out.tx = { kind: 'rich', body: parseTextBody(richEl) };
    } else {
      const strRef = findChild(txEl, STR_REF_TAG);
      if (strRef) {
        const ref = findChild(strRef, F_TAG)?.text ?? '';
        out.tx = { kind: 'strRef', ref };
      }
    }
  }
  return out;
};

const serializeDataLabel = (d: DataLabel): string => {
  if (d.delete) return `<c:dLbl><c:idx val="${d.idx}"/><c:delete val="1"/></c:dLbl>`;
  const parts: string[] = [`<c:dLbl><c:idx val="${d.idx}"/>`];
  if (d.tx) {
    if (d.tx.kind === 'rich') {
      parts.push(`<c:tx>${serializeTextBody(d.tx.body, 'c:rich')}</c:tx>`);
    } else {
      parts.push(`<c:tx><c:strRef><c:f>${escapeText(d.tx.ref)}</c:f></c:strRef></c:tx>`);
    }
  }
  parts.push(serializeDataLabelCommon(d));
  parts.push('</c:dLbl>');
  return parts.join('');
};

const parseDataLabelList = (el: XmlNode): DataLabelList => {
  const common = parseDataLabelCommon(el);
  const out: DataLabelList = { ...common };
  const showLeader = boolVal(findChild(el, SHOW_LEADER_LINES_TAG));
  if (showLeader !== undefined) out.showLeaderLines = showLeader;
  const dLbls: DataLabel[] = [];
  for (const child of findChildren(el, D_LBL_TAG)) {
    const dl = parseDataLabel(child);
    if (dl) dLbls.push(dl);
  }
  if (dLbls.length > 0) out.dLbl = dLbls;
  return out;
};

const serializeDataLabelList = (d: DataLabelList): string => {
  if (d.delete) return '<c:dLbls><c:delete val="1"/></c:dLbls>';
  const parts: string[] = ['<c:dLbls>'];
  if (d.dLbl) for (const dl of d.dLbl) parts.push(serializeDataLabel(dl));
  parts.push(serializeDataLabelCommon(d));
  if (d.showLeaderLines !== undefined) {
    parts.push(`<c:showLeaderLines val="${d.showLeaderLines ? '1' : '0'}"/>`);
  }
  parts.push('</c:dLbls>');
  return parts.join('');
};

const VALID_MARKER_SYMBOLS: ReadonlyArray<string> = [
  'auto',
  'circle',
  'dash',
  'diamond',
  'dot',
  'none',
  'picture',
  'plus',
  'square',
  'star',
  'triangle',
  'x',
];

const parseMarker = (el: XmlNode): Marker | undefined => {
  const symbolRaw = valAttr(findChild(el, SYMBOL_TAG));
  const size = intVal(findChild(el, SIZE_TAG));
  const spPr = parseSpPrSlot(el);
  const out: Marker = {};
  if (symbolRaw && VALID_MARKER_SYMBOLS.includes(symbolRaw)) out.symbol = symbolRaw as MarkerSymbol;
  if (size !== undefined) out.size = size;
  if (spPr) out.spPr = spPr;
  return Object.keys(out).length > 0 ? out : undefined;
};

const serializeMarker = (m: Marker): string => {
  const parts: string[] = ['<c:marker>'];
  if (m.symbol !== undefined) parts.push(`<c:symbol val="${m.symbol}"/>`);
  if (m.size !== undefined) parts.push(`<c:size val="${m.size}"/>`);
  if (m.spPr) parts.push(serializeShapeProperties(m.spPr));
  parts.push('</c:marker>');
  return parts.join('');
};

const parseDataPoint = (el: XmlNode): DataPoint | undefined => {
  const idx = intVal(findChild(el, IDX_TAG));
  if (idx === undefined) return undefined;
  const invertIfNegative = boolVal(findChild(el, INVERT_IF_NEGATIVE_TAG));
  const bubble3D = boolVal(findChild(el, BUBBLE_3D_TAG));
  const explosion = intVal(findChild(el, EXPLOSION_TAG));
  const markerEl = findChild(el, MARKER_TAG);
  const marker = markerEl ? parseMarker(markerEl) : undefined;
  const spPr = parseSpPrSlot(el);
  const out: DataPoint = { idx };
  if (invertIfNegative !== undefined) out.invertIfNegative = invertIfNegative;
  if (marker) out.marker = marker;
  if (bubble3D !== undefined) out.bubble3D = bubble3D;
  if (explosion !== undefined) out.explosion = explosion;
  if (spPr) out.spPr = spPr;
  return out;
};

const serializeDataPoint = (p: DataPoint): string => {
  // ECMA-376 order inside <c:dPt>: idx → invertIfNegative? → marker? → bubble3D?
  //   → explosion? → spPr? → pictureOptions?
  const parts: string[] = ['<c:dPt>', `<c:idx val="${p.idx}"/>`];
  if (p.invertIfNegative !== undefined) {
    parts.push(`<c:invertIfNegative val="${p.invertIfNegative ? '1' : '0'}"/>`);
  }
  if (p.marker) parts.push(serializeMarker(p.marker));
  if (p.bubble3D !== undefined) parts.push(`<c:bubble3D val="${p.bubble3D ? '1' : '0'}"/>`);
  if (p.explosion !== undefined) parts.push(`<c:explosion val="${p.explosion}"/>`);
  if (p.spPr) parts.push(serializeShapeProperties(p.spPr));
  parts.push('</c:dPt>');
  return parts.join('');
};

const parseDataPointList = (serEl: XmlNode): DataPoint[] | undefined => {
  const out: DataPoint[] = [];
  for (const el of findChildren(serEl, D_PT_TAG)) {
    const p = parseDataPoint(el);
    if (p) out.push(p);
  }
  return out.length > 0 ? out : undefined;
};

const VALID_TRENDLINE_TYPES: ReadonlyArray<string> = ['exp', 'linear', 'log', 'movingAvg', 'poly', 'power'];

const parseTrendline = (el: XmlNode): Trendline | undefined => {
  const typeRaw = valAttr(findChild(el, TRENDLINE_TYPE_TAG));
  if (!typeRaw || !VALID_TRENDLINE_TYPES.includes(typeRaw)) return undefined;
  const out: Trendline = { trendlineType: typeRaw as TrendlineType };
  const nameEl = findChild(el, TRENDLINE_NAME_TAG);
  if (nameEl?.text !== undefined) out.name = nameEl.text;
  const spPr = parseSpPrSlot(el);
  if (spPr) out.spPr = spPr;
  const order = intVal(findChild(el, TRENDLINE_ORDER_TAG));
  if (order !== undefined) out.order = order;
  const period = intVal(findChild(el, TRENDLINE_PERIOD_TAG));
  if (period !== undefined) out.period = period;
  // forward / backward / intercept can be floats.
  const fwdEl = findChild(el, TRENDLINE_FORWARD_TAG);
  if (fwdEl) {
    const v = valAttr(fwdEl);
    if (v !== undefined) {
      const n = Number.parseFloat(v);
      if (Number.isFinite(n)) out.forward = n;
    }
  }
  const bwdEl = findChild(el, TRENDLINE_BACKWARD_TAG);
  if (bwdEl) {
    const v = valAttr(bwdEl);
    if (v !== undefined) {
      const n = Number.parseFloat(v);
      if (Number.isFinite(n)) out.backward = n;
    }
  }
  const intEl = findChild(el, TRENDLINE_INTERCEPT_TAG);
  if (intEl) {
    const v = valAttr(intEl);
    if (v !== undefined) {
      const n = Number.parseFloat(v);
      if (Number.isFinite(n)) out.intercept = n;
    }
  }
  const dispRSqr = boolVal(findChild(el, TRENDLINE_DISP_RSQR_TAG));
  if (dispRSqr !== undefined) out.dispRSqr = dispRSqr;
  const dispEq = boolVal(findChild(el, TRENDLINE_DISP_EQ_TAG));
  if (dispEq !== undefined) out.dispEq = dispEq;
  return out;
};

const serializeTrendline = (t: Trendline): string => {
  // ECMA-376: name → spPr → trendlineType → order → period → forward → backward
  // → intercept → dispRSqr → dispEq → trendlineLbl.
  const parts: string[] = ['<c:trendline>'];
  if (t.name !== undefined) parts.push(`<c:name>${escapeText(t.name)}</c:name>`);
  if (t.spPr) parts.push(serializeShapeProperties(t.spPr));
  parts.push(`<c:trendlineType val="${t.trendlineType}"/>`);
  if (t.order !== undefined) parts.push(`<c:order val="${t.order}"/>`);
  if (t.period !== undefined) parts.push(`<c:period val="${t.period}"/>`);
  if (t.forward !== undefined) parts.push(`<c:forward val="${t.forward}"/>`);
  if (t.backward !== undefined) parts.push(`<c:backward val="${t.backward}"/>`);
  if (t.intercept !== undefined) parts.push(`<c:intercept val="${t.intercept}"/>`);
  if (t.dispRSqr !== undefined) parts.push(`<c:dispRSqr val="${t.dispRSqr ? '1' : '0'}"/>`);
  if (t.dispEq !== undefined) parts.push(`<c:dispEq val="${t.dispEq ? '1' : '0'}"/>`);
  parts.push('</c:trendline>');
  return parts.join('');
};

const VALID_ERR_DIR: ReadonlyArray<string> = ['x', 'y'];
const VALID_ERR_BAR_TYPE: ReadonlyArray<string> = ['both', 'minus', 'plus'];
const VALID_ERR_VAL_TYPE: ReadonlyArray<string> = ['cust', 'fixedVal', 'percentage', 'stdDev', 'stdErr'];

const parseErrBars = (el: XmlNode): ErrorBars | undefined => {
  const typeRaw = valAttr(findChild(el, ERR_BAR_TYPE_TAG));
  const valRaw = valAttr(findChild(el, ERR_VAL_TYPE_TAG));
  if (!typeRaw || !valRaw || !VALID_ERR_BAR_TYPE.includes(typeRaw) || !VALID_ERR_VAL_TYPE.includes(valRaw)) {
    return undefined;
  }
  const out: ErrorBars = {
    errBarType: typeRaw as ErrorBarType,
    errValType: valRaw as ErrorValType,
  };
  const dirRaw = valAttr(findChild(el, ERR_DIR_TAG));
  if (dirRaw && VALID_ERR_DIR.includes(dirRaw)) out.errDir = dirRaw as ErrorBarDirection;
  const noEnd = boolVal(findChild(el, NO_END_CAP_TAG));
  if (noEnd !== undefined) out.noEndCap = noEnd;
  const valEl = findChild(el, VAL_TAG);
  if (valEl) {
    const v = valAttr(valEl);
    if (v !== undefined) {
      const n = Number.parseFloat(v);
      if (Number.isFinite(n)) out.val = n;
    }
  }
  const plus = parseNumericRef(el, PLUS_TAG);
  if (plus) out.plus = plus;
  const minus = parseNumericRef(el, MINUS_TAG);
  if (minus) out.minus = minus;
  const spPr = parseSpPrSlot(el);
  if (spPr) out.spPr = spPr;
  return out;
};

const serializeErrBars = (e: ErrorBars): string => {
  // ECMA-376: errDir? → errBarType → errValType → noEndCap? → plus? → minus? →
  // val? → spPr?.
  const parts: string[] = ['<c:errBars>'];
  if (e.errDir) parts.push(`<c:errDir val="${e.errDir}"/>`);
  parts.push(`<c:errBarType val="${e.errBarType}"/>`);
  parts.push(`<c:errValType val="${e.errValType}"/>`);
  if (e.noEndCap !== undefined) parts.push(`<c:noEndCap val="${e.noEndCap ? '1' : '0'}"/>`);
  if (e.plus) parts.push(serializeNumericRef('plus', e.plus));
  if (e.minus) parts.push(serializeNumericRef('minus', e.minus));
  if (e.val !== undefined) parts.push(`<c:val val="${e.val}"/>`);
  if (e.spPr) parts.push(serializeShapeProperties(e.spPr));
  parts.push('</c:errBars>');
  return parts.join('');
};

/** Pull dLbls / trendline[] / errBars[] off a `<c:ser>` element. */
const parseSeriesDecorations = (
  serEl: XmlNode,
): { dLbls?: DataLabelList; trendline?: Trendline[]; errBars?: ErrorBars[] } => {
  const out: { dLbls?: DataLabelList; trendline?: Trendline[]; errBars?: ErrorBars[] } = {};
  const dLblsEl = findChild(serEl, D_LBLS_TAG);
  if (dLblsEl) out.dLbls = parseDataLabelList(dLblsEl);
  const trends: Trendline[] = [];
  for (const t of findChildren(serEl, TRENDLINE_TAG)) {
    const tl = parseTrendline(t);
    if (tl) trends.push(tl);
  }
  if (trends.length > 0) out.trendline = trends;
  const ebs: ErrorBars[] = [];
  for (const eb of findChildren(serEl, ERR_BARS_TAG)) {
    const e = parseErrBars(eb);
    if (e) ebs.push(e);
  }
  if (ebs.length > 0) out.errBars = ebs;
  return out;
};

const parseNumCache = (cacheEl: XmlNode): { values: number[]; formatCode?: string } => {
  const values: number[] = [];
  const fmt = findChild(cacheEl, FORMAT_CODE_TAG)?.text;
  for (const pt of findChildren(cacheEl, PT_TAG)) {
    const v = findChild(pt, V_TAG)?.text;
    if (v !== undefined) values.push(Number.parseFloat(v));
  }
  return fmt !== undefined ? { values, formatCode: fmt } : { values };
};

const parseStrCache = (cacheEl: XmlNode): string[] => {
  const values: string[] = [];
  for (const pt of findChildren(cacheEl, PT_TAG)) {
    const v = findChild(pt, V_TAG)?.text;
    if (v !== undefined) values.push(v);
  }
  return values;
};

const parseNumericRef = (parent: XmlNode, tag: string): NumericRef | undefined => {
  const wrap = findChild(parent, tag);
  if (!wrap) return undefined;
  const numRef = findChild(wrap, NUM_REF_TAG);
  if (!numRef) return undefined;
  const ref = findChild(numRef, F_TAG)?.text ?? '';
  const cacheEl = findChild(numRef, NUM_CACHE_TAG);
  if (!cacheEl) return { ref };
  const { values, formatCode } = parseNumCache(cacheEl);
  return {
    ref,
    cache: values,
    ...(formatCode !== undefined ? { formatCode } : {}),
  };
};

const parseCategoryRef = (parent: XmlNode): CategoryRef | undefined => {
  const cat = findChild(parent, CAT_TAG);
  if (!cat) return undefined;
  const numRef = findChild(cat, NUM_REF_TAG);
  if (numRef) {
    const ref = findChild(numRef, F_TAG)?.text ?? '';
    const cacheEl = findChild(numRef, NUM_CACHE_TAG);
    if (!cacheEl) return { ref, cacheKind: 'num' };
    const { values, formatCode } = parseNumCache(cacheEl);
    return {
      ref,
      cacheKind: 'num',
      cache: values,
      ...(formatCode !== undefined ? { formatCode } : {}),
    };
  }
  const strRef = findChild(cat, STR_REF_TAG);
  if (strRef) {
    const ref = findChild(strRef, F_TAG)?.text ?? '';
    const cacheEl = findChild(strRef, STR_CACHE_TAG);
    if (!cacheEl) return { ref, cacheKind: 'str' };
    return { ref, cacheKind: 'str', cache: parseStrCache(cacheEl) };
  }
  return undefined;
};

/** `<c:tx>` of a series: a cell reference, or a typed name (`<c:v>`, or a cache with no formula). */
const parseSeriesTx = (serEl: XmlNode): BarSeries['tx'] => {
  const txEl = findChild(serEl, TX_TAG);
  if (!txEl) return undefined;
  const literal = findChild(txEl, V_TAG)?.text;
  if (literal !== undefined) return { kind: 'literal', value: literal };
  const strRef = findChild(txEl, STR_REF_TAG);
  if (!strRef) return undefined;
  const ref = findChild(strRef, F_TAG)?.text ?? '';
  if (ref !== '') return { kind: 'ref', ref };
  const cacheEl = findChild(strRef, STR_CACHE_TAG);
  const cached = cacheEl ? parseStrCache(cacheEl)[0] : undefined;
  return cached === undefined ? undefined : { kind: 'literal', value: cached };
};

const parseSeries = (serEl: XmlNode): BarSeries | undefined => {
  const idx = intVal(findChild(serEl, IDX_TAG));
  const order = intVal(findChild(serEl, ORDER_TAG));
  if (idx === undefined) return undefined;
  const val = parseNumericRef(serEl, VAL_TAG);
  if (!val) return undefined;
  const opts: Parameters<typeof makeBarSeries>[0] = {
    idx,
    val,
  };
  if (order !== undefined) opts.order = order;
  const cat = parseCategoryRef(serEl);
  if (cat) opts.cat = cat;
  const tx = parseSeriesTx(serEl);
  if (tx) opts.tx = tx;
  const base = makeBarSeries(opts);
  const spPr = parseSpPrSlot(serEl);
  const invertIfNegative = boolVal(findChild(serEl, INVERT_IF_NEGATIVE_TAG));
  const explosion = intVal(findChild(serEl, EXPLOSION_TAG));
  const dPt = parseDataPointList(serEl);
  const deco = parseSeriesDecorations(serEl);
  return {
    ...base,
    ...(spPr ? { spPr } : {}),
    ...(invertIfNegative !== undefined ? { invertIfNegative } : {}),
    ...(explosion !== undefined ? { explosion } : {}),
    ...(dPt ? { dPt } : {}),
    ...deco,
  };
};

const parseAxIds = (chartEl: XmlNode): [number, number] => {
  const axIdNodes = findChildren(chartEl, AX_ID_TAG);
  return [intVal(axIdNodes[0]) ?? 1, intVal(axIdNodes[1]) ?? 2];
};

const parseBarSeriesList = (chartEl: XmlNode): BarSeries[] => {
  const series: BarSeries[] = [];
  for (const ser of findChildren(chartEl, SER_TAG)) {
    const s = parseSeries(ser);
    if (s) series.push(s);
  }
  return series;
};

const parseBarChart = (barEl: XmlNode): BarChart => {
  const barDir = (valAttr(findChild(barEl, BAR_DIR_TAG)) ?? 'col') as BarDirection;
  const grouping = (valAttr(findChild(barEl, GROUPING_TAG)) ?? 'clustered') as GroupingType;
  const varyColors = boolVal(findChild(barEl, VARY_COLORS_TAG));
  const gapWidth = intVal(findChild(barEl, GAP_WIDTH_TAG));
  const overlap = intVal(findChild(barEl, OVERLAP_TAG));
  return makeBarChart({
    barDir,
    grouping,
    series: parseBarSeriesList(barEl),
    axIds: parseAxIds(barEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapWidth !== undefined ? { gapWidth } : {}),
    ...(overlap !== undefined ? { overlap } : {}),
  });
};

const parseLineSeries = (serEl: XmlNode): LineSeries | undefined => {
  const base = parseSeries(serEl);
  if (!base) return undefined;
  const smooth = boolVal(findChild(serEl, SMOOTH_TAG));
  const markerEl = findChild(serEl, MARKER_TAG);
  const marker = markerEl ? parseMarker(markerEl) : undefined;
  return {
    ...base,
    ...(smooth !== undefined ? { smooth } : {}),
    ...(marker ? { marker } : {}),
  };
};

const parseLineChart = (lineEl: XmlNode): LineChart => {
  const grouping = (valAttr(findChild(lineEl, GROUPING_TAG)) ?? 'standard') as GroupingType;
  const varyColors = boolVal(findChild(lineEl, VARY_COLORS_TAG));
  const smooth = boolVal(findChild(lineEl, SMOOTH_TAG));
  const series: LineSeries[] = [];
  for (const ser of findChildren(lineEl, SER_TAG)) {
    const s = parseLineSeries(ser);
    if (s) series.push(s);
  }
  return makeLineChart({
    grouping,
    series,
    axIds: parseAxIds(lineEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(smooth !== undefined ? { smooth } : {}),
  });
};

const parseAreaChart = (areaEl: XmlNode): AreaChart => {
  const grouping = (valAttr(findChild(areaEl, GROUPING_TAG)) ?? 'standard') as GroupingType;
  const varyColors = boolVal(findChild(areaEl, VARY_COLORS_TAG));
  return makeAreaChart({
    grouping,
    series: parseBarSeriesList(areaEl),
    axIds: parseAxIds(areaEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
  });
};

const parsePieChart = (pieEl: XmlNode): PieChart => {
  const varyColors = boolVal(findChild(pieEl, VARY_COLORS_TAG));
  const firstSliceAng = intVal(findChild(pieEl, FIRST_SLICE_ANG_TAG));
  return makePieChart({
    series: parseBarSeriesList(pieEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(firstSliceAng !== undefined ? { firstSliceAng } : {}),
  });
};

const parseDoughnutChart = (donutEl: XmlNode): DoughnutChart => {
  const varyColors = boolVal(findChild(donutEl, VARY_COLORS_TAG));
  const holeSize = intVal(findChild(donutEl, HOLE_SIZE_TAG));
  const firstSliceAng = intVal(findChild(donutEl, FIRST_SLICE_ANG_TAG));
  return makeDoughnutChart({
    series: parseBarSeriesList(donutEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(holeSize !== undefined ? { holeSize } : {}),
    ...(firstSliceAng !== undefined ? { firstSliceAng } : {}),
  });
};

const parseScatterSeries = (serEl: XmlNode): ScatterSeries | undefined => {
  const idx = intVal(findChild(serEl, IDX_TAG));
  const order = intVal(findChild(serEl, ORDER_TAG));
  if (idx === undefined) return undefined;
  const yVal = parseNumericRef(serEl, Y_VAL_TAG);
  if (!yVal) return undefined;
  const xVal = parseNumericRef(serEl, X_VAL_TAG);
  const smooth = boolVal(findChild(serEl, SMOOTH_TAG));
  const opts: Parameters<typeof makeScatterSeries>[0] = { idx, yVal };
  if (order !== undefined) opts.order = order;
  if (xVal) opts.xVal = xVal;
  if (smooth !== undefined) opts.smooth = smooth;
  const markerEl = findChild(serEl, MARKER_TAG);
  const marker = markerEl ? parseMarker(markerEl) : undefined;
  if (marker) opts.marker = marker;
  const tx = parseSeriesTx(serEl);
  if (tx) opts.tx = tx;
  const base = makeScatterSeries(opts);
  const spPr = parseSpPrSlot(serEl);
  const dPt = parseDataPointList(serEl);
  const deco = parseSeriesDecorations(serEl);
  return {
    ...base,
    ...(spPr ? { spPr } : {}),
    ...(dPt ? { dPt } : {}),
    ...deco,
  };
};

const parseScatterChart = (scatterEl: XmlNode): ScatterChart => {
  const scatterStyle = (valAttr(findChild(scatterEl, SCATTER_STYLE_TAG)) ?? 'lineMarker') as ScatterStyle;
  const varyColors = boolVal(findChild(scatterEl, VARY_COLORS_TAG));
  const series: ScatterSeries[] = [];
  for (const ser of findChildren(scatterEl, SER_TAG)) {
    const s = parseScatterSeries(ser);
    if (s) series.push(s);
  }
  return makeScatterChart({
    scatterStyle,
    series,
    axIds: parseAxIds(scatterEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
  });
};

const parseRadarChart = (radarEl: XmlNode): RadarChart => {
  const radarStyle = (valAttr(findChild(radarEl, RADAR_STYLE_TAG)) ?? 'standard') as RadarStyle;
  const varyColors = boolVal(findChild(radarEl, VARY_COLORS_TAG));
  return makeRadarChart({
    radarStyle,
    series: parseBarSeriesList(radarEl),
    axIds: parseAxIds(radarEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
  });
};

const parseBubbleSeries = (serEl: XmlNode): BubbleSeries | undefined => {
  const idx = intVal(findChild(serEl, IDX_TAG));
  const order = intVal(findChild(serEl, ORDER_TAG));
  if (idx === undefined) return undefined;
  const yVal = parseNumericRef(serEl, Y_VAL_TAG);
  const bubbleSize = parseNumericRef(serEl, BUBBLE_SIZE_TAG);
  if (!yVal || !bubbleSize) return undefined;
  const xVal = parseNumericRef(serEl, X_VAL_TAG);
  const bubble3D = boolVal(findChild(serEl, BUBBLE_3D_TAG));
  const opts: Parameters<typeof makeBubbleSeries>[0] = { idx, yVal, bubbleSize };
  if (order !== undefined) opts.order = order;
  if (xVal) opts.xVal = xVal;
  if (bubble3D !== undefined) opts.bubble3D = bubble3D;
  const tx = parseSeriesTx(serEl);
  if (tx) opts.tx = tx;
  const base = makeBubbleSeries(opts);
  const spPr = parseSpPrSlot(serEl);
  const invertIfNegative = boolVal(findChild(serEl, INVERT_IF_NEGATIVE_TAG));
  const dPt = parseDataPointList(serEl);
  const deco = parseSeriesDecorations(serEl);
  return {
    ...base,
    ...(spPr ? { spPr } : {}),
    ...(invertIfNegative !== undefined ? { invertIfNegative } : {}),
    ...(dPt ? { dPt } : {}),
    ...deco,
  };
};

const parseBubbleChart = (bubbleEl: XmlNode): BubbleChart => {
  const varyColors = boolVal(findChild(bubbleEl, VARY_COLORS_TAG));
  const bubble3D = boolVal(findChild(bubbleEl, BUBBLE_3D_TAG));
  const bubbleScale = intVal(findChild(bubbleEl, BUBBLE_SCALE_TAG));
  const showNegBubbles = boolVal(findChild(bubbleEl, SHOW_NEG_BUBBLES_TAG));
  const sizeRepresentsRaw = valAttr(findChild(bubbleEl, SIZE_REPRESENTS_TAG));
  const sizeRepresents: BubbleSizeRepresents | undefined =
    sizeRepresentsRaw === 'area' || sizeRepresentsRaw === 'w' ? sizeRepresentsRaw : undefined;
  const series: BubbleSeries[] = [];
  for (const ser of findChildren(bubbleEl, SER_TAG)) {
    const s = parseBubbleSeries(ser);
    if (s) series.push(s);
  }
  return makeBubbleChart({
    series,
    axIds: parseAxIds(bubbleEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(bubble3D !== undefined ? { bubble3D } : {}),
    ...(bubbleScale !== undefined ? { bubbleScale } : {}),
    ...(showNegBubbles !== undefined ? { showNegBubbles } : {}),
    ...(sizeRepresents !== undefined ? { sizeRepresents } : {}),
  });
};

const UP_BARS_TAG = `{${CHART_NS}}upBars`;
const DOWN_BARS_TAG = `{${CHART_NS}}downBars`;

const parseHiLowLines = (el: XmlNode): boolean | HiLowLines => {
  const spPr = parseSpPrSlot(el);
  return spPr ? { spPr } : true;
};

const parseUpDownBars = (el: XmlNode): boolean | UpDownBars => {
  const gapWidth = intVal(findChild(el, GAP_WIDTH_TAG));
  const upBarsEl = findChild(el, UP_BARS_TAG);
  const downBarsEl = findChild(el, DOWN_BARS_TAG);
  const upBarsSpPr = upBarsEl ? parseSpPrSlot(upBarsEl) : undefined;
  const downBarsSpPr = downBarsEl ? parseSpPrSlot(downBarsEl) : undefined;
  const out: UpDownBars = {};
  // Treat the default gapWidth = 150 as "unset" to keep round-trip output minimal.
  if (gapWidth !== undefined && gapWidth !== 150) out.gapWidth = gapWidth;
  if (upBarsEl) out.upBars = upBarsSpPr ? { spPr: upBarsSpPr } : {};
  if (downBarsEl) out.downBars = downBarsSpPr ? { spPr: downBarsSpPr } : {};
  return Object.keys(out).length > 0 ? out : true;
};

const parseStockChart = (stockEl: XmlNode): StockChart => {
  const hiLowLinesEl = findChild(stockEl, HI_LOW_LINES_TAG);
  const hiLowLines = hiLowLinesEl ? parseHiLowLines(hiLowLinesEl) : undefined;
  const upDownBarsEl = findChild(stockEl, UP_DOWN_BARS_TAG);
  const upDownBars = upDownBarsEl ? parseUpDownBars(upDownBarsEl) : undefined;
  return makeStockChart({
    series: parseBarSeriesList(stockEl),
    axIds: parseAxIds(stockEl),
    ...(hiLowLines !== undefined ? { hiLowLines } : {}),
    ...(upDownBars !== undefined ? { upDownBars } : {}),
  });
};

const parseSurfaceChart = (surfaceEl: XmlNode): SurfaceChart => {
  const wireframe = boolVal(findChild(surfaceEl, WIREFRAME_TAG));
  const axIds3 = parseAxIds3(surfaceEl);
  return makeSurfaceChart({
    series: parseBarSeriesList(surfaceEl),
    axIds: axIds3,
    ...(wireframe !== undefined ? { wireframe } : {}),
  });
};

const parseAxIds3 = (chartEl: XmlNode): [number, number, number] => {
  const axIdNodes = findChildren(chartEl, AX_ID_TAG);
  return [intVal(axIdNodes[0]) ?? 1, intVal(axIdNodes[1]) ?? 2, intVal(axIdNodes[2]) ?? 3];
};

const SPLIT_TYPES: ReadonlyArray<SplitType> = ['auto', 'cust', 'percent', 'pos', 'val'];

const parseOfPieChart = (ofPieEl: XmlNode): OfPieChart => {
  const ofPieType = (valAttr(findChild(ofPieEl, OF_PIE_TYPE_TAG)) ?? 'pie') as OfPieType;
  const varyColors = boolVal(findChild(ofPieEl, VARY_COLORS_TAG));
  const gapWidth = intVal(findChild(ofPieEl, GAP_WIDTH_TAG));
  const splitTypeRaw = valAttr(findChild(ofPieEl, SPLIT_TYPE_TAG));
  const splitType =
    splitTypeRaw && (SPLIT_TYPES as ReadonlyArray<string>).includes(splitTypeRaw)
      ? (splitTypeRaw as SplitType)
      : undefined;
  const splitPos = intVal(findChild(ofPieEl, SPLIT_POS_TAG));
  const secondPieSize = intVal(findChild(ofPieEl, SECOND_PIE_SIZE_TAG));
  // <c:custSplit> contains <c:secondaryPt idx=N/> entries.
  let custSplit: number[] | undefined;
  const custSplitEl = findChild(ofPieEl, CUST_SPLIT_TAG);
  if (custSplitEl) {
    const points: number[] = [];
    for (const pt of findChildren(custSplitEl, SEC_BLOCK_PT_TAG)) {
      const idx = Number.parseInt(pt.attrs['idx'] ?? '', 10);
      if (Number.isInteger(idx)) points.push(idx);
    }
    custSplit = points;
  }
  return makeOfPieChart({
    ofPieType,
    series: parseBarSeriesList(ofPieEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapWidth !== undefined ? { gapWidth } : {}),
    ...(splitType !== undefined ? { splitType } : {}),
    ...(splitPos !== undefined ? { splitPos } : {}),
    ...(custSplit !== undefined ? { custSplit } : {}),
    ...(secondPieSize !== undefined ? { secondPieSize } : {}),
  });
};

// 3-D parsers — reuse 2-D helpers where the shape matches.

const parseBar3DChart = (barEl: XmlNode): Bar3DChart => {
  const barDir = (valAttr(findChild(barEl, BAR_DIR_TAG)) ?? 'col') as BarDirection;
  const grouping = (valAttr(findChild(barEl, GROUPING_TAG)) ?? 'clustered') as GroupingType;
  const varyColors = boolVal(findChild(barEl, VARY_COLORS_TAG));
  const gapWidth = intVal(findChild(barEl, GAP_WIDTH_TAG));
  const gapDepth = intVal(findChild(barEl, GAP_DEPTH_TAG));
  const shape = valAttr(findChild(barEl, SHAPE_TAG)) as Bar3DChart['shape'] | undefined;
  return makeBar3DChart({
    barDir,
    grouping,
    series: parseBarSeriesList(barEl),
    axIds: parseAxIds3(barEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapWidth !== undefined ? { gapWidth } : {}),
    ...(gapDepth !== undefined ? { gapDepth } : {}),
    ...(shape !== undefined ? { shape } : {}),
  });
};

const parseLine3DChart = (lineEl: XmlNode): Line3DChart => {
  const grouping = (valAttr(findChild(lineEl, GROUPING_TAG)) ?? 'standard') as GroupingType;
  const varyColors = boolVal(findChild(lineEl, VARY_COLORS_TAG));
  const gapDepth = intVal(findChild(lineEl, GAP_DEPTH_TAG));
  const series: LineSeries[] = [];
  for (const ser of findChildren(lineEl, SER_TAG)) {
    const s = parseLineSeries(ser);
    if (s) series.push(s);
  }
  return makeLine3DChart({
    grouping,
    series,
    axIds: parseAxIds3(lineEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapDepth !== undefined ? { gapDepth } : {}),
  });
};

const parsePie3DChart = (pieEl: XmlNode): Pie3DChart => {
  const varyColors = boolVal(findChild(pieEl, VARY_COLORS_TAG));
  return makePie3DChart({
    series: parseBarSeriesList(pieEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
  });
};

const parseArea3DChart = (areaEl: XmlNode): Area3DChart => {
  const grouping = (valAttr(findChild(areaEl, GROUPING_TAG)) ?? 'standard') as GroupingType;
  const varyColors = boolVal(findChild(areaEl, VARY_COLORS_TAG));
  const gapDepth = intVal(findChild(areaEl, GAP_DEPTH_TAG));
  return makeArea3DChart({
    grouping,
    series: parseBarSeriesList(areaEl),
    axIds: parseAxIds3(areaEl),
    ...(varyColors !== undefined ? { varyColors } : {}),
    ...(gapDepth !== undefined ? { gapDepth } : {}),
  });
};

const parseSurface3DChart = (surfaceEl: XmlNode): Surface3DChart => {
  const wireframe = boolVal(findChild(surfaceEl, WIREFRAME_TAG));
  return makeSurface3DChart({
    series: parseBarSeriesList(surfaceEl),
    axIds: parseAxIds3(surfaceEl),
    ...(wireframe !== undefined ? { wireframe } : {}),
  });
};

const parsePlotChart = (plotAreaEl: XmlNode): ChartKind => {
  const bar = findChild(plotAreaEl, BAR_CHART_TAG);
  if (bar) return parseBarChart(bar);
  const line = findChild(plotAreaEl, LINE_CHART_TAG);
  if (line) return parseLineChart(line);
  const area = findChild(plotAreaEl, AREA_CHART_TAG);
  if (area) return parseAreaChart(area);
  const pie = findChild(plotAreaEl, PIE_CHART_TAG);
  if (pie) return parsePieChart(pie);
  const donut = findChild(plotAreaEl, DOUGHNUT_CHART_TAG);
  if (donut) return parseDoughnutChart(donut);
  const scatter = findChild(plotAreaEl, SCATTER_CHART_TAG);
  if (scatter) return parseScatterChart(scatter);
  const radar = findChild(plotAreaEl, RADAR_CHART_TAG);
  if (radar) return parseRadarChart(radar);
  const bubble = findChild(plotAreaEl, BUBBLE_CHART_TAG);
  if (bubble) return parseBubbleChart(bubble);
  const stock = findChild(plotAreaEl, STOCK_CHART_TAG);
  if (stock) return parseStockChart(stock);
  const surface = findChild(plotAreaEl, SURFACE_CHART_TAG);
  if (surface) return parseSurfaceChart(surface);
  const ofPie = findChild(plotAreaEl, OF_PIE_CHART_TAG);
  if (ofPie) return parseOfPieChart(ofPie);
  const bar3D = findChild(plotAreaEl, BAR3D_CHART_TAG);
  if (bar3D) return parseBar3DChart(bar3D);
  const line3D = findChild(plotAreaEl, LINE3D_CHART_TAG);
  if (line3D) return parseLine3DChart(line3D);
  const pie3D = findChild(plotAreaEl, PIE3D_CHART_TAG);
  if (pie3D) return parsePie3DChart(pie3D);
  const area3D = findChild(plotAreaEl, AREA3D_CHART_TAG);
  if (area3D) return parseArea3DChart(area3D);
  const surface3D = findChild(plotAreaEl, SURFACE3D_CHART_TAG);
  if (surface3D) return parseSurface3DChart(surface3D);
  throw new OpenXmlSchemaError('parseChartXml: no supported chart kind found inside <plotArea>');
};

const VALID_LAYOUT_MODES: ReadonlyArray<string> = ['edge', 'factor'];
const VALID_LAYOUT_TARGETS: ReadonlyArray<string> = ['inner', 'outer'];

const parseLayout = (el: XmlNode): Layout | undefined => {
  const manualEl = findChild(el, MANUAL_LAYOUT_TAG);
  if (!manualEl) return undefined;
  const targetRaw = valAttr(findChild(manualEl, LAYOUT_TARGET_TAG));
  const layoutTarget = targetRaw && VALID_LAYOUT_TARGETS.includes(targetRaw)
    ? (targetRaw as LayoutTarget)
    : undefined;
  const xMode = valAttr(findChild(manualEl, X_MODE_TAG));
  const yMode = valAttr(findChild(manualEl, Y_MODE_TAG));
  const wMode = valAttr(findChild(manualEl, W_MODE_TAG));
  const hMode = valAttr(findChild(manualEl, H_MODE_TAG));
  const x = numberVal(findChild(manualEl, X_TAG));
  const y = numberVal(findChild(manualEl, Y_TAG));
  const w = numberVal(findChild(manualEl, W_TAG));
  const h = numberVal(findChild(manualEl, H_TAG));
  const ml: ManualLayout = {};
  if (layoutTarget) ml.layoutTarget = layoutTarget;
  if (xMode && VALID_LAYOUT_MODES.includes(xMode)) ml.xMode = xMode as LayoutMode;
  if (yMode && VALID_LAYOUT_MODES.includes(yMode)) ml.yMode = yMode as LayoutMode;
  if (wMode && VALID_LAYOUT_MODES.includes(wMode)) ml.wMode = wMode as LayoutMode;
  if (hMode && VALID_LAYOUT_MODES.includes(hMode)) ml.hMode = hMode as LayoutMode;
  if (x !== undefined) ml.x = x;
  if (y !== undefined) ml.y = y;
  if (w !== undefined) ml.w = w;
  if (h !== undefined) ml.h = h;
  return Object.keys(ml).length > 0 ? { manualLayout: ml } : undefined;
};

const serializeLayout = (layout: Layout | undefined): string => {
  const ml = layout?.manualLayout;
  if (!ml || Object.keys(ml).length === 0) return '<c:layout/>';
  const parts: string[] = ['<c:layout><c:manualLayout>'];
  if (ml.layoutTarget) parts.push(`<c:layoutTarget val="${ml.layoutTarget}"/>`);
  if (ml.xMode) parts.push(`<c:xMode val="${ml.xMode}"/>`);
  if (ml.yMode) parts.push(`<c:yMode val="${ml.yMode}"/>`);
  if (ml.wMode) parts.push(`<c:wMode val="${ml.wMode}"/>`);
  if (ml.hMode) parts.push(`<c:hMode val="${ml.hMode}"/>`);
  if (ml.x !== undefined) parts.push(`<c:x val="${ml.x}"/>`);
  if (ml.y !== undefined) parts.push(`<c:y val="${ml.y}"/>`);
  if (ml.w !== undefined) parts.push(`<c:w val="${ml.w}"/>`);
  if (ml.h !== undefined) parts.push(`<c:h val="${ml.h}"/>`);
  parts.push('</c:manualLayout></c:layout>');
  return parts.join('');
};

const parseView3D = (el: XmlNode): View3D | undefined => {
  const rotX = intVal(findChild(el, ROT_X_TAG));
  const rotY = intVal(findChild(el, ROT_Y_TAG));
  const hPercent = intVal(findChild(el, H_PERCENT_TAG));
  const depthPercent = intVal(findChild(el, DEPTH_PERCENT_TAG));
  const rAngAx = boolVal(findChild(el, R_ANG_AX_TAG));
  const perspective = intVal(findChild(el, PERSPECTIVE_TAG));
  const out: View3D = {};
  if (rotX !== undefined) out.rotX = rotX;
  if (rotY !== undefined) out.rotY = rotY;
  if (depthPercent !== undefined) out.depthPercent = depthPercent;
  if (hPercent !== undefined) out.hPercent = hPercent;
  if (rAngAx !== undefined) out.rAngAx = rAngAx;
  if (perspective !== undefined) out.perspective = perspective;
  return Object.keys(out).length > 0 ? out : undefined;
};

const serializeView3D = (v: View3D): string => {
  // ECMA-376 order: rotX, hPercent, rotY, depthPercent, rAngAx, perspective.
  const parts: string[] = ['<c:view3D>'];
  if (v.rotX !== undefined) parts.push(`<c:rotX val="${v.rotX}"/>`);
  if (v.hPercent !== undefined) parts.push(`<c:hPercent val="${v.hPercent}"/>`);
  if (v.rotY !== undefined) parts.push(`<c:rotY val="${v.rotY}"/>`);
  if (v.depthPercent !== undefined) parts.push(`<c:depthPercent val="${v.depthPercent}"/>`);
  if (v.rAngAx !== undefined) parts.push(`<c:rAngAx val="${v.rAngAx ? '1' : '0'}"/>`);
  if (v.perspective !== undefined) parts.push(`<c:perspective val="${v.perspective}"/>`);
  parts.push('</c:view3D>');
  return parts.join('');
};

const parseSurfaceFrame = (el: XmlNode): SurfaceFrame | undefined => {
  const thickness = intVal(findChild(el, THICKNESS_TAG));
  const spPr = parseSpPrSlot(el);
  const out: SurfaceFrame = {};
  if (thickness !== undefined) out.thickness = thickness;
  if (spPr) out.spPr = spPr;
  return Object.keys(out).length > 0 ? out : undefined;
};

const serializeSurfaceFrame = (tag: 'floor' | 'sideWall' | 'backWall', f: SurfaceFrame): string => {
  const parts: string[] = [`<c:${tag}>`];
  if (f.thickness !== undefined) parts.push(`<c:thickness val="${f.thickness}"/>`);
  if (f.spPr) parts.push(serializeShapeProperties(f.spPr));
  parts.push(`</c:${tag}>`);
  return parts.join('');
};

const VALID_TICK_MARKS: ReadonlyArray<string> = ['cross', 'in', 'none', 'out'];
const VALID_TICK_LBL_POS: ReadonlyArray<string> = ['high', 'low', 'nextTo', 'none'];
const VALID_AXIS_CROSSES: ReadonlyArray<string> = ['autoZero', 'max', 'min'];
const VALID_ORIENTATIONS: ReadonlyArray<string> = ['maxMin', 'minMax'];
const VALID_CROSS_BETWEEN: ReadonlyArray<string> = ['between', 'midCat'];
const VALID_CAT_LABEL_ALIGN: ReadonlyArray<string> = ['ctr', 'l', 'r'];

const numberVal = (el: XmlNode | undefined): number | undefined => {
  if (!el) return undefined;
  const raw = valAttr(el);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
};

const parseScaling = (axEl: XmlNode): AxisScaling | undefined => {
  const sc = findChild(axEl, SCALING_TAG);
  if (!sc) return undefined;
  const orientationRaw = valAttr(findChild(sc, ORIENTATION_TAG));
  // `minMax` is the implicit ECMA-376 default and the value the serializer
  // emits when scaling is unset, so treat it as "no user override".
  const orientation = orientationRaw && VALID_ORIENTATIONS.includes(orientationRaw) && orientationRaw !== 'minMax'
    ? (orientationRaw as AxisOrientation)
    : undefined;
  const min = numberVal(findChild(sc, MIN_TAG));
  const max = numberVal(findChild(sc, MAX_TAG));
  const logBase = numberVal(findChild(sc, LOG_BASE_TAG));
  const out: AxisScaling = {};
  if (orientation !== undefined) out.orientation = orientation;
  if (min !== undefined) out.min = min;
  if (max !== undefined) out.max = max;
  if (logBase !== undefined) out.logBase = logBase;
  return Object.keys(out).length > 0 ? out : undefined;
};

const serializeScaling = (s: AxisScaling | undefined): string => {
  const parts: string[] = ['<c:scaling>'];
  parts.push(`<c:orientation val="${s?.orientation ?? 'minMax'}"/>`);
  if (s?.logBase !== undefined) parts.push(`<c:logBase val="${s.logBase}"/>`);
  if (s?.max !== undefined) parts.push(`<c:max val="${s.max}"/>`);
  if (s?.min !== undefined) parts.push(`<c:min val="${s.min}"/>`);
  parts.push('</c:scaling>');
  return parts.join('');
};

const parseAxisCommon = (
  axEl: XmlNode,
): {
  axId: number;
  crossAx: number;
  position?: 'b' | 't' | 'l' | 'r';
  delete?: boolean;
  majorGridlines?: boolean | Gridlines;
  minorGridlines?: boolean | Gridlines;
  spPr?: ShapeProperties;
  txPr?: TextBody;
  title?: ChartTitle;
  numFmt?: ChartNumberFormat;
  majorTickMark?: TickMark;
  minorTickMark?: TickMark;
  tickLblPos?: TickLabelPosition;
  scaling?: AxisScaling;
  crosses?: AxisCrosses;
  crossesAt?: number;
} => {
  const axId = intVal(findChild(axEl, AX_ID_TAG)) ?? 0;
  const crossAx = intVal(findChild(axEl, CROSS_AX_TAG)) ?? 0;
  const positionRaw = valAttr(findChild(axEl, AX_POS_TAG));
  const validPos = positionRaw === 'b' || positionRaw === 't' || positionRaw === 'l' || positionRaw === 'r';
  const del = boolVal(findChild(axEl, DELETE_TAG));
  const majorGridlines = parseGridlines(axEl, MAJOR_GRIDLINES_TAG);
  const minorGridlines = parseGridlines(axEl, MINOR_GRIDLINES_TAG);
  const spPr = parseSpPrSlot(axEl);
  const txPr = parseTxPrSlot(axEl);
  const titleEl = findChild(axEl, TITLE_TAG);
  const title = titleEl ? parseChartTitle(titleEl) : undefined;
  const numFmt = parseNumberFormat(findChild(axEl, NUM_FMT_TAG));
  // Excel inserts `<c:numFmt formatCode="General" sourceLinked="1"/>` as a default,
  // so we treat that exact pair as "no user-provided override" to keep round-trip
  // output minimal for callers who didn't set numFmt.
  const userNumFmt = numFmt && (numFmt.formatCode !== 'General' || numFmt.sourceLinked !== true) ? numFmt : undefined;
  const majorRaw = valAttr(findChild(axEl, MAJOR_TICK_MARK_TAG));
  const majorTickMark = majorRaw && VALID_TICK_MARKS.includes(majorRaw) && majorRaw !== 'out'
    ? (majorRaw as TickMark)
    : undefined;
  const minorRaw = valAttr(findChild(axEl, MINOR_TICK_MARK_TAG));
  const minorTickMark = minorRaw && VALID_TICK_MARKS.includes(minorRaw) && minorRaw !== 'none'
    ? (minorRaw as TickMark)
    : undefined;
  const tickLblPosRaw = valAttr(findChild(axEl, TICK_LBL_POS_TAG));
  const tickLblPos = tickLblPosRaw && VALID_TICK_LBL_POS.includes(tickLblPosRaw) && tickLblPosRaw !== 'nextTo'
    ? (tickLblPosRaw as TickLabelPosition)
    : undefined;
  const crossesRaw = valAttr(findChild(axEl, CROSSES_TAG));
  const crosses = crossesRaw && VALID_AXIS_CROSSES.includes(crossesRaw) && crossesRaw !== 'autoZero'
    ? (crossesRaw as AxisCrosses)
    : undefined;
  const crossesAt = numberVal(findChild(axEl, CROSSES_AT_TAG));
  const scaling = parseScaling(axEl);
  return {
    axId,
    crossAx,
    ...(validPos ? { position: positionRaw as 'b' | 't' | 'l' | 'r' } : {}),
    ...(del !== undefined ? { delete: del } : {}),
    ...(majorGridlines !== undefined ? { majorGridlines } : {}),
    ...(minorGridlines !== undefined ? { minorGridlines } : {}),
    ...(spPr ? { spPr } : {}),
    ...(txPr ? { txPr } : {}),
    ...(title ? { title } : {}),
    ...(userNumFmt ? { numFmt: userNumFmt } : {}),
    ...(majorTickMark ? { majorTickMark } : {}),
    ...(minorTickMark ? { minorTickMark } : {}),
    ...(tickLblPos ? { tickLblPos } : {}),
    ...(scaling ? { scaling } : {}),
    ...(crosses ? { crosses } : {}),
    ...(crossesAt !== undefined ? { crossesAt } : {}),
  };
};

const parseCategoryAxis = (axEl: XmlNode): CategoryAxis => {
  const base = parseAxisCommon(axEl);
  const auto = boolVal(findChild(axEl, AUTO_TAG));
  // <c:auto val="1"/> is the default Excel emits, so treat the literal "1" as
  // "no user override" to keep round-trip output minimal.
  const userAuto = auto === false ? false : undefined;
  const lblAlgnRaw = valAttr(findChild(axEl, LBL_ALGN_TAG));
  const lblAlgn = lblAlgnRaw && VALID_CAT_LABEL_ALIGN.includes(lblAlgnRaw) && lblAlgnRaw !== 'ctr'
    ? (lblAlgnRaw as CategoryLabelAlignment)
    : undefined;
  const lblOffsetRaw = intVal(findChild(axEl, LBL_OFFSET_TAG));
  const lblOffset = lblOffsetRaw !== undefined && lblOffsetRaw !== 100 ? lblOffsetRaw : undefined;
  const noMultiLvlLbl = boolVal(findChild(axEl, NO_MULTI_LVL_LBL_TAG));
  const userNoMultiLvlLbl = noMultiLvlLbl === true ? true : undefined;
  return {
    ...base,
    ...(userAuto !== undefined ? { auto: userAuto } : {}),
    ...(lblAlgn ? { lblAlgn } : {}),
    ...(lblOffset !== undefined ? { lblOffset } : {}),
    ...(userNoMultiLvlLbl !== undefined ? { noMultiLvlLbl: userNoMultiLvlLbl } : {}),
  };
};

const VALID_TIME_UNITS: ReadonlyArray<string> = ['days', 'months', 'years'];

const BASE_TIME_UNIT_TAG = `{${CHART_NS}}baseTimeUnit`;
const MAJOR_TIME_UNIT_TAG = `{${CHART_NS}}majorTimeUnit`;
const MINOR_TIME_UNIT_TAG = `{${CHART_NS}}minorTimeUnit`;
const TICK_LBL_SKIP_TAG = `{${CHART_NS}}tickLblSkip`;
const TICK_MARK_SKIP_TAG = `{${CHART_NS}}tickMarkSkip`;

const parseDateAxis = (axEl: XmlNode): DateAxis => {
  const base = parseAxisCommon(axEl);
  const auto = boolVal(findChild(axEl, AUTO_TAG));
  const userAuto = auto === false ? false : undefined;
  const lblOffsetRaw = intVal(findChild(axEl, LBL_OFFSET_TAG));
  const lblOffset = lblOffsetRaw !== undefined && lblOffsetRaw !== 100 ? lblOffsetRaw : undefined;
  const baseTimeUnitRaw = valAttr(findChild(axEl, BASE_TIME_UNIT_TAG));
  const baseTimeUnit = baseTimeUnitRaw && VALID_TIME_UNITS.includes(baseTimeUnitRaw)
    ? (baseTimeUnitRaw as TimeUnit)
    : undefined;
  const majorUnit = numberVal(findChild(axEl, MAJOR_UNIT_TAG));
  const majorTimeUnitRaw = valAttr(findChild(axEl, MAJOR_TIME_UNIT_TAG));
  const majorTimeUnit = majorTimeUnitRaw && VALID_TIME_UNITS.includes(majorTimeUnitRaw)
    ? (majorTimeUnitRaw as TimeUnit)
    : undefined;
  const minorUnit = numberVal(findChild(axEl, MINOR_UNIT_TAG));
  const minorTimeUnitRaw = valAttr(findChild(axEl, MINOR_TIME_UNIT_TAG));
  const minorTimeUnit = minorTimeUnitRaw && VALID_TIME_UNITS.includes(minorTimeUnitRaw)
    ? (minorTimeUnitRaw as TimeUnit)
    : undefined;
  return {
    ...base,
    ...(userAuto !== undefined ? { auto: userAuto } : {}),
    ...(lblOffset !== undefined ? { lblOffset } : {}),
    ...(baseTimeUnit ? { baseTimeUnit } : {}),
    ...(majorUnit !== undefined ? { majorUnit } : {}),
    ...(majorTimeUnit ? { majorTimeUnit } : {}),
    ...(minorUnit !== undefined ? { minorUnit } : {}),
    ...(minorTimeUnit ? { minorTimeUnit } : {}),
  };
};

const parseSeriesAxis = (axEl: XmlNode): SeriesAxis => {
  const base = parseAxisCommon(axEl);
  const tickLblSkip = intVal(findChild(axEl, TICK_LBL_SKIP_TAG));
  const tickMarkSkip = intVal(findChild(axEl, TICK_MARK_SKIP_TAG));
  return {
    ...base,
    ...(tickLblSkip !== undefined ? { tickLblSkip } : {}),
    ...(tickMarkSkip !== undefined ? { tickMarkSkip } : {}),
  };
};

const parseValueAxis = (axEl: XmlNode): ValueAxis => {
  const base = parseAxisCommon(axEl);
  const crossBetweenRaw = valAttr(findChild(axEl, CROSS_BETWEEN_TAG));
  const crossBetween = crossBetweenRaw && VALID_CROSS_BETWEEN.includes(crossBetweenRaw) && crossBetweenRaw !== 'between'
    ? (crossBetweenRaw as AxisCrossBetween)
    : undefined;
  const majorUnit = numberVal(findChild(axEl, MAJOR_UNIT_TAG));
  const minorUnit = numberVal(findChild(axEl, MINOR_UNIT_TAG));
  return {
    ...base,
    ...(crossBetween ? { crossBetween } : {}),
    ...(majorUnit !== undefined ? { majorUnit } : {}),
    ...(minorUnit !== undefined ? { minorUnit } : {}),
  };
};

/** Parse a `xl/charts/chartN.xml` payload. */
export function parseChartXml(bytes: Uint8Array | string): ChartSpace {
  const root = parseXml(bytes);
  if (root.name !== CHART_SPACE_TAG) {
    throw new OpenXmlSchemaError(`parseChartXml: root is "${root.name}", expected chartSpace`);
  }
  const chartEl = findChild(root, CHART_TAG);
  if (!chartEl) throw new OpenXmlSchemaError('parseChartXml: <chartSpace> missing <chart>');
  const plotAreaEl = findChild(chartEl, PLOT_AREA_TAG);
  if (!plotAreaEl) throw new OpenXmlSchemaError('parseChartXml: <chart> missing <plotArea>');
  const chart = parsePlotChart(plotAreaEl);
  const catAxEl = findChild(plotAreaEl, CAT_AX_TAG);
  // Scatter and bubble charts have two value axes; the X one is first in the chart's axId list.
  const valAxEls = findChildren(plotAreaEl, VAL_AX_TAG);
  const xAxId = 'axIds' in chart && (chart.kind === 'scatter' || chart.kind === 'bubble') ? chart.axIds?.[0] : undefined;
  const xValAxEl = xAxId === undefined || valAxEls.length < 2 ? undefined : valAxEls.find((el) => intVal(findChild(el, AX_ID_TAG)) === xAxId);
  const valAxEl = valAxEls.find((el) => el !== xValAxEl);
  const plotAreaSpPr = parseSpPrSlot(plotAreaEl);
  const plotArea: PlotArea = {
    chart,
    ...(catAxEl ? { catAx: parseCategoryAxis(catAxEl) } : {}),
    ...(valAxEl ? { valAx: parseValueAxis(valAxEl) } : {}),
    ...(xValAxEl ? { xValAx: parseValueAxis(xValAxEl) } : {}),
    ...((): { dateAx?: DateAxis; serAx?: SeriesAxis; layout?: Layout } => {
      const out: { dateAx?: DateAxis; serAx?: SeriesAxis; layout?: Layout } = {};
      const dateAxEl = findChild(plotAreaEl, DATE_AX_TAG);
      if (dateAxEl) out.dateAx = parseDateAxis(dateAxEl);
      const serAxEl = findChild(plotAreaEl, SER_AX_TAG);
      if (serAxEl) out.serAx = parseSeriesAxis(serAxEl);
      const layoutEl = findChild(plotAreaEl, LAYOUT_TAG);
      if (layoutEl) {
        const layout = parseLayout(layoutEl);
        if (layout) out.layout = layout;
      }
      return out;
    })(),
    ...(plotAreaSpPr ? { spPr: plotAreaSpPr } : {}),
  };
  const titleEl = findChild(chartEl, TITLE_TAG);
  const title = titleEl ? parseChartTitle(titleEl) : undefined;
  const legendEl = findChild(chartEl, LEGEND_TAG);
  let legend: Legend | undefined;
  if (legendEl) {
    const posRaw = valAttr(findChild(legendEl, LEGEND_POS_TAG)) as LegendPosition | undefined;
    const overlay = boolVal(findChild(legendEl, OVERLAY_TAG));
    const legendSpPr = parseSpPrSlot(legendEl);
    const legendTxPr = parseTxPrSlot(legendEl);
    const legendLayoutEl = findChild(legendEl, LAYOUT_TAG);
    const legendLayout = legendLayoutEl ? parseLayout(legendLayoutEl) : undefined;
    legend = {
      position: posRaw ?? 'r',
      ...(overlay !== undefined ? { overlay } : {}),
      ...(legendLayout ? { layout: legendLayout } : {}),
      ...(legendSpPr ? { spPr: legendSpPr } : {}),
      ...(legendTxPr ? { txPr: legendTxPr } : {}),
    };
  }
  const plotVisOnly = boolVal(findChild(chartEl, PLOT_VIS_ONLY_TAG));
  const dispBlanksAs = valAttr(findChild(chartEl, DISP_BLANKS_AS_TAG)) as ChartSpace['dispBlanksAs'];
  const style = intVal(findChild(root, STYLE_TAG));
  const view3DEl = findChild(chartEl, VIEW3D_TAG);
  const view3D = view3DEl ? parseView3D(view3DEl) : undefined;
  const floorEl = findChild(chartEl, FLOOR_TAG);
  const floor = floorEl ? parseSurfaceFrame(floorEl) : undefined;
  const sideWallEl = findChild(chartEl, SIDE_WALL_TAG);
  const sideWall = sideWallEl ? parseSurfaceFrame(sideWallEl) : undefined;
  const backWallEl = findChild(chartEl, BACK_WALL_TAG);
  const backWall = backWallEl ? parseSurfaceFrame(backWallEl) : undefined;
  // Top-level spPr / txPr live on chartSpace (sibling of <c:chart>), not inside
  // <c:chart>.
  const spaceSpPr = parseSpPrSlot(root);
  const spaceTxPr = parseTxPrSlot(root);
  return makeChartSpace({
    plotArea,
    ...(title !== undefined ? { title } : {}),
    ...(legend ? { legend } : {}),
    ...(style !== undefined ? { style } : {}),
    ...(view3D ? { view3D } : {}),
    ...(floor ? { floor } : {}),
    ...(sideWall ? { sideWall } : {}),
    ...(backWall ? { backWall } : {}),
    ...(plotVisOnly !== undefined ? { plotVisOnly } : {}),
    ...(dispBlanksAs ? { dispBlanksAs } : {}),
    ...(spaceSpPr ? { spPr: spaceSpPr } : {}),
    ...(spaceTxPr ? { txPr: spaceTxPr } : {}),
  });
}

const serializeNumCache = (cache: ReadonlyArray<number>, formatCode?: string): string => {
  const parts: string[] = ['<c:numCache>'];
  if (formatCode) parts.push(`<c:formatCode>${escapeText(formatCode)}</c:formatCode>`);
  parts.push(`<c:ptCount val="${cache.length}"/>`);
  for (let i = 0; i < cache.length; i++) {
    const v = cache[i] as number;
    parts.push(`<c:pt idx="${i}"><c:v>${escapeText(String(v))}</c:v></c:pt>`);
  }
  parts.push('</c:numCache>');
  return parts.join('');
};

const serializeStrCache = (cache: ReadonlyArray<string | number>): string => {
  const parts: string[] = ['<c:strCache>', `<c:ptCount val="${cache.length}"/>`];
  for (let i = 0; i < cache.length; i++) {
    const v = cache[i] as string | number;
    parts.push(`<c:pt idx="${i}"><c:v>${escapeText(String(v))}</c:v></c:pt>`);
  }
  parts.push('</c:strCache>');
  return parts.join('');
};

const serializeNumericRef = (tag: string, ref: NumericRef): string => {
  const inner = ref.cache !== undefined ? serializeNumCache(ref.cache, ref.formatCode) : '';
  return `<c:${tag}><c:numRef><c:f>${escapeText(ref.ref)}</c:f>${inner}</c:numRef></c:${tag}>`;
};

const serializeCategoryRef = (cat: CategoryRef): string => {
  if (
    cat.cacheKind === 'str' ||
    (cat.cacheKind !== 'num' && cat.cache && cat.cache.some((v) => typeof v === 'string'))
  ) {
    const inner = cat.cache !== undefined ? serializeStrCache(cat.cache) : '';
    return `<c:cat><c:strRef><c:f>${escapeText(cat.ref)}</c:f>${inner}</c:strRef></c:cat>`;
  }
  const numericCache = cat.cache as number[] | undefined;
  const inner = numericCache !== undefined ? serializeNumCache(numericCache, cat.formatCode) : '';
  return `<c:cat><c:numRef><c:f>${escapeText(cat.ref)}</c:f>${inner}</c:numRef></c:cat>`;
};

const serializeSeries = (s: BarSeries, marker?: Marker): string => {
  const parts: string[] = ['<c:ser>', `<c:idx val="${s.idx}"/>`, `<c:order val="${s.order}"/>`];
  if (s.tx) {
    if (s.tx.kind === 'literal') {
      parts.push(`<c:tx><c:v>${escapeText(s.tx.value)}</c:v></c:tx>`);
    } else {
      parts.push(`<c:tx><c:strRef><c:f>${escapeText(s.tx.ref)}</c:f></c:strRef></c:tx>`);
    }
  }
  if (s.spPr) parts.push(serializeShapeProperties(s.spPr));
  // ECMA-376 places per-chart-kind elements after <c:spPr> and before <c:dPt>.
  // Each chart kind only sets the relevant one of marker / invertIfNegative /
  // explosion, so the merged order below produces a schema-valid sequence for
  // every kind that shares this helper (bar / line / area / pie / doughnut /
  // radar / stock / surface).
  if (marker) parts.push(serializeMarker(marker));
  if (s.invertIfNegative !== undefined) {
    parts.push(`<c:invertIfNegative val="${s.invertIfNegative ? '1' : '0'}"/>`);
  }
  if (s.explosion !== undefined) parts.push(`<c:explosion val="${s.explosion}"/>`);
  if (s.dPt) for (const p of s.dPt) parts.push(serializeDataPoint(p));
  if (s.dLbls) parts.push(serializeDataLabelList(s.dLbls));
  if (s.trendline) for (const t of s.trendline) parts.push(serializeTrendline(t));
  if (s.errBars) for (const e of s.errBars) parts.push(serializeErrBars(e));
  if (s.cat) parts.push(serializeCategoryRef(s.cat));
  parts.push(serializeNumericRef('val', s.val));
  parts.push('</c:ser>');
  return parts.join('');
};

const serializeBarChart = (chart: BarChart): string => {
  const parts: string[] = [
    '<c:barChart>',
    `<c:barDir val="${chart.barDir}"/>`,
    `<c:grouping val="${chart.grouping}"/>`,
    `<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`,
  ];
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:gapWidth val="${chart.gapWidth ?? 150}"/>`);
  if (chart.overlap !== undefined) {
    parts.push(`<c:overlap val="${chart.overlap}"/>`);
  } else if (chart.grouping === 'stacked' || chart.grouping === 'percentStacked') {
    parts.push('<c:overlap val="100"/>');
  }
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:barChart>');
  return parts.join('');
};

const serializeLineSeries = (s: LineSeries): string => {
  const base = serializeSeries(s, s.marker);
  if (s.smooth === undefined) return base;
  // Inject smooth before </c:ser> (cheap inline patch).
  return base.replace('</c:ser>', `<c:smooth val="${s.smooth ? '1' : '0'}"/></c:ser>`);
};

const serializeLineChart = (chart: LineChart): string => {
  const parts: string[] = [
    '<c:lineChart>',
    `<c:grouping val="${chart.grouping}"/>`,
    `<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`,
  ];
  for (const s of chart.series) parts.push(serializeLineSeries(s));
  if (chart.smooth !== undefined) parts.push(`<c:smooth val="${chart.smooth ? '1' : '0'}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:lineChart>');
  return parts.join('');
};

const serializeAreaChart = (chart: AreaChart): string => {
  const parts: string[] = [
    '<c:areaChart>',
    `<c:grouping val="${chart.grouping}"/>`,
    `<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`,
  ];
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:areaChart>');
  return parts.join('');
};

const serializePieChart = (chart: PieChart): string => {
  const parts: string[] = ['<c:pieChart>', `<c:varyColors val="${(chart.varyColors ?? true) ? '1' : '0'}"/>`];
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:firstSliceAng val="${chart.firstSliceAng ?? 0}"/>`);
  parts.push('</c:pieChart>');
  return parts.join('');
};

const serializeDoughnutChart = (chart: DoughnutChart): string => {
  const parts: string[] = ['<c:doughnutChart>', `<c:varyColors val="${(chart.varyColors ?? true) ? '1' : '0'}"/>`];
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:firstSliceAng val="${chart.firstSliceAng ?? 0}"/>`);
  parts.push(`<c:holeSize val="${chart.holeSize ?? 50}"/>`);
  parts.push('</c:doughnutChart>');
  return parts.join('');
};

const serializeScatterSeries = (s: ScatterSeries): string => {
  const parts: string[] = ['<c:ser>', `<c:idx val="${s.idx}"/>`, `<c:order val="${s.order}"/>`];
  if (s.tx) {
    if (s.tx.kind === 'literal') {
      parts.push(`<c:tx><c:v>${escapeText(s.tx.value)}</c:v></c:tx>`);
    } else {
      parts.push(`<c:tx><c:strRef><c:f>${escapeText(s.tx.ref)}</c:f></c:strRef></c:tx>`);
    }
  }
  if (s.spPr) parts.push(serializeShapeProperties(s.spPr));
  if (s.marker) parts.push(serializeMarker(s.marker));
  if (s.dPt) for (const p of s.dPt) parts.push(serializeDataPoint(p));
  if (s.dLbls) parts.push(serializeDataLabelList(s.dLbls));
  if (s.trendline) for (const t of s.trendline) parts.push(serializeTrendline(t));
  if (s.errBars) for (const e of s.errBars) parts.push(serializeErrBars(e));
  if (s.xVal) parts.push(serializeNumericRef('xVal', s.xVal));
  parts.push(serializeNumericRef('yVal', s.yVal));
  if (s.smooth !== undefined) parts.push(`<c:smooth val="${s.smooth ? '1' : '0'}"/>`);
  parts.push('</c:ser>');
  return parts.join('');
};

const serializeScatterChart = (chart: ScatterChart): string => {
  const parts: string[] = [
    '<c:scatterChart>',
    `<c:scatterStyle val="${chart.scatterStyle}"/>`,
    `<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`,
  ];
  for (const s of chart.series) parts.push(serializeScatterSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:scatterChart>');
  return parts.join('');
};

const serializeRadarChart = (chart: RadarChart): string => {
  const parts: string[] = [
    '<c:radarChart>',
    `<c:radarStyle val="${chart.radarStyle}"/>`,
    `<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`,
  ];
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:radarChart>');
  return parts.join('');
};

const serializeBubbleSeries = (s: BubbleSeries): string => {
  const parts: string[] = ['<c:ser>', `<c:idx val="${s.idx}"/>`, `<c:order val="${s.order}"/>`];
  if (s.tx) {
    if (s.tx.kind === 'literal') {
      parts.push(`<c:tx><c:v>${escapeText(s.tx.value)}</c:v></c:tx>`);
    } else {
      parts.push(`<c:tx><c:strRef><c:f>${escapeText(s.tx.ref)}</c:f></c:strRef></c:tx>`);
    }
  }
  if (s.spPr) parts.push(serializeShapeProperties(s.spPr));
  if (s.invertIfNegative !== undefined) {
    parts.push(`<c:invertIfNegative val="${s.invertIfNegative ? '1' : '0'}"/>`);
  }
  if (s.dPt) for (const p of s.dPt) parts.push(serializeDataPoint(p));
  if (s.dLbls) parts.push(serializeDataLabelList(s.dLbls));
  if (s.trendline) for (const t of s.trendline) parts.push(serializeTrendline(t));
  if (s.errBars) for (const e of s.errBars) parts.push(serializeErrBars(e));
  if (s.xVal) parts.push(serializeNumericRef('xVal', s.xVal));
  parts.push(serializeNumericRef('yVal', s.yVal));
  parts.push(serializeNumericRef('bubbleSize', s.bubbleSize));
  if (s.bubble3D !== undefined) parts.push(`<c:bubble3D val="${s.bubble3D ? '1' : '0'}"/>`);
  parts.push('</c:ser>');
  return parts.join('');
};

const serializeBubbleChart = (chart: BubbleChart): string => {
  const parts: string[] = ['<c:bubbleChart>'];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeBubbleSeries(s));
  if (chart.bubble3D !== undefined) parts.push(`<c:bubble3D val="${chart.bubble3D ? '1' : '0'}"/>`);
  if (chart.bubbleScale !== undefined) parts.push(`<c:bubbleScale val="${chart.bubbleScale}"/>`);
  if (chart.showNegBubbles !== undefined) parts.push(`<c:showNegBubbles val="${chart.showNegBubbles ? '1' : '0'}"/>`);
  if (chart.sizeRepresents !== undefined) parts.push(`<c:sizeRepresents val="${chart.sizeRepresents}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:bubbleChart>');
  return parts.join('');
};

const serializeHiLowLines = (h: boolean | HiLowLines): string => {
  if (h === true) return '<c:hiLowLines/>';
  if (h === false) return '';
  if (!h.spPr) return '<c:hiLowLines/>';
  return `<c:hiLowLines>${serializeShapeProperties(h.spPr)}</c:hiLowLines>`;
};

const serializeUpDownBars = (u: boolean | UpDownBars): string => {
  if (u === true) return '<c:upDownBars><c:gapWidth val="150"/></c:upDownBars>';
  if (u === false) return '';
  const parts: string[] = ['<c:upDownBars>'];
  parts.push(`<c:gapWidth val="${u.gapWidth ?? 150}"/>`);
  if (u.upBars) {
    parts.push(u.upBars.spPr ? `<c:upBars>${serializeShapeProperties(u.upBars.spPr)}</c:upBars>` : '<c:upBars/>');
  }
  if (u.downBars) {
    parts.push(
      u.downBars.spPr ? `<c:downBars>${serializeShapeProperties(u.downBars.spPr)}</c:downBars>` : '<c:downBars/>',
    );
  }
  parts.push('</c:upDownBars>');
  return parts.join('');
};

const serializeStockChart = (chart: StockChart): string => {
  const parts: string[] = ['<c:stockChart>'];
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.hiLowLines !== undefined) parts.push(serializeHiLowLines(chart.hiLowLines));
  if (chart.upDownBars !== undefined) parts.push(serializeUpDownBars(chart.upDownBars));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push('</c:stockChart>');
  return parts.join('');
};

const serializeSurfaceChart = (chart: SurfaceChart): string => {
  const parts: string[] = ['<c:surfaceChart>'];
  if (chart.wireframe !== undefined) parts.push(`<c:wireframe val="${chart.wireframe ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:surfaceChart>');
  return parts.join('');
};

const serializeOfPieChart = (chart: OfPieChart): string => {
  const parts: string[] = ['<c:ofPieChart>', `<c:ofPieType val="${chart.ofPieType}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.gapWidth !== undefined) parts.push(`<c:gapWidth val="${chart.gapWidth}"/>`);
  if (chart.splitType !== undefined) parts.push(`<c:splitType val="${chart.splitType}"/>`);
  if (chart.splitPos !== undefined) parts.push(`<c:splitPos val="${chart.splitPos}"/>`);
  if (chart.custSplit !== undefined) {
    parts.push('<c:custSplit>');
    for (const idx of chart.custSplit) parts.push(`<c:secondaryPt idx="${idx}"/>`);
    parts.push('</c:custSplit>');
  }
  if (chart.secondPieSize !== undefined) parts.push(`<c:secondPieSize val="${chart.secondPieSize}"/>`);
  parts.push('</c:ofPieChart>');
  return parts.join('');
};

const serializeBar3DChart = (chart: Bar3DChart): string => {
  const parts: string[] = [
    '<c:bar3DChart>',
    `<c:barDir val="${chart.barDir}"/>`,
    `<c:grouping val="${chart.grouping}"/>`,
  ];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.gapWidth !== undefined) parts.push(`<c:gapWidth val="${chart.gapWidth}"/>`);
  if (chart.gapDepth !== undefined) parts.push(`<c:gapDepth val="${chart.gapDepth}"/>`);
  if (chart.shape !== undefined) parts.push(`<c:shape val="${chart.shape}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:bar3DChart>');
  return parts.join('');
};

const serializeLine3DChart = (chart: Line3DChart): string => {
  const parts: string[] = ['<c:line3DChart>', `<c:grouping val="${chart.grouping}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeLineSeries(s));
  if (chart.gapDepth !== undefined) parts.push(`<c:gapDepth val="${chart.gapDepth}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:line3DChart>');
  return parts.join('');
};

const serializePie3DChart = (chart: Pie3DChart): string => {
  const parts: string[] = ['<c:pie3DChart>'];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push('</c:pie3DChart>');
  return parts.join('');
};

const serializeArea3DChart = (chart: Area3DChart): string => {
  const parts: string[] = ['<c:area3DChart>', `<c:grouping val="${chart.grouping}"/>`];
  if (chart.varyColors !== undefined) parts.push(`<c:varyColors val="${chart.varyColors ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  if (chart.gapDepth !== undefined) parts.push(`<c:gapDepth val="${chart.gapDepth}"/>`);
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:area3DChart>');
  return parts.join('');
};

const serializeSurface3DChart = (chart: Surface3DChart): string => {
  const parts: string[] = ['<c:surface3DChart>'];
  if (chart.wireframe !== undefined) parts.push(`<c:wireframe val="${chart.wireframe ? '1' : '0'}"/>`);
  for (const s of chart.series) parts.push(serializeSeries(s));
  parts.push(`<c:axId val="${chart.axIds[0]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[1]}"/>`);
  parts.push(`<c:axId val="${chart.axIds[2]}"/>`);
  parts.push('</c:surface3DChart>');
  return parts.join('');
};

type AxisTag = 'catAx' | 'valAx' | 'dateAx' | 'serAx';
type AxisLike = CategoryAxis | ValueAxis | DateAxis | SeriesAxis;

const serializeGridlines = (tag: 'c:majorGridlines' | 'c:minorGridlines', g: boolean | Gridlines): string => {
  // `true` -> self-closing element (Excel default rendering).
  // Object form with `spPr` -> wrapped element so the line can be styled.
  if (g === true || g === false || !g.spPr) return `<${tag}/>`;
  return `<${tag}>${serializeShapeProperties(g.spPr)}</${tag}>`;
};

const parseGridlines = (parent: XmlNode, tagName: string): boolean | Gridlines | undefined => {
  const el = findChild(parent, tagName);
  if (!el) return undefined;
  const spPr = parseSpPrSlot(el);
  return spPr ? { spPr } : true;
};

const serializeAxis = (tag: AxisTag, ax: AxisLike): string => {
  // ECMA-376 sequence inside <c:catAx>/<c:valAx>:
  //   axId, scaling, delete, axPos, majorGridlines?, minorGridlines?, title?,
  //   numFmt, majorTickMark, minorTickMark, tickLblPos, spPr?, txPr?, crossAx,
  //   crosses | crossesAt, [catAx: auto, lblAlgn, lblOffset, tickLblSkip,
  //   tickMarkSkip, noMultiLvlLbl] | [valAx: crossBetween, majorUnit, minorUnit,
  //   dispUnits]
  const parts: string[] = [
    `<c:${tag}>`,
    `<c:axId val="${ax.axId}"/>`,
    serializeScaling(ax.scaling),
    `<c:delete val="${ax.delete ? '1' : '0'}"/>`,
    `<c:axPos val="${ax.position ?? (tag === 'catAx' ? 'b' : 'l')}"/>`,
  ];
  if (ax.majorGridlines) parts.push(serializeGridlines('c:majorGridlines', ax.majorGridlines));
  if (ax.minorGridlines) parts.push(serializeGridlines('c:minorGridlines', ax.minorGridlines));
  if (ax.title) parts.push(serializeChartTitle(ax.title));
  parts.push(
    ax.numFmt
      ? serializeNumberFormat(ax.numFmt)
      : '<c:numFmt formatCode="General" sourceLinked="1"/>',
  );
  parts.push(`<c:majorTickMark val="${ax.majorTickMark ?? 'out'}"/>`);
  parts.push(`<c:minorTickMark val="${ax.minorTickMark ?? 'none'}"/>`);
  parts.push(`<c:tickLblPos val="${ax.tickLblPos ?? 'nextTo'}"/>`);
  if (ax.spPr) parts.push(serializeShapeProperties(ax.spPr));
  if (ax.txPr) parts.push(serializeTextBody(ax.txPr, 'c:txPr'));
  parts.push(`<c:crossAx val="${ax.crossAx}"/>`);
  // crossesAt overrides crosses when both are present (mutually exclusive in
  // ECMA-376 — Excel ignores <c:crosses> when <c:crossesAt> is set).
  if (ax.crossesAt !== undefined) {
    parts.push(`<c:crossesAt val="${ax.crossesAt}"/>`);
  } else {
    parts.push(`<c:crosses val="${ax.crosses ?? 'autoZero'}"/>`);
  }
  if (tag === 'catAx') {
    const cat = ax as CategoryAxis;
    parts.push(`<c:auto val="${cat.auto === false ? '0' : '1'}"/>`);
    parts.push(`<c:lblAlgn val="${cat.lblAlgn ?? 'ctr'}"/>`);
    parts.push(`<c:lblOffset val="${cat.lblOffset ?? 100}"/>`);
    parts.push(`<c:noMultiLvlLbl val="${cat.noMultiLvlLbl ? '1' : '0'}"/>`);
  } else if (tag === 'valAx') {
    const val = ax as ValueAxis;
    parts.push(`<c:crossBetween val="${val.crossBetween ?? 'between'}"/>`);
    if (val.majorUnit !== undefined) parts.push(`<c:majorUnit val="${val.majorUnit}"/>`);
    if (val.minorUnit !== undefined) parts.push(`<c:minorUnit val="${val.minorUnit}"/>`);
  } else if (tag === 'dateAx') {
    const date = ax as DateAxis;
    parts.push(`<c:auto val="${date.auto === false ? '0' : '1'}"/>`);
    if (date.lblOffset !== undefined) parts.push(`<c:lblOffset val="${date.lblOffset}"/>`);
    if (date.baseTimeUnit !== undefined) parts.push(`<c:baseTimeUnit val="${date.baseTimeUnit}"/>`);
    if (date.majorUnit !== undefined) parts.push(`<c:majorUnit val="${date.majorUnit}"/>`);
    if (date.majorTimeUnit !== undefined) parts.push(`<c:majorTimeUnit val="${date.majorTimeUnit}"/>`);
    if (date.minorUnit !== undefined) parts.push(`<c:minorUnit val="${date.minorUnit}"/>`);
    if (date.minorTimeUnit !== undefined) parts.push(`<c:minorTimeUnit val="${date.minorTimeUnit}"/>`);
  } else {
    const ser = ax as SeriesAxis;
    if (ser.tickLblSkip !== undefined) parts.push(`<c:tickLblSkip val="${ser.tickLblSkip}"/>`);
    if (ser.tickMarkSkip !== undefined) parts.push(`<c:tickMarkSkip val="${ser.tickMarkSkip}"/>`);
  }
  parts.push(`</c:${tag}>`);
  return parts.join('');
};

const serializeChartTitle = (title: ChartTitle): string => {
  const parts: string[] = ['<c:title>'];
  if (title.tx) {
    // Wrap a TextBody as <c:tx><c:rich>...</c:rich></c:tx>.
    parts.push('<c:tx>');
    parts.push(serializeTextBody(title.tx, 'c:rich'));
    parts.push('</c:tx>');
  } else if (title.text !== undefined) {
    parts.push(
      '<c:tx>',
      '<c:rich>',
      '<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>',
      '<a:lstStyle/>',
      '<a:p><a:pPr><a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">',
      '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>',
      '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>',
      '</a:defRPr></a:pPr>',
      `<a:r><a:rPr lang="en-US"/><a:t>${escapeText(title.text)}</a:t></a:r>`,
      '</a:p>',
      '</c:rich>',
      '</c:tx>',
    );
  }
  // <c:layout/> + <c:overlay> match the order Excel emits.
  parts.push(serializeLayout(title.layout));
  if (title.overlay !== undefined) {
    parts.push(`<c:overlay val="${title.overlay ? '1' : '0'}"/>`);
  } else {
    parts.push('<c:overlay val="0"/>');
  }
  if (title.spPr) parts.push(serializeShapeProperties(title.spPr));
  if (title.txPr) parts.push(serializeTextBody(title.txPr, 'c:txPr'));
  parts.push('</c:title>');
  return parts.join('');
};

const serializeChartKind = (chart: ChartKind): string => {
  switch (chart.kind) {
    case 'bar':
      return serializeBarChart(chart);
    case 'line':
      return serializeLineChart(chart);
    case 'area':
      return serializeAreaChart(chart);
    case 'pie':
      return serializePieChart(chart);
    case 'doughnut':
      return serializeDoughnutChart(chart);
    case 'scatter':
      return serializeScatterChart(chart);
    case 'radar':
      return serializeRadarChart(chart);
    case 'bubble':
      return serializeBubbleChart(chart);
    case 'stock':
      return serializeStockChart(chart);
    case 'surface':
      return serializeSurfaceChart(chart);
    case 'ofPie':
      return serializeOfPieChart(chart);
    case 'bar3D':
      return serializeBar3DChart(chart);
    case 'line3D':
      return serializeLine3DChart(chart);
    case 'pie3D':
      return serializePie3DChart(chart);
    case 'area3D':
      return serializeArea3DChart(chart);
    case 'surface3D':
      return serializeSurface3DChart(chart);
  }
};

const serializePlotArea = (plotArea: PlotArea): string => {
  const parts: string[] = ['<c:plotArea>', serializeLayout(plotArea.layout)];
  parts.push(serializeChartKind(plotArea.chart));
  // Excel rejects charts that reference axIds in a chart-kind element but don't
  // define matching catAx/valAx siblings under plotArea. Fill in sensible
  // defaults when callers omit them.
  const axes = inferAxesForChart(plotArea);
  for (const ax of axes) parts.push(ax);
  if (plotArea.dateAx) parts.push(serializeAxis('dateAx', plotArea.dateAx));
  if (plotArea.serAx) parts.push(serializeAxis('serAx', plotArea.serAx));
  if (plotArea.spPr) parts.push(serializeShapeProperties(plotArea.spPr));
  parts.push('</c:plotArea>');
  return parts.join('');
};

const inferAxesForChart = (plotArea: PlotArea): string[] => {
  const chart = plotArea.chart;
  if (!('axIds' in chart)) return [];
  const ids = chart.axIds;
  if (!ids || ids.length < 2) return [];
  // Scatter / bubble have two value axes (x and y are both numeric).
  const isXValueAxis = chart.kind === 'scatter' || chart.kind === 'bubble';
  const cat = (isXValueAxis ? plotArea.xValAx : undefined) ?? plotArea.catAx ?? { axId: ids[0], crossAx: ids[1], position: 'b' as const };
  const val = plotArea.valAx ?? { axId: ids[1], crossAx: ids[0], position: 'l' as const };
  const out: string[] = [];
  out.push(serializeAxis(isXValueAxis ? 'valAx' : 'catAx', cat));
  out.push(serializeAxis('valAx', val));
  return out;
};

const serializeLegend = (legend: Legend): string => {
  // ECMA-376 sequence inside <c:legend>: legendPos, legendEntry*, layout?,
  // overlay?, spPr?, txPr?.
  const parts: string[] = ['<c:legend>', `<c:legendPos val="${legend.position}"/>`];
  if (legend.layout) parts.push(serializeLayout(legend.layout));
  if (legend.overlay !== undefined) parts.push(`<c:overlay val="${legend.overlay ? '1' : '0'}"/>`);
  if (legend.spPr) parts.push(serializeShapeProperties(legend.spPr));
  if (legend.txPr) parts.push(serializeTextBody(legend.txPr, 'c:txPr'));
  parts.push('</c:legend>');
  return parts.join('');
};

/**
 * Inspect a parsed chart-space root and return the rId referenced by
 * `<c:userShapes r:id="...">` if any. The loader calls this after parsing the
 * chart so it can resolve the chartDrawing part via the chart's rels.
 */
export function findUserShapesRId(bytes: Uint8Array | string): string | undefined {
  const root = parseXml(bytes);
  if (root.name !== CHART_SPACE_TAG) return undefined;
  const us = findChild(root, USER_SHAPES_TAG);
  return us?.attrs[`{${REL_NS}}id`];
}

/** Optional rels rIds the writer resolves before bytes are emitted. */
export interface ChartSerializeOptions {
  /** Drawing-rels id for the `<c:userShapes r:id="...">` reference. */
  userShapesRId?: string;
}

/** Serialise a ChartSpace to its `xl/charts/chartN.xml` bytes. */
export function chartToBytes(space: ChartSpace, opts: ChartSerializeOptions = {}): Uint8Array {
  return new TextEncoder().encode(serializeChartSpace(space, opts));
}

function serializeChartSpace(space: ChartSpace, opts: ChartSerializeOptions = {}): string {
  const parts: string[] = [
    XML_HEADER,
    `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}">`,
    // chartSpace prelude — `date1904`, `lang`, `roundedCorners` mirror what
    // real Excel emits and a few Excel versions reject the file entirely
    // without them.
    '<c:date1904 val="0"/>',
    '<c:lang val="en-US"/>',
    '<c:roundedCorners val="0"/>',
  ];
  if (space.style !== undefined) {
    parts.push(`<c:style val="${space.style}"/>`);
  }
  parts.push('<c:chart>');
  if (space.title !== undefined) parts.push(serializeChartTitle(space.title));
  // <c:autoTitleDeleted> follows <c:title> when present, and is required by
  // some Excel versions when a title is set; emit `0` when we have a title and
  // `1` (= no auto-title) otherwise.
  parts.push(`<c:autoTitleDeleted val="${space.title === undefined ? '1' : '0'}"/>`);
  if (space.view3D) parts.push(serializeView3D(space.view3D));
  if (space.floor) parts.push(serializeSurfaceFrame('floor', space.floor));
  if (space.sideWall) parts.push(serializeSurfaceFrame('sideWall', space.sideWall));
  if (space.backWall) parts.push(serializeSurfaceFrame('backWall', space.backWall));
  parts.push(serializePlotArea(space.plotArea));
  if (space.legend) parts.push(serializeLegend(space.legend));
  parts.push(`<c:plotVisOnly val="${space.plotVisOnly === false ? '0' : '1'}"/>`);
  parts.push(`<c:dispBlanksAs val="${space.dispBlanksAs ?? 'gap'}"/>`);
  parts.push('</c:chart>');
  // chartSpace-level spPr / txPr are siblings of <c:chart>, emitted after.
  if (space.spPr) parts.push(serializeShapeProperties(space.spPr));
  if (space.txPr) parts.push(serializeTextBody(space.txPr, 'c:txPr'));
  // <c:userShapes r:id="..."> sits at the very end of <c:chartSpace>.
  if (opts.userShapesRId !== undefined) {
    parts.push(`<c:userShapes r:id="${opts.userShapesRId}"/>`);
  }
  parts.push('</c:chartSpace>');
  return parts.join('');
}

// SHEET_DRAWING_NS is imported only to keep import surface stable; the chart
// serialiser doesn't need it directly.
void SHEET_DRAWING_NS;
