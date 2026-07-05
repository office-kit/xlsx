import type { ShapeProperties } from '../drawing/dml/shape-properties';
import type { TextBody } from '../drawing/dml/text';
import type { ChartDrawing } from './user-shapes';

// ChartML data model.
//
// **Stage 1**: BarChart end-to-end. The full 17 SpreadsheetML chart kinds + 8
// chartex kinds land across the upcoming iterations; this commit introduces the
// shared shape so future kinds can plug in without breaking the public surface:
//
// ChartSpace
//     ├─ title?: string
//     ├─ legend?: { position: ... }
//     └─ plotArea
//          ├─ chart: BarChart | LineChart | PieChart | …  (discriminated union)
//          ├─ catAx?: CategoryAxis  (or Date / Series axis)
//          └─ valAx?: ValueAxis
//
// References to worksheet ranges (`Sheet1!$A$1:$A$5`) are kept as strings; the
// optional `cache` field carries the last-known data so charts can be rendered
// without reading the source data.

export type LegendPosition = 'r' | 't' | 'l' | 'b' | 'tr';
export type GroupingType = 'clustered' | 'stacked' | 'percentStacked' | 'standard';
export type BarDirection = 'bar' | 'col';

/** Reference to a worksheet range plus an optional client-side cache of the resolved values. */
export interface NumericRef {
  /** Worksheet-qualified range string, e.g. `Sheet1!$B$1:$B$5`. */
  ref: string;
  /** Optional cached numeric values — Excel writes these for offline rendering. */
  cache?: number[];
  /** Optional `formatCode` Excel uses when rendering each cached value. */
  formatCode?: string;
}

export interface CategoryRef {
  ref: string;
  /** Whether the cache is numeric or string. */
  cacheKind?: 'num' | 'str';
  /** String values (when `cacheKind === 'str'`) or numeric values. */
  cache?: ReadonlyArray<string | number>;
  formatCode?: string;
}

// ---- Series decorations: data labels, trendlines, error bars --------------

/**
 * Number-format payload for chart axes / data labels (`<c:numFmt formatCode="…"
 * sourceLinked="0|1"/>`).
 *
 * Named `ChartNumberFormat` to avoid colliding with the cell-stylesheet
 * `NumberFormat` exported from `@office-kit/xlsx/styles`, which is a different shape
 * (`{ numFmtId, formatCode }`).
 */
export interface ChartNumberFormat {
  formatCode: string;
  sourceLinked?: boolean;
}

export type DataLabelPosition = 'bestFit' | 'b' | 'ctr' | 'inBase' | 'inEnd' | 'l' | 'outEnd' | 'r' | 't';

/** Per-point data label override. Lives inside `<c:dLbls><c:dLbl idx="N">…</c:dLbl></c:dLbls>`. */
export interface DataLabel {
  idx: number;
  /** Delete this individual label (suppresses display even if the parent dLbls would show it). */
  delete?: boolean;
  /** Inline rich text or a cell reference. Mutually exclusive. */
  tx?: { kind: 'rich'; body: TextBody } | { kind: 'strRef'; ref: string };
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

/** Series-wide data label settings. `<c:dLbls>` element. */
export interface DataLabelList {
  /** When true, suppress all labels (`<c:delete val="1"/>`). */
  delete?: boolean;
  /** Per-point overrides. */
  dLbl?: DataLabel[];
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
  showLeaderLines?: boolean;
}

export type MarkerSymbol =
  | 'auto'
  | 'circle'
  | 'dash'
  | 'diamond'
  | 'dot'
  | 'none'
  | 'picture'
  | 'plus'
  | 'square'
  | 'star'
  | 'triangle'
  | 'x';

/** Data-point marker on line / scatter / radar series (`<c:marker>` child of `<c:ser>`). */
export interface Marker {
  symbol?: MarkerSymbol;
  /** Marker size in points (2..72). */
  size?: number;
  spPr?: ShapeProperties;
}

/**
 * Per-point override on a series (`<c:dPt>` child of `<c:ser>`). Lets each bar
 * / slice / data point pick up its own colour, marker, explosion, etc.
 */
export interface DataPoint {
  /** 0-based index of the data point within the series. */
  idx: number;
  /** Invert fill when the data point's value is negative (bar / area / bubble). */
  invertIfNegative?: boolean;
  /** Per-point marker override (line / scatter / radar). */
  marker?: Marker;
  /** 3-D bubble flag (bubble series only). */
  bubble3D?: boolean;
  /** Slice explosion in % for pie / doughnut points (0..400). */
  explosion?: number;
  /** Per-point shape properties (fill / line colour). */
  spPr?: ShapeProperties;
}

export type TrendlineType = 'exp' | 'linear' | 'log' | 'movingAvg' | 'poly' | 'power';

export interface Trendline {
  /** Display name shown in the chart legend (`<c:name>` text). */
  name?: string;
  spPr?: ShapeProperties;
  trendlineType: TrendlineType;
  /** Polynomial order (2..6) when `trendlineType === 'poly'`. */
  order?: number;
  /** Moving-average period when `trendlineType === 'movingAvg'`. */
  period?: number;
  /** Forecast forward periods. */
  forward?: number;
  /** Forecast backward periods. */
  backward?: number;
  /** Y-axis intercept (linear only). */
  intercept?: number;
  /** Display R² value on the chart. */
  dispRSqr?: boolean;
  /** Display the trendline equation. */
  dispEq?: boolean;
}

export type ErrorBarDirection = 'x' | 'y';
export type ErrorBarType = 'both' | 'minus' | 'plus';
export type ErrorValType = 'cust' | 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr';

export interface ErrorBars {
  /** Direction. Required only on scatter / bubble series; bar / line / area imply `'y'`. */
  errDir?: ErrorBarDirection;
  errBarType: ErrorBarType;
  errValType: ErrorValType;
  noEndCap?: boolean;
  /** Numeric magnitude for fixedVal / percentage / stdDev / stdErr. */
  val?: number;
  /** Custom plus-side data (NumericRef) — required when `errValType === 'cust'`. */
  plus?: NumericRef;
  /** Custom minus-side data (NumericRef) — required when `errValType === 'cust'`. */
  minus?: NumericRef;
  spPr?: ShapeProperties;
}

export interface BarSeries {
  /** 0-based slot in the chart (`<c:idx>`). */
  idx: number;
  /** Render order (`<c:order>`). Usually equals `idx`. */
  order: number;
  /** Series title — either a static string or a cell reference. */
  tx?: { kind: 'literal'; value: string } | { kind: 'ref'; ref: string };
  /** Per-series shape properties (fill / line / effects). */
  spPr?: ShapeProperties;
  /** Invert fill on negative values (bar / line / area series). */
  invertIfNegative?: boolean;
  /**
   * Slice explosion in % (pie / doughnut series only; 0..400). Bar / line / area writers
   * ignore the field — leave it unset for those series.
   */
  explosion?: number;
  /** Per-point overrides. Empty / omitted means every point inherits the series defaults. */
  dPt?: DataPoint[];
  /** Series-wide data labels. */
  dLbls?: DataLabelList;
  /** Trendlines attached to this series. */
  trendline?: Trendline[];
  /** Error bars. Bar / line / area / radar default to a single y-direction entry. */
  errBars?: ErrorBars[];
  /** Categories. */
  cat?: CategoryRef;
  /** Values (always required for a bar series). */
  val: NumericRef;
}

export interface BarChart {
  kind: 'bar';
  /** `bar` for horizontal bars, `col` for vertical columns. */
  barDir: BarDirection;
  /** Excel default is `clustered`. */
  grouping: GroupingType;
  varyColors?: boolean;
  series: BarSeries[];
  /** Bar gap width in % of bar width (Excel default 150). */
  gapWidth?: number;
  /**
   * Bar overlap in -100..100 % of bar width. Positive values overlap clustered bars; negative values
   * space them apart. For stacked / percentStacked grouping, omit and the serializer emits the
   * standard Excel default of 100 (flush stacking).
   */
  overlap?: number;
  /** Internal axis ids. The category and value axes carry the same numbers. */
  axIds: [number, number];
}

export interface LineSeries extends BarSeries {
  /** Per-series smoothing toggle. */
  smooth?: boolean;
  /** Data-point marker (`<c:marker>` child of `<c:ser>`). */
  marker?: Marker;
}

export interface LineChart {
  kind: 'line';
  grouping: GroupingType;
  varyColors?: boolean;
  series: LineSeries[];
  /** Whether to round corners between data points (chart-level default). */
  smooth?: boolean;
  axIds: [number, number];
}

export interface AreaChart {
  kind: 'area';
  grouping: GroupingType;
  varyColors?: boolean;
  series: BarSeries[];
  axIds: [number, number];
}

export interface PieChart {
  kind: 'pie';
  varyColors?: boolean;
  /** Pie / Doughnut have a single ring of slices — but Excel allows multiple series; we mirror that. */
  series: BarSeries[];
}

export interface DoughnutChart {
  kind: 'doughnut';
  varyColors?: boolean;
  series: BarSeries[];
  /** Hole size in % of outer radius (10..90, Excel default 50). */
  holeSize?: number;
  /** First-slice rotation angle in degrees. */
  firstSliceAng?: number;
}

export type ScatterStyle = 'line' | 'lineMarker' | 'marker' | 'none' | 'smooth' | 'smoothMarker';

export interface ScatterSeries {
  idx: number;
  order: number;
  tx?: BarSeries['tx'];
  spPr?: ShapeProperties;
  dPt?: DataPoint[];
  dLbls?: DataLabelList;
  trendline?: Trendline[];
  /** Up to 2 entries (one per direction). */
  errBars?: ErrorBars[];
  /** Data-point marker. */
  marker?: Marker;
  xVal?: NumericRef;
  yVal: NumericRef;
  smooth?: boolean;
}

export interface ScatterChart {
  kind: 'scatter';
  scatterStyle: ScatterStyle;
  varyColors?: boolean;
  series: ScatterSeries[];
  axIds: [number, number];
}

export type RadarStyle = 'standard' | 'marker' | 'filled';

export interface RadarChart {
  kind: 'radar';
  radarStyle: RadarStyle;
  varyColors?: boolean;
  series: BarSeries[];
  axIds: [number, number];
}

export interface BubbleSeries {
  idx: number;
  order: number;
  tx?: BarSeries['tx'];
  spPr?: ShapeProperties;
  invertIfNegative?: boolean;
  dPt?: DataPoint[];
  dLbls?: DataLabelList;
  trendline?: Trendline[];
  /** Up to 2 entries (one per direction). */
  errBars?: ErrorBars[];
  xVal?: NumericRef;
  yVal: NumericRef;
  /** Bubble size — required for a real bubble chart. */
  bubbleSize: NumericRef;
  /** Per-series 3-D toggle. */
  bubble3D?: boolean;
}

export type BubbleSizeRepresents = 'area' | 'w';

export interface BubbleChart {
  kind: 'bubble';
  varyColors?: boolean;
  series: BubbleSeries[];
  bubble3D?: boolean;
  /** Bubble scale 0..300 %. Excel default is 100. */
  bubbleScale?: number;
  showNegBubbles?: boolean;
  sizeRepresents?: BubbleSizeRepresents;
  axIds: [number, number];
}

/** `<c:hiLowLines>` child of stock charts. */
export interface HiLowLines {
  spPr?: ShapeProperties;
}

/** `<c:upBars>` / `<c:downBars>` child frame styling. */
export interface BarFrame {
  spPr?: ShapeProperties;
}

/** `<c:upDownBars>` child of stock charts. */
export interface UpDownBars {
  /** Gap width 0..500 between up/down bars. Excel default is 150. */
  gapWidth?: number;
  upBars?: BarFrame;
  downBars?: BarFrame;
}

export interface StockChart {
  kind: 'stock';
  /** Up to 4 series — typically open / high / low / close. */
  series: BarSeries[];
  /** Boolean flag for the simple `<c:hiLowLines/>` form, or an object with `spPr` for custom line styling. */
  hiLowLines?: boolean | HiLowLines;
  /** Boolean for the simple form, or an object with gapWidth / upBars / downBars detail. */
  upDownBars?: boolean | UpDownBars;
  axIds: [number, number];
}

export interface SurfaceChart {
  kind: 'surface';
  series: BarSeries[];
  /** Wireframe (line-only) when true; smoothed surface fill when false. */
  wireframe?: boolean;
  /** Surfaces use 3 axes: cat + val + ser. */
  axIds: [number, number, number];
}

export type OfPieType = 'bar' | 'pie';
export type SplitType = 'auto' | 'cust' | 'percent' | 'pos' | 'val';

export interface OfPieChart {
  kind: 'ofPie';
  /** `bar` for "Bar of Pie", `pie` for "Pie of Pie". */
  ofPieType: OfPieType;
  varyColors?: boolean;
  series: BarSeries[];
  gapWidth?: number;
  splitType?: SplitType;
  /** Position threshold paired with `splitType='pos'`. */
  splitPos?: number;
  /** Indices of data points moved to the secondary plot when `splitType='cust'`. */
  custSplit?: number[];
  /** Secondary plot size as % of primary (5..200). */
  secondPieSize?: number;
}

// ---- 3-D chart variants ---------------------------------------------------
//
// 3-D charts share most of their attributes with their 2-D counterparts but
// land on different XML tag names (<c:bar3DChart>, etc) and use 3 axes (cat /
// val / ser). We keep them as distinct discriminated-union variants so the
// chart kind stays type-narrowable.

export interface Bar3DChart {
  kind: 'bar3D';
  barDir: BarDirection;
  grouping: GroupingType;
  varyColors?: boolean;
  series: BarSeries[];
  gapWidth?: number;
  /** Bar 3-D adds a `gapDepth` attribute. */
  gapDepth?: number;
  /** Cluster | percentStacked | stacked … plus 'standard' which 2-D doesn't take. */
  shape?: 'cone' | 'coneToMax' | 'box' | 'cylinder' | 'pyramid' | 'pyramidToMax';
  axIds: [number, number, number];
}

export interface Line3DChart {
  kind: 'line3D';
  grouping: GroupingType;
  varyColors?: boolean;
  series: LineSeries[];
  gapDepth?: number;
  axIds: [number, number, number];
}

export interface Pie3DChart {
  kind: 'pie3D';
  varyColors?: boolean;
  series: BarSeries[];
}

export interface Area3DChart {
  kind: 'area3D';
  grouping: GroupingType;
  varyColors?: boolean;
  series: BarSeries[];
  gapDepth?: number;
  axIds: [number, number, number];
}

export interface Surface3DChart {
  kind: 'surface3D';
  series: BarSeries[];
  wireframe?: boolean;
  axIds: [number, number, number];
}

/** Discriminator union of all SpreadsheetML chart kinds modelled so far. */
export type ChartKind =
  | BarChart
  | LineChart
  | AreaChart
  | PieChart
  | DoughnutChart
  | ScatterChart
  | RadarChart
  | BubbleChart
  | StockChart
  | SurfaceChart
  | OfPieChart
  | Bar3DChart
  | Line3DChart
  | Pie3DChart
  | Area3DChart
  | Surface3DChart;

export type TickMark = 'cross' | 'in' | 'none' | 'out';
export type TickLabelPosition = 'high' | 'low' | 'nextTo' | 'none';
export type AxisCrosses = 'autoZero' | 'max' | 'min';
export type AxisOrientation = 'maxMin' | 'minMax';
export type AxisCrossBetween = 'between' | 'midCat';
export type CategoryLabelAlignment = 'ctr' | 'l' | 'r';

/** `<c:scaling>` child of axes. */
export interface AxisScaling {
  orientation?: AxisOrientation;
  min?: number;
  max?: number;
  logBase?: number;
}

interface AxisShared {
  axId: number;
  /** Crosses partner axis id. */
  crossAx: number;
  position?: 'b' | 't' | 'l' | 'r';
  delete?: boolean;
  /** Axis-line / tick formatting. */
  spPr?: ShapeProperties;
  /** Tick-label text formatting. */
  txPr?: TextBody;
  /** Axis title (`<c:title>`). Reuses the chart-title structure. */
  title?: ChartTitle;
  /** Tick-label number format. Default emitted is `General` / `sourceLinked=1`. */
  numFmt?: ChartNumberFormat;
  /** Major tick mark style. Default emitted is `out`. */
  majorTickMark?: TickMark;
  /** Minor tick mark style. Default emitted is `none`. */
  minorTickMark?: TickMark;
  /** Tick-label position. Default emitted is `nextTo`. */
  tickLblPos?: TickLabelPosition;
  /** Axis scaling (`<c:scaling>`). Default emitted is `orientation: 'minMax'`. */
  scaling?: AxisScaling;
  /** `<c:crosses>`. Default emitted is `autoZero`. Mutually exclusive with `crossesAt`. */
  crosses?: AxisCrosses;
  /** `<c:crossesAt val="..."/>`. Numeric cross point on the partner axis. */
  crossesAt?: number;
  /**
   * Major gridlines. `true` keeps Excel's default rendering; pass
   * `{ spPr }` to override the line colour / width / dash. `false` /
   * `undefined` omit the element entirely.
   */
  majorGridlines?: boolean | Gridlines;
  /** Minor gridlines. Same shape as {@link AxisShared.majorGridlines}. */
  minorGridlines?: boolean | Gridlines;
}

/**
 * Rich `<c:majorGridlines>` / `<c:minorGridlines>` form. Wraps a
 * `ShapeProperties` so the gridline line can be styled (e.g. corporate-style
 * light-grey `D9D9D9`).
 */
export interface Gridlines {
  spPr?: ShapeProperties;
}

export interface CategoryAxis extends AxisShared {
  /** `<c:auto>` — whether Excel auto-selects the axis type from the data. */
  auto?: boolean;
  /** Tick-label alignment. Default emitted is `ctr`. */
  lblAlgn?: CategoryLabelAlignment;
  /** Tick-label offset 0..1000. Default emitted is `100`. */
  lblOffset?: number;
  /** Suppress multi-level labels for hierarchical categories. */
  noMultiLvlLbl?: boolean;
}

export interface ValueAxis extends AxisShared {
  /** Tick-axis crossing rule. Default emitted is `between`. */
  crossBetween?: AxisCrossBetween;
  /** Major unit spacing on the value axis. */
  majorUnit?: number;
  /** Minor unit spacing on the value axis. */
  minorUnit?: number;
}

export type TimeUnit = 'days' | 'months' | 'years';

export interface DateAxis extends AxisShared {
  /** `<c:auto>` — auto-select axis type. */
  auto?: boolean;
  /** Tick-label offset 0..1000. Default emitted is `100`. */
  lblOffset?: number;
  /** Smallest interval the axis represents. */
  baseTimeUnit?: TimeUnit;
  /** Major unit spacing on the date axis (count of `majorTimeUnit`s). */
  majorUnit?: number;
  majorTimeUnit?: TimeUnit;
  /** Minor unit spacing on the date axis. */
  minorUnit?: number;
  minorTimeUnit?: TimeUnit;
}

export interface SeriesAxis extends AxisShared {
  /** Skip every Nth tick label (1 = every label). */
  tickLblSkip?: number;
  /** Skip every Nth tick mark. */
  tickMarkSkip?: number;
}

export interface PlotArea {
  chart: ChartKind;
  catAx?: CategoryAxis;
  valAx?: ValueAxis;
  /** Date / time axis. Replaces `catAx` when the categories are dates. */
  dateAx?: DateAxis;
  /** Series axis (used by `surface3DChart` / `surfaceChart`). */
  serAx?: SeriesAxis;
  /** Manual layout for the plot area inside the chart. */
  layout?: Layout;
  /** Plot-area shape properties (background fill, border line). */
  spPr?: ShapeProperties;
}

export type LayoutMode = 'edge' | 'factor';
export type LayoutTarget = 'inner' | 'outer';

/** `<c:manualLayout>` child of `<c:layout>`. All values are fractions of the chart space (0..1). */
export interface ManualLayout {
  layoutTarget?: LayoutTarget;
  xMode?: LayoutMode;
  yMode?: LayoutMode;
  wMode?: LayoutMode;
  hMode?: LayoutMode;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

/** `<c:layout>` wrapper. Currently only `manualLayout` is exposed; the lone `<layoutId>` form is for chartex. */
export interface Layout {
  manualLayout?: ManualLayout;
}

export interface Legend {
  position: LegendPosition;
  overlay?: boolean;
  layout?: Layout;
  spPr?: ShapeProperties;
  txPr?: TextBody;
}

/** Chart title with full DrawingML formatting support. */
export interface ChartTitle {
  /**
   * Plain title text. When set the serializer emits
   * `<c:tx><c:rich><a:p><a:r><a:t>text</a:t></a:r></c:rich></c:tx>`. Mutually
   * exclusive with `tx`.
   */
  text?: string;
  /** Rich text body — overrides `text` when both are present. */
  tx?: TextBody;
  overlay?: boolean;
  layout?: Layout;
  spPr?: ShapeProperties;
  txPr?: TextBody;
}

/** 3-D viewing options. Used by `bar3DChart`, `line3DChart`, `pie3DChart`, `area3DChart`, `surface3DChart`. */
export interface View3D {
  /** X-axis rotation in degrees, -90..90. */
  rotX?: number;
  /** Y-axis rotation in degrees, 0..359. */
  rotY?: number;
  /** Depth as a percentage of chart width, 20..2000. */
  depthPercent?: number;
  /** Height as a percentage of chart width, 5..500. */
  hPercent?: number;
  /** Use right-angle axes (orthographic). When true, perspective is ignored. */
  rAngAx?: boolean;
  /** Perspective angle 0..240 (0 = isometric, 30 = Excel default). */
  perspective?: number;
}

/** 3-D wall / floor frame. Children of `<c:chart>` for bar3D / line3D / area3D / surface3D / pie3D. */
export interface SurfaceFrame {
  /** Wall thickness in % of chart width (0..100). */
  thickness?: number;
  spPr?: ShapeProperties;
}

export interface ChartSpace {
  /** Optional chart title. */
  title?: ChartTitle;
  legend?: Legend;
  plotArea: PlotArea;
  /**
   * Built-in Excel chart-style preset (1..48). Mapped to `<c:style val="N"/>` inside
   * `<c:chartSpace>` and selects one entry of Excel's "Chart Styles" gallery.
   */
  style?: number;
  /** 3-D viewing options (applies to 3-D chart kinds; ignored otherwise by Excel). */
  view3D?: View3D;
  /** 3-D chart floor frame. */
  floor?: SurfaceFrame;
  /** 3-D chart side wall frame. */
  sideWall?: SurfaceFrame;
  /** 3-D chart back wall frame. */
  backWall?: SurfaceFrame;
  /** Honour the formatting hints in cached numeric data when rendering. */
  plotVisOnly?: boolean;
  /** Display blanks as gap, zero, or span — Excel default is `gap`. */
  dispBlanksAs?: 'gap' | 'zero' | 'span';
  /** Chart-space level shape properties (overall frame). */
  spPr?: ShapeProperties;
  /** Chart-space level default text properties. */
  txPr?: TextBody;
  /** Annotations / text boxes / arrows drawn over the chart. Serialised as `xl/drawings/chartDrawingN.xml`. */
  userShapes?: ChartDrawing;
}

export function makeBarChart(opts: {
  barDir?: BarDirection;
  grouping?: GroupingType;
  series?: BarSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
  gapWidth?: number;
  overlap?: number;
}): BarChart {
  return {
    kind: 'bar',
    barDir: opts.barDir ?? 'col',
    grouping: opts.grouping ?? 'clustered',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapWidth !== undefined ? { gapWidth: opts.gapWidth } : {}),
    ...(opts.overlap !== undefined ? { overlap: opts.overlap } : {}),
  };
}

export function makeBarSeries(opts: {
  idx: number;
  order?: number;
  val: NumericRef;
  cat?: CategoryRef;
  tx?: BarSeries['tx'];
}): BarSeries {
  return {
    idx: opts.idx,
    order: opts.order ?? opts.idx,
    val: opts.val,
    ...(opts.cat ? { cat: opts.cat } : {}),
    ...(opts.tx ? { tx: opts.tx } : {}),
  };
}

export function makeChartSpace(opts: {
  plotArea: PlotArea;
  /** Plain string is wrapped in `{ text }`; pass `ChartTitle` for full formatting. */
  title?: string | ChartTitle;
  legend?: Legend;
  style?: number;
  view3D?: View3D;
  floor?: SurfaceFrame;
  sideWall?: SurfaceFrame;
  backWall?: SurfaceFrame;
  plotVisOnly?: boolean;
  dispBlanksAs?: ChartSpace['dispBlanksAs'];
  spPr?: ShapeProperties;
  txPr?: TextBody;
}): ChartSpace {
  const title: ChartTitle | undefined = typeof opts.title === 'string' ? { text: opts.title } : opts.title;
  return {
    plotArea: opts.plotArea,
    ...(title !== undefined ? { title } : {}),
    ...(opts.legend ? { legend: opts.legend } : {}),
    ...(opts.style !== undefined ? { style: opts.style } : {}),
    ...(opts.view3D ? { view3D: opts.view3D } : {}),
    ...(opts.floor ? { floor: opts.floor } : {}),
    ...(opts.sideWall ? { sideWall: opts.sideWall } : {}),
    ...(opts.backWall ? { backWall: opts.backWall } : {}),
    ...(opts.plotVisOnly !== undefined ? { plotVisOnly: opts.plotVisOnly } : {}),
    ...(opts.dispBlanksAs !== undefined ? { dispBlanksAs: opts.dispBlanksAs } : {}),
    ...(opts.spPr ? { spPr: opts.spPr } : {}),
    ...(opts.txPr ? { txPr: opts.txPr } : {}),
  };
}

export function makeLineChart(opts: {
  grouping?: GroupingType;
  series?: LineSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
  smooth?: boolean;
}): LineChart {
  return {
    kind: 'line',
    grouping: opts.grouping ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.smooth !== undefined ? { smooth: opts.smooth } : {}),
  };
}

export function makeAreaChart(opts: {
  grouping?: GroupingType;
  series?: BarSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
}): AreaChart {
  return {
    kind: 'area',
    grouping: opts.grouping ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makePieChart(opts: { series?: BarSeries[]; varyColors?: boolean }): PieChart {
  return {
    kind: 'pie',
    series: opts.series ?? [],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makeDoughnutChart(opts: {
  series?: BarSeries[];
  varyColors?: boolean;
  holeSize?: number;
  firstSliceAng?: number;
}): DoughnutChart {
  return {
    kind: 'doughnut',
    series: opts.series ?? [],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.holeSize !== undefined ? { holeSize: opts.holeSize } : {}),
    ...(opts.firstSliceAng !== undefined ? { firstSliceAng: opts.firstSliceAng } : {}),
  };
}

export function makeScatterChart(opts: {
  scatterStyle?: ScatterStyle;
  series?: ScatterSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
}): ScatterChart {
  return {
    kind: 'scatter',
    scatterStyle: opts.scatterStyle ?? 'lineMarker',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makeScatterSeries(opts: {
  idx: number;
  order?: number;
  tx?: BarSeries['tx'];
  xVal?: NumericRef;
  yVal: NumericRef;
  smooth?: boolean;
  marker?: Marker;
}): ScatterSeries {
  return {
    idx: opts.idx,
    order: opts.order ?? opts.idx,
    yVal: opts.yVal,
    ...(opts.tx ? { tx: opts.tx } : {}),
    ...(opts.xVal ? { xVal: opts.xVal } : {}),
    ...(opts.smooth !== undefined ? { smooth: opts.smooth } : {}),
    ...(opts.marker ? { marker: opts.marker } : {}),
  };
}

export function makeRadarChart(opts: {
  radarStyle?: RadarStyle;
  series?: BarSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
}): RadarChart {
  return {
    kind: 'radar',
    radarStyle: opts.radarStyle ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makeBubbleChart(opts: {
  series?: BubbleSeries[];
  axIds?: [number, number];
  varyColors?: boolean;
  bubble3D?: boolean;
  bubbleScale?: number;
  showNegBubbles?: boolean;
  sizeRepresents?: BubbleSizeRepresents;
}): BubbleChart {
  return {
    kind: 'bubble',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.bubble3D !== undefined ? { bubble3D: opts.bubble3D } : {}),
    ...(opts.bubbleScale !== undefined ? { bubbleScale: opts.bubbleScale } : {}),
    ...(opts.showNegBubbles !== undefined ? { showNegBubbles: opts.showNegBubbles } : {}),
    ...(opts.sizeRepresents !== undefined ? { sizeRepresents: opts.sizeRepresents } : {}),
  };
}

export function makeBubbleSeries(opts: {
  idx: number;
  order?: number;
  tx?: BarSeries['tx'];
  xVal?: NumericRef;
  yVal: NumericRef;
  bubbleSize: NumericRef;
  bubble3D?: boolean;
}): BubbleSeries {
  return {
    idx: opts.idx,
    order: opts.order ?? opts.idx,
    yVal: opts.yVal,
    bubbleSize: opts.bubbleSize,
    ...(opts.tx ? { tx: opts.tx } : {}),
    ...(opts.xVal ? { xVal: opts.xVal } : {}),
    ...(opts.bubble3D !== undefined ? { bubble3D: opts.bubble3D } : {}),
  };
}

export function makeStockChart(opts: {
  series?: BarSeries[];
  axIds?: [number, number];
  hiLowLines?: boolean | HiLowLines;
  upDownBars?: boolean | UpDownBars;
}): StockChart {
  return {
    kind: 'stock',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2],
    ...(opts.hiLowLines !== undefined ? { hiLowLines: opts.hiLowLines } : {}),
    ...(opts.upDownBars !== undefined ? { upDownBars: opts.upDownBars } : {}),
  };
}

export function makeSurfaceChart(opts: {
  series?: BarSeries[];
  wireframe?: boolean;
  axIds?: [number, number, number];
}): SurfaceChart {
  return {
    kind: 'surface',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.wireframe !== undefined ? { wireframe: opts.wireframe } : {}),
  };
}

export function makeOfPieChart(opts: {
  ofPieType?: OfPieType;
  series?: BarSeries[];
  varyColors?: boolean;
  gapWidth?: number;
  splitType?: SplitType;
  splitPos?: number;
  custSplit?: number[];
  secondPieSize?: number;
}): OfPieChart {
  return {
    kind: 'ofPie',
    ofPieType: opts.ofPieType ?? 'pie',
    series: opts.series ?? [],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapWidth !== undefined ? { gapWidth: opts.gapWidth } : {}),
    ...(opts.splitType !== undefined ? { splitType: opts.splitType } : {}),
    ...(opts.splitPos !== undefined ? { splitPos: opts.splitPos } : {}),
    ...(opts.custSplit ? { custSplit: opts.custSplit } : {}),
    ...(opts.secondPieSize !== undefined ? { secondPieSize: opts.secondPieSize } : {}),
  };
}

export function makeBar3DChart(opts: {
  barDir?: BarDirection;
  grouping?: GroupingType;
  series?: BarSeries[];
  axIds?: [number, number, number];
  varyColors?: boolean;
  gapWidth?: number;
  gapDepth?: number;
  shape?: Bar3DChart['shape'];
}): Bar3DChart {
  return {
    kind: 'bar3D',
    barDir: opts.barDir ?? 'col',
    grouping: opts.grouping ?? 'clustered',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapWidth !== undefined ? { gapWidth: opts.gapWidth } : {}),
    ...(opts.gapDepth !== undefined ? { gapDepth: opts.gapDepth } : {}),
    ...(opts.shape !== undefined ? { shape: opts.shape } : {}),
  };
}

export function makeLine3DChart(opts: {
  grouping?: GroupingType;
  series?: LineSeries[];
  axIds?: [number, number, number];
  varyColors?: boolean;
  gapDepth?: number;
}): Line3DChart {
  return {
    kind: 'line3D',
    grouping: opts.grouping ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapDepth !== undefined ? { gapDepth: opts.gapDepth } : {}),
  };
}

export function makePie3DChart(opts: { series?: BarSeries[]; varyColors?: boolean }): Pie3DChart {
  return {
    kind: 'pie3D',
    series: opts.series ?? [],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
  };
}

export function makeArea3DChart(opts: {
  grouping?: GroupingType;
  series?: BarSeries[];
  axIds?: [number, number, number];
  varyColors?: boolean;
  gapDepth?: number;
}): Area3DChart {
  return {
    kind: 'area3D',
    grouping: opts.grouping ?? 'standard',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.varyColors !== undefined ? { varyColors: opts.varyColors } : {}),
    ...(opts.gapDepth !== undefined ? { gapDepth: opts.gapDepth } : {}),
  };
}

export function makeSurface3DChart(opts: {
  series?: BarSeries[];
  wireframe?: boolean;
  axIds?: [number, number, number];
}): Surface3DChart {
  return {
    kind: 'surface3D',
    series: opts.series ?? [],
    axIds: opts.axIds ?? [1, 2, 3],
    ...(opts.wireframe !== undefined ? { wireframe: opts.wireframe } : {}),
  };
}
