// Charts: building a chart from a block of cells (Insert ▸ Chart), and
// rebuilding one from its own data refs (Switch Row/Column, Select Data,
// Change Chart Type). Reading a chart for display lives in chart-data.ts.
//
// Every chart is built from the same split of the block into series names,
// categories and values (`seriesSpecs`), so changing the type of an existing
// chart is "find its source block, build the new type over it, carry the
// formatting across".

import type {
  BarSeries,
  BubbleSeries,
  ChartSpace,
  LineSeries,
  PlotArea,
  ScatterSeries,
  ScatterStyle,
} from '@office-kit/xlsx/chart';
import {
  makeArea3DChart,
  makeAreaChart,
  makeBar3DChart,
  makeBarChart,
  makeBarSeries,
  makeBoxWhiskerChart,
  makeBubbleChart,
  makeBubbleSeries,
  makeChartSpace,
  makeDoughnutChart,
  makeFunnelChart,
  makeHistogramChart,
  makeLine3DChart,
  makeLineChart,
  makeOfPieChart,
  makeParetoChart,
  makePie3DChart,
  makePieChart,
  makeRadarChart,
  makeScatterChart,
  makeScatterSeries,
  makeStockChart,
  makeSunburstChart,
  makeSurface3DChart,
  makeSurfaceChart,
  makeTreemapChart,
  makeWaterfallChart,
  type CxChartSpace,
  type CxDim,
} from '@office-kit/xlsx/chart';
import type { ChartReference } from '@office-kit/xlsx/drawing';
import { addChartAt } from '@office-kit/xlsx/drawing';
import { coerceToText, type CalcScalar } from '../calc/index.ts';
import type { Range } from './address.ts';
import { cellAddress, colLetter, MAX_COL, parseRangeAddress, quoteSheetName, unionRange } from './address.ts';
import { isBlank as isBlankCell, usedRange } from './cells.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';
import { currentRegion } from './navigation.ts';
import { currentRange } from './selection.ts';

/** Every chart type the Insert ribbon and Change Chart Type offer, in Excel's gallery order. */
export const CHART_GROUPS = {
  column: ['columnClustered', 'columnStacked', 'columnPercent', 'column3DClustered', 'column3DStacked', 'column3DPercent', 'column3D'],
  bar: ['barClustered', 'barStacked', 'barPercent', 'bar3DClustered', 'bar3DStacked', 'bar3DPercent'],
  line: ['line', 'lineStacked', 'linePercent', 'lineMarkers', 'lineStackedMarkers', 'linePercentMarkers', 'line3D'],
  area: ['area', 'areaStacked', 'areaPercent', 'area3D', 'area3DStacked', 'area3DPercent'],
  pie: ['pie', 'pie3D', 'pieOfPie', 'barOfPie', 'doughnut'],
  scatter: ['scatter', 'scatterSmoothMarkers', 'scatterSmooth', 'scatterLinesMarkers', 'scatterLines'],
  bubble: ['bubble', 'bubble3D'],
  radar: ['radar', 'radarMarkers', 'radarFilled'],
  stock: ['stockHLC', 'stockOHLC'],
  surface: ['surface3D', 'surfaceWireframe3D', 'contour', 'contourWireframe'],
  hierarchy: ['treemap', 'sunburst'],
  statistic: ['histogram', 'pareto', 'boxWhisker'],
  waterfall: ['waterfall', 'funnel'],
} as const;

export type ChartGroup = keyof typeof CHART_GROUPS;
export type ChartChoice = (typeof CHART_GROUPS)[ChartGroup][number];

export const CHART_GROUP_IDS: readonly ChartGroup[] = ['column', 'bar', 'line', 'area', 'pie', 'scatter', 'bubble', 'radar', 'stock', 'surface', 'hierarchy', 'statistic', 'waterfall'];

export const CHART_CHOICES: readonly ChartChoice[] = CHART_GROUP_IDS.flatMap((g) => CHART_GROUPS[g]);

/** The gallery group a chart type is listed under. */
export function chartGroupOf(choice: ChartChoice): ChartGroup {
  return CHART_GROUP_IDS.find((g) => CHART_GROUPS[g].some((c) => c === choice)) ?? 'column';
}

/** How the selected block splits into series names, categories and values. */
export interface SourceLayout {
  /** The first row holds series names (or, by rows, categories). */
  readonly headerRow: boolean;
  /** The first column holds categories (or, by rows, series names). */
  readonly headerCol: boolean;
  /** One series per column (Excel's choice when there are at least as many rows as columns). */
  readonly byColumns: boolean;
}

export type ValueAt = (row: number, col: number) => CalcScalar;

const isText = (v: CalcScalar): boolean => typeof v === 'string' && v !== '';
const isBlank = (v: CalcScalar): boolean => v === null || v === '';

/**
 * Excel's guess: the first row is a header when it has text and no numbers
 * (ignoring its first cell); the first column is categories when its body has
 * text and no numbers, or when the top-left cell is empty above a header row.
 */
export function analyzeSource(range: Range, valueAt: ValueAt): SourceLayout {
  const rows = range.r2 - range.r1 + 1;
  const cols = range.c2 - range.c1 + 1;
  const allLabels = (cells: CalcScalar[]): boolean => cells.some(isText) && cells.every((v) => isText(v) || isBlank(v));
  const firstRow: CalcScalar[] = [];
  for (let c = range.c1 + 1; c <= range.c2; c++) firstRow.push(valueAt(range.r1, c));
  const headerRow = rows > 1 && (cols === 1 ? isText(valueAt(range.r1, range.c1)) : allLabels(firstRow));
  const firstCol: CalcScalar[] = [];
  for (let r = range.r1 + (headerRow ? 1 : 0); r <= range.r2; r++) firstCol.push(valueAt(r, range.c1));
  const topLeftBlank = isBlank(valueAt(range.r1, range.c1));
  const headerCol = cols > 1 && (allLabels(firstCol) || (headerRow && topLeftBlank));
  const dataRows = rows - (headerRow ? 1 : 0);
  const dataCols = cols - (headerCol ? 1 : 0);
  return { headerRow, headerCol, byColumns: dataRows >= dataCols };
}

function absRef(sheet: string, r1: number, c1: number, r2: number, c2: number): string {
  const a = cellAddress(r1, c1, true);
  return `${quoteSheetName(sheet)}!${r1 === r2 && c1 === c2 ? a : `${a}:${cellAddress(r2, c2, true)}`}`;
}

interface SeriesSpec {
  readonly name?: string;
  readonly nameRef?: string;
  readonly valRef: string;
  readonly values: number[];
}

interface Specs {
  readonly series: SeriesSpec[];
  readonly catRef?: string;
  readonly categories?: string[];
}

const textOf = (v: CalcScalar): string => {
  if (v === null) return '';
  const t = coerceToText(v);
  return typeof t === 'string' ? t : t.code;
};

/** Split `range` into series value refs, a category ref and series name refs per `layout`. */
export function seriesSpecs(sheet: string, range: Range, layout: SourceLayout, valueAt: ValueAt): Specs {
  const r0 = range.r1 + (layout.headerRow ? 1 : 0);
  const c0 = range.c1 + (layout.headerCol ? 1 : 0);
  const num = (v: CalcScalar): number => (typeof v === 'number' ? v : 0);
  const series: SeriesSpec[] = [];
  if (layout.byColumns) {
    for (let c = c0; c <= range.c2; c++) {
      const values: number[] = [];
      for (let r = r0; r <= range.r2; r++) values.push(num(valueAt(r, c)));
      series.push({
        valRef: absRef(sheet, r0, c, range.r2, c),
        values,
        ...(layout.headerRow ? { nameRef: absRef(sheet, range.r1, c, range.r1, c), name: textOf(valueAt(range.r1, c)) } : {}),
      });
    }
    if (!layout.headerCol) return { series };
    const categories: string[] = [];
    for (let r = r0; r <= range.r2; r++) categories.push(textOf(valueAt(r, range.c1)));
    return { series, catRef: absRef(sheet, r0, range.c1, range.r2, range.c1), categories };
  }
  for (let r = r0; r <= range.r2; r++) {
    const values: number[] = [];
    for (let c = c0; c <= range.c2; c++) values.push(num(valueAt(r, c)));
    series.push({
      valRef: absRef(sheet, r, c0, r, range.c2),
      values,
      ...(layout.headerCol ? { nameRef: absRef(sheet, r, range.c1, r, range.c1), name: textOf(valueAt(r, range.c1)) } : {}),
    });
  }
  if (!layout.headerRow) return { series };
  const categories: string[] = [];
  for (let c = c0; c <= range.c2; c++) categories.push(textOf(valueAt(range.r1, c)));
  return { series, catRef: absRef(sheet, range.r1, c0, range.r1, range.c2), categories };
}

// ---- building ---------------------------------------------------------------------

type Grouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';

const groupingOf = (choice: string): Grouping => (choice.includes('Percent') ? 'percentStacked' : choice.includes('Stacked') ? 'stacked' : 'clustered');

const VIEW_3D = { rotX: 15, rotY: 20, rAngAx: true } as const;
const SERIES_AXIS = { serAx: { axId: 3, crossAx: 2, position: 'b' as const } };

function barSeriesOf(specs: Specs): BarSeries[] {
  return specs.series.map((s, idx) =>
    makeBarSeries({
      idx,
      val: { ref: s.valRef, cache: s.values },
      ...(specs.catRef ? { cat: { ref: specs.catRef, cacheKind: 'str' as const, ...(specs.categories ? { cache: specs.categories } : {}) } } : {}),
      ...(s.nameRef ? { tx: { kind: 'ref' as const, ref: s.nameRef } } : {}),
    }),
  );
}

const noLine = { ln: { fill: { kind: 'noFill' as const } } };

/** Scatter / bubble: the first value column (or the category column) is X. */
function xyParts(specs: Specs): { xs: { ref: string; cache?: number[] } | undefined; rest: SeriesSpec[] } {
  if (specs.catRef) return { xs: { ref: specs.catRef }, rest: specs.series };
  const first = specs.series[0];
  if (!first || specs.series.length < 2) return { xs: undefined, rest: specs.series };
  return { xs: { ref: first.valRef, cache: first.values }, rest: specs.series.slice(1) };
}

function scatterChart(specs: Specs, choice: ChartChoice): ScatterSeries[] {
  const { xs, rest } = xyParts(specs);
  const smooth = choice === 'scatterSmooth' || choice === 'scatterSmoothMarkers';
  const markers = choice === 'scatter' || choice.endsWith('Markers');
  const lines = choice !== 'scatter';
  return rest.map((s, idx) => {
    const series = makeScatterSeries({
      idx,
      yVal: { ref: s.valRef, cache: s.values },
      ...(xs ? { xVal: xs } : {}),
      ...(s.nameRef ? { tx: { kind: 'ref' as const, ref: s.nameRef } } : {}),
      marker: markers ? { symbol: 'circle', size: 5 } : { symbol: 'none' },
      smooth,
    });
    // Excel's plain Scatter is markers only: the series line is switched off.
    if (!lines) series.spPr = noLine;
    return series;
  });
}

function bubbleSeriesOf(specs: Specs, bubble3D = false): BubbleSeries[] {
  const { xs, rest } = xyParts(specs);
  const out: BubbleSeries[] = [];
  // After X, columns come in (Y, size) pairs.
  for (let i = 0; i < rest.length; i += 2) {
    const y = rest[i];
    if (!y) break;
    // A lone Y column doubles as its own sizes rather than producing no bubbles.
    const size = rest[i + 1] ?? y;
    out.push(
      makeBubbleSeries({
        idx: out.length,
        yVal: { ref: y.valRef, cache: y.values },
        bubbleSize: { ref: size.valRef, cache: size.values },
        bubble3D,
        ...(xs ? { xVal: xs } : {}),
        ...(y.nameRef ? { tx: { kind: 'ref' as const, ref: y.nameRef } } : {}),
      }),
    );
  }
  return out;
}

function plotAreaFor(choice: ChartChoice, specs: Specs): { plotArea: PlotArea; view3D?: ChartSpace['view3D'] } {
  const series = barSeriesOf(specs);
  const horizontal = choice.startsWith('bar') && choice !== 'barOfPie';
  const axes = {
    catAx: { axId: 1, crossAx: 2, position: horizontal ? ('l' as const) : ('b' as const) },
    valAx: { axId: 2, crossAx: 1, position: horizontal ? ('b' as const) : ('l' as const), majorGridlines: true },
  };
  const grouping = groupingOf(choice);
  const stackedish = grouping !== 'clustered';
  switch (choice) {
    case 'columnClustered':
    case 'columnStacked':
    case 'columnPercent':
    case 'barClustered':
    case 'barStacked':
    case 'barPercent':
      return {
        plotArea: {
          chart: makeBarChart({ barDir: horizontal ? 'bar' : 'col', grouping, series, gapWidth: stackedish ? 150 : 219, overlap: stackedish ? 100 : -27 }),
          ...axes,
        },
      };
    case 'column3DClustered':
    case 'column3DStacked':
    case 'column3DPercent':
    case 'bar3DClustered':
    case 'bar3DStacked':
    case 'bar3DPercent':
      // Excel writes a placeholder third axis id of 0 and no series axis for non-standard 3-D bars.
      return {
        plotArea: { chart: makeBar3DChart({ barDir: horizontal ? 'bar' : 'col', grouping, series, gapWidth: 150, shape: 'box', axIds: [1, 2, 0] }), ...axes },
        view3D: VIEW_3D,
      };
    case 'column3D':
      return { plotArea: { chart: makeBar3DChart({ barDir: 'col', grouping: 'standard', series, gapWidth: 150, shape: 'box', axIds: [1, 2, 3] }), ...axes, ...SERIES_AXIS }, view3D: { ...VIEW_3D, rAngAx: false } };
    case 'line':
    case 'lineStacked':
    case 'linePercent':
    case 'lineMarkers':
    case 'lineStackedMarkers':
    case 'linePercentMarkers': {
      const markers = choice.endsWith('Markers');
      const lineSeries: LineSeries[] = series.map((s) => ({ ...s, marker: markers ? { symbol: 'circle', size: 5 } : { symbol: 'none' } }));
      return { plotArea: { chart: makeLineChart({ grouping: grouping === 'clustered' ? 'standard' : grouping, series: lineSeries }), ...axes } };
    }
    case 'line3D':
      return { plotArea: { chart: makeLine3DChart({ series, axIds: [1, 2, 3] }), ...axes, ...SERIES_AXIS }, view3D: { ...VIEW_3D, rAngAx: false } };
    case 'area':
    case 'areaStacked':
    case 'areaPercent':
      return { plotArea: { chart: makeAreaChart({ grouping: grouping === 'clustered' ? 'standard' : grouping, series }), ...axes } };
    case 'area3D':
      return { plotArea: { chart: makeArea3DChart({ grouping: 'standard', series, axIds: [1, 2, 3] }), ...axes, ...SERIES_AXIS }, view3D: { ...VIEW_3D, rAngAx: false } };
    case 'area3DStacked':
    case 'area3DPercent':
      return { plotArea: { chart: makeArea3DChart({ grouping, series, axIds: [1, 2, 0] }), ...axes }, view3D: VIEW_3D };
    case 'pie':
      return { plotArea: { chart: makePieChart({ series: series.slice(0, 1), varyColors: true }) } };
    case 'pie3D':
      return { plotArea: { chart: makePie3DChart({ series: series.slice(0, 1), varyColors: true }) }, view3D: { rotX: 30, rotY: 0, rAngAx: false } };
    case 'pieOfPie':
    case 'barOfPie':
      return { plotArea: { chart: makeOfPieChart({ ofPieType: choice === 'pieOfPie' ? 'pie' : 'bar', series: series.slice(0, 1), varyColors: true, gapWidth: 100, secondPieSize: 75 }) } };
    case 'doughnut':
      return { plotArea: { chart: makeDoughnutChart({ series, varyColors: true, holeSize: 75, firstSliceAng: 0 }) } };
    case 'scatter':
    case 'scatterSmoothMarkers':
    case 'scatterSmooth':
    case 'scatterLinesMarkers':
    case 'scatterLines': {
      const style: ScatterStyle = choice === 'scatterSmooth' || choice === 'scatterSmoothMarkers' ? 'smoothMarker' : 'lineMarker';
      return {
        plotArea: {
          chart: makeScatterChart({ scatterStyle: style, series: scatterChart(specs, choice) }),
          catAx: { axId: 1, crossAx: 2, position: 'b' },
          valAx: { axId: 2, crossAx: 1, position: 'l', majorGridlines: true },
        },
      };
    }
    case 'bubble':
    case 'bubble3D':
      return {
        plotArea: {
          chart: makeBubbleChart({ series: bubbleSeriesOf(specs, choice === 'bubble3D'), bubbleScale: 100, sizeRepresents: 'area' }),
          catAx: { axId: 1, crossAx: 2, position: 'b' },
          valAx: { axId: 2, crossAx: 1, position: 'l', majorGridlines: true },
        },
      };
    case 'radar':
    case 'radarMarkers':
    case 'radarFilled': {
      // Radar series carry no <c:marker>; the chart-level style alone decides whether markers show.
      const style = choice === 'radarFilled' ? 'filled' : choice === 'radarMarkers' ? 'marker' : 'standard';
      return { plotArea: { chart: makeRadarChart({ radarStyle: style, series }), ...axes } };
    }
    case 'stockHLC':
    case 'stockOHLC': {
      // Stock series are drawn by their hi-low lines and up/down bars, not by series lines.
      const stockSeries: LineSeries[] = series.map((s) => ({ ...s, spPr: noLine, marker: { symbol: 'none' } }));
      return {
        plotArea: {
          chart: makeStockChart({ series: stockSeries, hiLowLines: true, ...(choice === 'stockOHLC' ? { upDownBars: { gapWidth: 150 } } : {}) }),
          ...axes,
        },
      };
    }
    case 'surface3D':
    case 'surfaceWireframe3D':
      return {
        plotArea: { chart: makeSurface3DChart({ series, wireframe: choice === 'surfaceWireframe3D' }), ...axes, ...SERIES_AXIS },
        view3D: { rotX: 15, rotY: 20, rAngAx: false, perspective: 30 },
      };
    case 'contour':
    case 'contourWireframe':
      return {
        plotArea: { chart: makeSurfaceChart({ series, wireframe: choice === 'contourWireframe' }), ...axes, ...SERIES_AXIS },
        view3D: { rotX: 90, rotY: 0, rAngAx: false, perspective: 0 },
      };
    // Chartex kinds are built by `cxSpaceFor`; this path is never taken for them.
    case 'treemap':
    case 'sunburst':
    case 'histogram':
    case 'pareto':
    case 'boxWhisker':
    case 'waterfall':
    case 'funnel':
      return { plotArea: { chart: makeBarChart({ series }), ...axes } };
  }
}

const CX_CHOICES: ReadonlySet<ChartChoice> = new Set(['treemap', 'sunburst', 'histogram', 'pareto', 'boxWhisker', 'waterfall', 'funnel']);

/** Whether `choice` is an Excel 2016 chart (`cx:` chartex part) rather than a classic `c:` chart. */
export function isCxChoice(choice: ChartChoice): boolean {
  return CX_CHOICES.has(choice);
}

/**
 * Chartex charts take one category column and one value column. The point
 * caches are filled in because Excel draws a chartex chart from its cache
 * until the workbook recalculates.
 */
function cxSpaceFor(choice: ChartChoice, specs: Specs): CxChartSpace {
  const first = specs.series[0];
  const catRef = specs.catRef;
  const valRef = first?.valRef;
  const refs = { ...(catRef ? { catRef } : {}), ...(valRef ? { valRef } : {}) };
  let space: CxChartSpace;
  switch (choice) {
    case 'treemap':
      space = makeTreemapChart(refs);
      break;
    case 'sunburst':
      space = makeSunburstChart(refs);
      break;
    case 'histogram':
      space = makeHistogramChart({ ...(valRef ? { valRef } : {}), binCountAuto: true });
      break;
    case 'pareto':
      space = makeParetoChart(refs);
      break;
    case 'boxWhisker':
      space = makeBoxWhiskerChart({ ...refs, meanMarker: true, outliers: true, quartileMethod: 'exclusive' });
      break;
    case 'funnel':
      space = makeFunnelChart(refs);
      break;
    default:
      space = makeWaterfallChart(refs);
  }
  for (const data of space.chartData.data) {
    for (const dim of data.dims) fillDim(dim, dim.kind === 'str' ? (specs.categories ?? []) : (first?.values ?? []).map(String));
  }
  if (first?.name) space.chart.title = { text: first.name };
  return space;
}

function fillDim(dim: CxDim, values: readonly string[]): void {
  dim.ptCount = values.length;
  dim.pts = values.map((v, idx) => ({ idx, v }));
}

/** Build the chart Excel would create for `choice` over `range`; `byColumns` overrides the orientation guess. */
export function buildChart(choice: ChartChoice, sheet: string, range: Range, valueAt: ValueAt, byColumns?: boolean): ChartReference {
  const guessed = analyzeSource(range, valueAt);
  const layout = byColumns === undefined ? guessed : { ...guessed, byColumns };
  const specs = seriesSpecs(sheet, range, layout, valueAt);
  if (isCxChoice(choice)) return { cxSpace: cxSpaceFor(choice, specs) };
  const { plotArea, view3D } = plotAreaFor(choice, specs);
  const pieLike = CHART_GROUPS.pie.some((c) => c === choice);
  const single = specs.series.length === 1 || (pieLike && choice !== 'doughnut');
  const title = single ? specs.series[0]?.name : undefined;
  const space = makeChartSpace({
    plotArea,
    ...(title ? { title } : {}),
    ...(single && !pieLike ? {} : { legend: { position: 'b' } }),
    ...(view3D ? { view3D } : {}),
  });
  return { space };
}

/**
 * The block a new chart plots: the selection, or the data region around a
 * single selected cell (blank or not, so the empty corner of a table works) as in Excel, cut to the used range so a chart over
 * whole columns costs the data rather than a million rows. Undefined when
 * there is nothing to plot.
 */
export function chartRangeOf(doc: SpreadsheetEditor, selected: Range = currentRange(doc.selection)): Range | undefined {
  const single = selected.r1 === selected.r2 && selected.c1 === selected.c2;
  const r = single ? currentRegion(doc.ws, doc.selection.active) : selected;
  const used = usedRange(doc.ws);
  const lone = r.r1 === r.r2 && r.c1 === r.c2;
  if (!used || (lone && isBlankCell(doc.cellAt(r.r1, r.c1)))) return undefined;
  const cut = { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, used.r2), c2: Math.min(r.c2, used.c2) };
  return cut.r1 <= cut.r2 && cut.c1 <= cut.c2 ? cut : undefined;
}

/** Excel's default chart size: 5 × 3 inches. */
export const DEFAULT_CHART_PX = { width: 480, height: 288 } as const;

/** Insert ▸ Chart: build the chart for `range` and place it to the right of the data. One undo step; returns the new item's index (undefined when sheet protection refused it). */
export function insertChart(doc: SpreadsheetEditor, choice: ChartChoice, range: Range, byColumns?: boolean): number | undefined {
  const ws = doc.ws;
  const sheet = ws.title;
  const chart = buildChart(choice, sheet, range, (r, c) => doc.calc.cellValue(sheet, r, c), byColumns);
  const at = `${colLetter(Math.min(MAX_COL, range.c2 + 2))}${range.r1}`;
  return doc.transact('Insert Chart', (tx) => {
    tx.sheet(ws, 'drawing');
    addChartAt(ws, at, chart, { widthPx: DEFAULT_CHART_PX.width, heightPx: DEFAULT_CHART_PX.height });
    return (ws.drawing?.items.length ?? 1) - 1;
  });
}

// ---- rebuilding from an existing chart -----------------------------------------

/** The cell block a chart was built from, recovered from its series refs. */
export interface ChartSource {
  readonly sheet: string;
  readonly range: Range;
  readonly byColumns: boolean;
}

/** Every worksheet ref a chart reads: series names, categories, values, X / sizes. */
export function chartRefs(chart: ChartReference): { values: string[]; others: string[] } {
  const values: string[] = [];
  const others: string[] = [];
  const space = chart.space;
  if (space) {
    for (const s of space.plotArea.chart.series) {
      if ('val' in s) values.push(s.val.ref);
      if ('yVal' in s) values.push(s.yVal.ref);
      if ('xVal' in s && s.xVal) others.push(s.xVal.ref);
      if ('bubbleSize' in s) values.push(s.bubbleSize.ref);
      if ('cat' in s && s.cat) others.push(s.cat.ref);
      if (s.tx?.kind === 'ref') others.push(s.tx.ref);
    }
  }
  for (const data of chart.cxSpace?.chartData.data ?? []) {
    for (const dim of data.dims) if (dim.f) (dim.kind === 'num' ? values : others).push(dim.f);
  }
  for (const s of chart.cxSpace?.chart.plotArea.series ?? []) if (s.tx?.f) others.push(s.tx.f);
  return { values, others };
}

/**
 * The single block on one sheet that all of a chart's refs fall in, with the
 * orientation its value refs run in. Undefined when the refs span sheets or
 * cannot be parsed (a chart built by hand from scattered ranges).
 */
export function chartSource(chart: ChartReference): ChartSource | undefined {
  const { values, others } = chartRefs(chart);
  let sheet: string | undefined;
  let range: Range | undefined;
  let byColumns: boolean | undefined;
  for (const ref of [...values, ...others]) {
    const parsed = parseRangeAddress(ref);
    if (!parsed?.sheet) return undefined;
    if (sheet !== undefined && parsed.sheet !== sheet) return undefined;
    sheet = parsed.sheet;
    range = range ? unionRange(range, parsed.range) : parsed.range;
    if (byColumns === undefined && values.includes(ref)) {
      const r = parsed.range;
      if (r.r1 !== r.r2 || r.c1 !== r.c2) byColumns = r.c1 === r.c2;
    }
  }
  if (sheet === undefined || !range) return undefined;
  return { sheet, range, byColumns: byColumns ?? true };
}

/** The type a chart would be inserted as, for Change Chart Type's current selection. */
export function chartChoiceOf(chart: ChartReference): ChartChoice | undefined {
  const cx = chart.cxSpace?.chart.plotArea.series[0];
  if (cx) {
    switch (cx.layoutId) {
      case 'treemap':
      case 'sunburst':
      case 'funnel':
      case 'waterfall':
      case 'boxWhisker':
        return cx.layoutId;
      case 'clusteredColumn':
        return chart.cxSpace?.chart.plotArea.series.some((s) => s.layoutId === 'paretoLine') ? 'pareto' : 'histogram';
      default:
        return undefined;
    }
  }
  const c = chart.space?.plotArea.chart;
  if (!c) return undefined;
  const suffix = (g: Grouping): string => (g === 'stacked' ? 'Stacked' : g === 'percentStacked' ? 'Percent' : '');
  const pick = (id: string): ChartChoice | undefined => CHART_CHOICES.find((x) => x === id);
  switch (c.kind) {
    case 'bar':
      return pick(`${c.barDir === 'bar' ? 'bar' : 'column'}${suffix(c.grouping) || 'Clustered'}`);
    case 'bar3D':
      return c.grouping === 'standard' ? 'column3D' : pick(`${c.barDir === 'bar' ? 'bar' : 'column'}3D${suffix(c.grouping) || 'Clustered'}`);
    case 'line': {
      const markers = c.series.some((s) => s.marker?.symbol !== 'none');
      return pick(`line${suffix(c.grouping)}${markers ? 'Markers' : ''}`);
    }
    case 'line3D':
      return 'line3D';
    case 'area':
      return pick(`area${suffix(c.grouping)}`);
    case 'area3D':
      return pick(`area3D${suffix(c.grouping)}`);
    case 'pie':
      return 'pie';
    case 'pie3D':
      return 'pie3D';
    case 'ofPie':
      return c.ofPieType === 'pie' ? 'pieOfPie' : 'barOfPie';
    case 'doughnut':
      return 'doughnut';
    case 'scatter': {
      const markers = c.series.some((s) => s.marker?.symbol !== 'none');
      const lines = c.series.some((s) => s.spPr?.ln?.fill?.kind !== 'noFill');
      if (!lines) return 'scatter';
      if (c.scatterStyle.startsWith('smooth')) return markers ? 'scatterSmoothMarkers' : 'scatterSmooth';
      return markers ? 'scatterLinesMarkers' : 'scatterLines';
    }
    case 'bubble':
      return c.series.some((s) => s.bubble3D) ? 'bubble3D' : 'bubble';
    case 'radar':
      return c.radarStyle === 'filled' ? 'radarFilled' : c.radarStyle === 'marker' ? 'radarMarkers' : 'radar';
    case 'stock':
      return c.upDownBars ? 'stockOHLC' : 'stockHLC';
    case 'surface3D':
      return c.wireframe ? 'surfaceWireframe3D' : 'surface3D';
    case 'surface':
      return c.wireframe ? 'contourWireframe' : 'contour';
  }
}

/**
 * Rebuild `chart` as `choice` over `source`, keeping what the user set on the
 * old one that still applies: title, legend, chart and plot area formatting,
 * axis titles and per-series formatting by position.
 */
export function rebuildChart(chart: ChartReference, choice: ChartChoice, source: ChartSource, valueAt: ValueAt): ChartReference {
  const next = buildChart(choice, source.sheet, source.range, valueAt, source.byColumns);
  const old = chart.space;
  const space = next.space;
  if (old && space) {
    if (old.title) space.title = old.title;
    else delete space.title;
    if (old.legend) space.legend = old.legend;
    else delete space.legend;
    if (old.spPr) space.spPr = old.spPr;
    if (old.txPr) space.txPr = old.txPr;
    if (old.plotArea.spPr) space.plotArea.spPr = old.plotArea.spPr;
    const oldAx = old.plotArea;
    if (space.plotArea.catAx && oldAx.catAx?.title) space.plotArea.catAx.title = oldAx.catAx.title;
    if (space.plotArea.valAx && oldAx.valAx) {
      if (oldAx.valAx.title) space.plotArea.valAx.title = oldAx.valAx.title;
      if (oldAx.valAx.majorGridlines !== undefined) space.plotArea.valAx.majorGridlines = oldAx.valAx.majorGridlines;
    }
    const oldSeries = old.plotArea.chart.series;
    space.plotArea.chart.series.forEach((s, i) => {
      const prev = oldSeries[i];
      if (prev?.spPr && !s.spPr) s.spPr = prev.spPr;
      if (prev?.dLbls) s.dLbls = prev.dLbls;
    });
  }
  const cx = next.cxSpace;
  if (cx) {
    const oldTitle = old?.title?.text ?? chart.cxSpace?.chart.title?.text;
    if (oldTitle) cx.chart.title = { text: oldTitle };
    const spPr = old?.spPr ?? chart.cxSpace?.spPr;
    if (spPr) cx.spPr = spPr;
  }
  return next;
}
