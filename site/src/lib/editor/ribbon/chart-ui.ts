// Names and gallery thumbnails for the chart types, shared by the Insert tab's
// chart menus, Insert Chart / Change Chart Type and the Chart Design tab.

import type { CalcScalar } from '../calc/index.ts';
import { chartData, type ChartCalc } from '../core/chart-data.ts';
import { buildChart, CHART_GROUP_IDS, CHART_GROUPS, isCxChoice, type ChartChoice, type ChartGroup } from '../core/charts.ts';
import type { ThemePalette } from '../core/theme.ts';
import { renderChartSvg } from '../grid/chart-render.ts';
import type { MessageKey } from '../i18n/i18n.svelte.ts';

export const CHART_LABELS: Readonly<Record<ChartChoice, MessageKey>> = {
  columnClustered: 'chColumnClustered',
  columnStacked: 'chColumnStacked',
  columnPercent: 'chColumnPercent',
  column3DClustered: 'chColumn3DClustered',
  column3DStacked: 'chColumn3DStacked',
  column3DPercent: 'chColumn3DPercent',
  column3D: 'chColumn3D',
  barClustered: 'chBarClustered',
  barStacked: 'chBarStacked',
  barPercent: 'chBarPercent',
  bar3DClustered: 'chBar3DClustered',
  bar3DStacked: 'chBar3DStacked',
  bar3DPercent: 'chBar3DPercent',
  line: 'chLine',
  lineStacked: 'chLineStacked',
  linePercent: 'chLinePercent',
  lineMarkers: 'chLineMarkers',
  lineStackedMarkers: 'chLineStackedMarkers',
  linePercentMarkers: 'chLinePercentMarkers',
  line3D: 'chLine3D',
  area: 'chArea',
  areaStacked: 'chAreaStacked',
  areaPercent: 'chAreaPercent',
  area3D: 'chArea3D',
  area3DStacked: 'chArea3DStacked',
  area3DPercent: 'chArea3DPercent',
  pie: 'chPie',
  pie3D: 'chPie3D',
  pieOfPie: 'chPieOfPie',
  barOfPie: 'chBarOfPie',
  doughnut: 'chDoughnut',
  scatter: 'chScatter',
  scatterSmoothMarkers: 'chScatterSmoothMarkers',
  scatterSmooth: 'chScatterSmooth',
  scatterLinesMarkers: 'chScatterLinesMarkers',
  scatterLines: 'chScatterLines',
  bubble: 'chBubble',
  bubble3D: 'chBubble3D',
  radar: 'chRadar',
  radarMarkers: 'chRadarMarkers',
  radarFilled: 'chRadarFilled',
  stockHLC: 'chStockHLC',
  stockOHLC: 'chStockOHLC',
  surface3D: 'chSurface3D',
  surfaceWireframe3D: 'chSurfaceWireframe3D',
  contour: 'chContour',
  contourWireframe: 'chContourWireframe',
  treemap: 'chTreemap',
  sunburst: 'chSunburst',
  histogram: 'chHistogram',
  pareto: 'chPareto',
  boxWhisker: 'chBoxWhisker',
  waterfall: 'chWaterfall',
  funnel: 'chFunnel',
};

export const GROUP_LABELS: Readonly<Record<ChartGroup, MessageKey>> = {
  column: 'chGrpColumn',
  bar: 'chGrpBar',
  line: 'chGrpLine',
  area: 'chGrpArea',
  pie: 'chGrpPie',
  scatter: 'chGrpScatter',
  bubble: 'chGrpBubble',
  radar: 'chGrpRadar',
  stock: 'chGrpStock',
  surface: 'chGrpSurface',
  hierarchy: 'chGrpHierarchy',
  statistic: 'chGrpStatistic',
  waterfall: 'chGrpWaterfall',
};

// Excel 2016 chart types (treemap, sunburst, histogram, Pareto, box & whisker,
// waterfall, funnel) are built and drawn, but the library's chartex writer
// produces parts Excel rejects with a repair prompt. Until it does not, they
// are left out of every gallery; existing chartex charts still render and edit.
const offered = (g: ChartGroup): boolean => CHART_GROUPS[g].every((c) => !isCxChoice(c));

/** Chart groups offered by Insert Chart and Change Chart Type. */
export const OFFERED_GROUPS: readonly ChartGroup[] = CHART_GROUP_IDS.filter(offered);

interface InsertMenu {
  readonly label: MessageKey;
  readonly tip: MessageKey;
  readonly groups: readonly ChartGroup[];
  readonly icon: string;
  /** Grid cell in the Charts group's three-row block of small buttons, 1-based. */
  readonly col: number;
  readonly row: number;
}

/** The Insert tab's chart buttons, each a menu over one or more groups, in Excel's grid positions. */
const ALL_INSERT_MENUS: readonly InsertMenu[] = [
  { label: 'chInsColumn', icon: 'chart-column', tip: 'chInsColumnTip', groups: ['column', 'bar'], col: 1, row: 1 },
  { label: 'chInsLine', icon: 'chart-line', tip: 'chInsLineTip', groups: ['line', 'area'], col: 1, row: 2 },
  { label: 'chInsPie', icon: 'chart-pie', tip: 'chInsPieTip', groups: ['pie'], col: 1, row: 3 },
  { label: 'chInsHierarchy', icon: 'chart', tip: 'chInsHierarchyTip', groups: ['hierarchy'], col: 2, row: 1 },
  { label: 'chInsStatistic', icon: 'chart', tip: 'chInsStatisticTip', groups: ['statistic'], col: 2, row: 2 },
  { label: 'chInsScatter', icon: 'chart-scatter', tip: 'chInsScatterTip', groups: ['scatter', 'bubble'], col: 2, row: 3 },
  { label: 'chInsWaterfall', icon: 'chart-stock', tip: 'chInsWaterfallTip', groups: ['waterfall', 'stock', 'surface', 'radar'], col: 3, row: 1 },
];

export const INSERT_MENUS: readonly InsertMenu[] = ALL_INSERT_MENUS.flatMap((m) => {
  const groups = m.groups.filter(offered);
  return groups.length > 0 ? [{ ...m, groups }] : [];
});

// Thumbnails draw the type over fixed sample data, like Excel's gallery icons.
// Four series suit every type: stock reads them as open/high/low/close and
// bubble as X/Y/size pairs. The sheet name matches no real sheet, so the chart
// is drawn from the caches the builder writes.
const SAMPLE: readonly CalcScalar[][] = [
  [null, 'A', 'B', 'C', 'D'],
  ['1', 30, 38, 24, 34],
  ['2', 34, 42, 30, 31],
  ['3', 26, 36, 20, 28],
  ['4', 31, 40, 27, 37],
  ['5', 22, 30, 18, 25],
];
const SAMPLE_RANGE = { r1: 1, c1: 1, r2: SAMPLE.length, c2: 5 };
const SAMPLE_SHEET = '⁣thumbnail';

const noCalc: ChartCalc = { sheetName: () => undefined, cellValue: () => null };

const thumbs = new Map<string, string>();

/** A small text-free picture of a chart type, cached per type, size and palette. */
export function chartThumbnail(choice: ChartChoice, width: number, height: number, palette: ThemePalette): string {
  const key = `${choice}|${width}|${height}|${palette.join()}`;
  const hit = thumbs.get(key);
  if (hit !== undefined) return hit;
  const chart = buildChart(choice, SAMPLE_SHEET, SAMPLE_RANGE, (r, c) => SAMPLE[r - 1]?.[c - 1] ?? null);
  const data = chartData(chart, noCalc, new Map(), palette);
  const svg = data
    ? shrinkSvg(renderChartSvg(
        { ...data, title: undefined, legend: undefined, catAxis: false, valAxis: false, catTitle: undefined, valTitle: undefined, chartLine: null, chartFill: null, series: data.series.map((s) => Object.assign({}, s, { name: '', labels: undefined })) },
        width * 3,
        height * 3,
        palette,
      ), width, height)
    : '';
  thumbs.set(key, svg);
  return svg;
}

/**
 * Show an SVG drawn at a larger size in a `width` × `height` box. Gallery
 * previews are drawn big and scaled down so margins, text and strokes shrink
 * with the chart, as in Excel's galleries.
 */
export function shrinkSvg(svg: string, width: number, height: number): string {
  return svg.replace(/^<svg ([^>]*?)width="\d+" height="\d+"/, `<svg $1width="${width}" height="${height}"`);
}
