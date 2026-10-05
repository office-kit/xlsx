// Chart Design / Format edits on a chart model. Each function mutates the
// chart in place; callers wrap it in one `doc.transact` (see `editChart`), so
// every ribbon click is a single undo step.
//
// Classic (`c:`) charts support everything here. Chartex (`cx:`) charts
// model only a title, a legend, data labels, gridlines and fills, so the other
// edits leave them unchanged.

import type { ChartSpace, ChartTitle, DataLabelList, DataLabelPosition, LegendPosition } from '@office-kit/xlsx/chart';
import type { ChartReference, DmlColorWithMods, Fill, ShapeProperties } from '@office-kit/xlsx/drawing';
import type { CalcScalar } from '../calc/index.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';
import { titleText } from './chart-data.ts';

/** Run `fn` on the chart at `index` of the active sheet's drawing as one undo step. */
export function editChart(doc: SpreadsheetEditor, index: number, label: string, fn: (chart: ChartReference) => void): void {
  const ws = doc.ws;
  const item = ws.drawing?.items[index];
  if (item?.content.kind !== 'chart') return;
  doc.transact(label, (tx) => {
    tx.sheet(ws, 'drawing');
    // The snapshot above cloned the drawing for undo; edit the live item.
    const live = ws.drawing?.items[index];
    if (live?.content.kind === 'chart') fn(live.content.chart);
  });
}

export type TitleMode = 'none' | 'above' | 'overlay';
export type LegendChoice = LegendPosition | 'none';
export type LabelChoice = 'none' | 'center' | 'insideEnd' | 'insideBase' | 'outsideEnd' | 'bestFit';
export type Gridline = 'valMajor' | 'valMinor' | 'catMajor';
export type Axis = 'cat' | 'val';

/** The current Add Chart Element settings, for the menus' check marks. */
export interface ChartElements {
  readonly title: TitleMode;
  readonly titleText: string;
  readonly catTitle: boolean;
  readonly valTitle: boolean;
  readonly legend: LegendChoice;
  readonly labels: LabelChoice;
  readonly gridlines: Readonly<Record<Gridline, boolean>>;
  readonly axes: Readonly<Record<Axis, boolean>>;
  /** Category/value axes exist (false for pies and chartex charts). */
  readonly hasAxes: boolean;
}

const LABEL_POS: Readonly<Record<Exclude<LabelChoice, 'none'>, DataLabelPosition>> = {
  center: 'ctr',
  insideEnd: 'inEnd',
  insideBase: 'inBase',
  outsideEnd: 'outEnd',
  bestFit: 'bestFit',
};

const on = (g: boolean | object | undefined): boolean => g !== undefined && g !== false;

function hasAxes(space: ChartSpace): boolean {
  return 'axIds' in space.plotArea.chart;
}

export function chartElements(chart: ChartReference): ChartElements {
  const space = chart.space;
  if (space) {
    const pa = space.plotArea;
    const dl = pa.chart.series[0]?.dLbls;
    const labelPos = dl && !dl.delete && (dl.showVal || dl.showPercent) ? (Object.entries(LABEL_POS).find(([, v]) => v === dl.dLblPos)?.[0] ?? 'bestFit') : 'none';
    return {
      title: space.title ? (space.title.overlay ? 'overlay' : 'above') : 'none',
      titleText: titleText(space.title) ?? '',
      catTitle: pa.catAx?.title !== undefined,
      valTitle: pa.valAx?.title !== undefined,
      legend: space.legend?.position ?? 'none',
      labels: isLabelChoice(labelPos) ? labelPos : 'none',
      gridlines: { valMajor: on(pa.valAx?.majorGridlines), valMinor: on(pa.valAx?.minorGridlines), catMajor: on(pa.catAx?.majorGridlines) },
      axes: { cat: pa.catAx?.delete !== true, val: pa.valAx?.delete !== true },
      hasAxes: hasAxes(space),
    };
  }
  const cx = chart.cxSpace;
  const val = cx?.chart.plotArea.axes.find((a) => a.id === 1);
  const legendPos = cx?.chart.legend?.pos;
  return {
    title: cx?.chart.title ? (cx.chart.title.overlay ? 'overlay' : 'above') : 'none',
    titleText: cx?.chart.title?.text ?? '',
    catTitle: false,
    valTitle: false,
    legend: legendPos ?? 'none',
    labels: cx?.chart.plotArea.series[0]?.dataLabels ? 'outsideEnd' : 'none',
    gridlines: { valMajor: val?.majorGridlines === true, valMinor: false, catMajor: false },
    axes: { cat: true, val: true },
    hasAxes: false,
  };
}

function isLabelChoice(v: string): v is LabelChoice {
  return v === 'none' || v in LABEL_POS;
}

export function setChartTitle(chart: ChartReference, mode: TitleMode, text?: string): void {
  const space = chart.space;
  if (space) {
    if (mode === 'none') {
      delete space.title;
      return;
    }
    const prev: ChartTitle = space.title ?? {};
    const value = text ?? titleText(prev) ?? '';
    const next: ChartTitle = { ...prev, overlay: mode === 'overlay' };
    // A typed title replaces rich text; `text` is what the serializer and renderer read first.
    delete next.tx;
    space.title = { ...next, text: value || 'Chart Title' };
    return;
  }
  const cx = chart.cxSpace;
  if (!cx) return;
  if (mode === 'none') delete cx.chart.title;
  else cx.chart.title = { ...cx.chart.title, overlay: mode === 'overlay', text: text ?? cx.chart.title?.text ?? 'Chart Title' };
}

export function setAxisTitle(chart: ChartReference, axis: Axis, show: boolean): void {
  const pa = chart.space?.plotArea;
  const ax = axis === 'cat' ? pa?.catAx : pa?.valAx;
  if (!ax) return;
  if (show) ax.title ??= { text: 'Axis Title' };
  else delete ax.title;
}

export function setLegend(chart: ChartReference, pos: LegendChoice): void {
  const space = chart.space;
  if (space) {
    if (pos === 'none') delete space.legend;
    else space.legend = { ...space.legend, position: pos, overlay: false };
    return;
  }
  const cx = chart.cxSpace;
  if (!cx) return;
  if (pos === 'none') delete cx.chart.legend;
  else cx.chart.legend = { ...cx.chart.legend, pos: pos === 'tr' ? 'r' : pos, overlay: false };
}

export function setDataLabels(chart: ChartReference, choice: LabelChoice): void {
  const space = chart.space;
  if (space) {
    const pie = ['pie', 'pie3D', 'doughnut', 'ofPie'].includes(space.plotArea.chart.kind);
    for (const s of space.plotArea.chart.series) {
      if (choice === 'none') {
        delete s.dLbls;
        continue;
      }
      const list: DataLabelList = { showLegendKey: false, showVal: true, showCatName: false, showSerName: false, showPercent: false, showBubbleSize: false };
      // Doughnuts reject dLblPos; Excel positions their labels itself.
      if (space.plotArea.chart.kind !== 'doughnut') list.dLblPos = pie && choice === 'insideBase' ? 'inEnd' : LABEL_POS[choice];
      s.dLbls = list;
    }
    return;
  }
  for (const s of chart.cxSpace?.chart.plotArea.series ?? []) {
    if (choice === 'none') delete s.dataLabels;
    else s.dataLabels = { pos: choice === 'center' ? 'ctr' : choice === 'insideEnd' ? 'inEnd' : choice === 'insideBase' ? 'inBase' : 'outEnd', visibility: { value: true, seriesName: false, categoryName: false } };
  }
}

export function setGridline(chart: ChartReference, which: Gridline, show: boolean): void {
  const pa = chart.space?.plotArea;
  if (pa) {
    const ax = which === 'catMajor' ? pa.catAx : pa.valAx;
    if (!ax) return;
    if (which === 'valMinor') ax.minorGridlines = show;
    else ax.majorGridlines = show;
    return;
  }
  const val = chart.cxSpace?.chart.plotArea.axes.find((a) => a.id === 1);
  if (val && which === 'valMajor') val.majorGridlines = show;
}

export function setAxisVisible(chart: ChartReference, axis: Axis, show: boolean): void {
  const pa = chart.space?.plotArea;
  const ax = axis === 'cat' ? pa?.catAx : pa?.valAx;
  if (ax) ax.delete = !show;
}

// ---- Quick Layout ------------------------------------------------------------------------

export const QUICK_LAYOUTS = [1, 2, 3, 4, 5, 6] as const;
export type QuickLayout = (typeof QUICK_LAYOUTS)[number];

interface LayoutPreset {
  readonly legend: LegendChoice;
  readonly labels: LabelChoice;
  readonly catTitle: boolean;
  readonly valTitle: boolean;
  readonly grid: boolean;
}

/** A few of Excel's Quick Layout presets: combinations of title, legend, labels, axis titles and gridlines. */
const QUICK_LAYOUT_PRESETS: Readonly<Record<QuickLayout, LayoutPreset>> = {
  1: { legend: 'r', labels: 'none', catTitle: false, valTitle: false, grid: true },
  2: { legend: 't', labels: 'outsideEnd', catTitle: false, valTitle: false, grid: false },
  3: { legend: 'b', labels: 'none', catTitle: false, valTitle: false, grid: true },
  4: { legend: 'b', labels: 'outsideEnd', catTitle: false, valTitle: false, grid: false },
  5: { legend: 'none', labels: 'none', catTitle: true, valTitle: true, grid: true },
  6: { legend: 'none', labels: 'none', catTitle: false, valTitle: true, grid: true },
};

export function applyQuickLayout(chart: ChartReference, layout: QuickLayout): void {
  const preset = QUICK_LAYOUT_PRESETS[layout];
  setChartTitle(chart, 'above');
  setLegend(chart, preset.legend);
  setDataLabels(chart, preset.labels);
  setAxisTitle(chart, 'cat', preset.catTitle);
  setAxisTitle(chart, 'val', preset.valTitle);
  setGridline(chart, 'valMajor', preset.grid);
  setGridline(chart, 'valMinor', false);
}

// ---- Change Colors -----------------------------------------------------------------------

type Shade = readonly [lumMod: number, lumOff: number];

/** Shades Excel's monochromatic palettes step through, as lumMod / lumOff percentages. */
const MONO_SHADES: readonly Shade[] = [
  [100, 0],
  [60, 40],
  [75, 0],
  [40, 60],
  [50, 0],
  [80, 20],
];

const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const;

export interface ColorSet {
  readonly id: string;
  readonly colors: readonly DmlColorWithMods[];
}

const scheme = (name: (typeof ACCENTS)[number], shade: Shade = [100, 0]): DmlColorWithMods => ({
  base: { kind: 'schemeClr', value: name },
  mods: [...(shade[0] !== 100 ? [{ kind: 'lumMod' as const, val: shade[0] * 1000 }] : []), ...(shade[1] !== 0 ? [{ kind: 'lumOff' as const, val: shade[1] * 1000 }] : [])],
});

const rotate = <T>(list: readonly T[], by: number): T[] => [...list.slice(by), ...list.slice(0, by)];

/** Excel's Change Colors gallery: four colourful sets, then one monochromatic set per accent. */
export const COLOR_SETS: readonly ColorSet[] = [
  { id: 'colorful1', colors: ACCENTS.map((a) => scheme(a)) },
  { id: 'colorful2', colors: rotate(ACCENTS, 1).map((a) => scheme(a)) },
  { id: 'colorful3', colors: rotate(ACCENTS, 2).map((a) => scheme(a)) },
  { id: 'colorful4', colors: rotate(ACCENTS, 3).map((a) => scheme(a)) },
  ...ACCENTS.map((a, i) => ({ id: `mono${i + 1}`, colors: MONO_SHADES.map((s) => scheme(a, s)) })),
];

const solid = (color: DmlColorWithMods): Fill => ({ kind: 'solidFill', color });

function lineish(space: ChartSpace): boolean {
  const k = space.plotArea.chart.kind;
  return k === 'line' || k === 'line3D' || k === 'scatter' || (k === 'radar' && space.plotArea.chart.radarStyle !== 'filled');
}

export function applyColorSet(chart: ChartReference, set: ColorSet): void {
  const color = (i: number): DmlColorWithMods => set.colors[i % set.colors.length] ?? scheme('accent1');
  const space = chart.space;
  if (space) {
    const c = space.plotArea.chart;
    const byPoint = 'varyColors' in c && c.varyColors === true && c.series.length === 1;
    c.series.forEach((s, i) => {
      if (byPoint) {
        const n = 'val' in s ? (s.val.cache?.length ?? 6) : 6;
        s.dPt = Array.from({ length: n }, (_, idx) => ({ idx, spPr: { fill: solid(color(idx)) } }));
        return;
      }
      const spPr: ShapeProperties = { ...s.spPr };
      if (lineish(space)) {
        // A hidden series line (markers-only scatter) stays hidden.
        if (spPr.ln?.fill?.kind !== 'noFill') spPr.ln = { ...spPr.ln, fill: solid(color(i)), w: spPr.ln?.w ?? 28_575 };
      } else spPr.fill = solid(color(i));
      s.spPr = spPr;
      delete s.dPt;
    });
    return;
  }
  chart.cxSpace?.chart.plotArea.series.forEach((s, i) => {
    s.spPr = { ...s.spPr, fill: solid(color(i)) };
  });
}

// ---- Chart Styles ---------------------------------------------------------------------------

export const CHART_STYLES = [1, 2, 3, 4, 5, 6] as const;
export type ChartStyle = (typeof CHART_STYLES)[number];

const srgb = (hex: string): Fill => solid({ base: { kind: 'srgb', value: hex }, mods: [] });

/**
 * A handful of looks in the spirit of Excel's Chart Styles gallery, built only
 * from properties that round-trip in the file: frame and plot fills,
 * gridlines, gap width and data labels.
 */
export function applyChartStyle(chart: ChartReference, style: ChartStyle): void {
  const frame: Record<ChartStyle, ShapeProperties> = {
    1: { fill: srgb('FFFFFF'), ln: { fill: srgb('D9D9D9'), w: 9525 } },
    2: { fill: srgb('FFFFFF'), ln: { fill: srgb('D9D9D9'), w: 9525 } },
    3: { fill: srgb('F2F2F2'), ln: { fill: { kind: 'noFill' } } },
    4: { fill: srgb('404040'), ln: { fill: { kind: 'noFill' } } },
    5: { fill: srgb('FFFFFF'), ln: { fill: { kind: 'noFill' } } },
    6: { fill: srgb('DEEBF7'), ln: { fill: srgb('9DC3E6'), w: 9525 } },
  };
  const space = chart.space;
  if (space) {
    space.spPr = frame[style];
    delete space.plotArea.spPr;
    const c = space.plotArea.chart;
    if (c.kind === 'bar' || c.kind === 'bar3D') c.gapWidth = style === 5 ? 50 : style === 2 ? 100 : 219;
  } else if (chart.cxSpace) chart.cxSpace.spPr = frame[style];
  setGridline(chart, 'valMajor', style !== 2 && style !== 5);
  setDataLabels(chart, style === 2 || style === 5 ? 'outsideEnd' : style === 6 ? 'insideEnd' : 'none');
}

// ---- Format: fill and outline -----------------------------------------------------------------

/** Chart area fill: `#RRGGBB`, or null for No Fill. */
export function setChartFill(chart: ChartReference, hex: string | null): void {
  const fill: Fill = hex === null ? { kind: 'noFill' } : srgb(hex.replace('#', '').toUpperCase());
  if (chart.space) chart.space.spPr = { ...chart.space.spPr, fill };
  else if (chart.cxSpace) chart.cxSpace.spPr = { ...chart.cxSpace.spPr, fill };
}

/** Chart area outline: colour or null for No Line, width in points. */
export function setChartOutline(chart: ChartReference, hex: string | null, widthPt = 0.75): void {
  const ln = hex === null ? { fill: { kind: 'noFill' as const } } : { fill: srgb(hex.replace('#', '').toUpperCase()), w: Math.round(widthPt * 12_700) };
  if (chart.space) chart.space.spPr = { ...chart.space.spPr, ln };
  else if (chart.cxSpace) chart.cxSpace.spPr = { ...chart.cxSpace.spPr, ln };
}

// ---- Select Data: series refs -------------------------------------------------------------

/** One series as Select Data shows it: refs as text, `name` either a ref or literal text. */
export interface SeriesRefs {
  name: string;
  values: string;
  /** X values (scatter, bubble). */
  x: string;
  /** Bubble sizes. */
  size: string;
}

/** Whether the chart's series take X values instead of shared categories. */
export function isXYChart(chart: ChartReference): boolean {
  const k = chart.space?.plotArea.chart.kind;
  return k === 'scatter' || k === 'bubble';
}

export function seriesRefs(chart: ChartReference): { series: SeriesRefs[]; categories: string } {
  const c = chart.space?.plotArea.chart;
  if (!c) return { series: [], categories: '' };
  const series: SeriesRefs[] = [];
  let categories = '';
  for (const s of c.series) {
    const name = s.tx?.kind === 'ref' ? `=${s.tx.ref}` : (s.tx?.value ?? '');
    if ('yVal' in s) series.push({ name, values: s.yVal.ref, x: s.xVal?.ref ?? '', size: 'bubbleSize' in s ? s.bubbleSize.ref : '' });
    else {
      series.push({ name, values: s.val.ref, x: '', size: '' });
      categories ||= s.cat?.ref ?? '';
    }
  }
  return { series, categories };
}

/** Grow or shrink `list` to `n` series, new ones copied from the last, and renumber them. */
function resizeSeries<T extends { idx: number; order: number; dPt?: unknown; dLbls?: unknown }>(list: T[], n: number): void {
  const template = list[list.length - 1];
  list.length = Math.min(list.length, n);
  if (template) {
    for (let i = list.length; i < n; i++) {
      const copy = structuredClone(template);
      delete copy.dPt;
      list.push(copy);
    }
  }
  list.forEach((s, i) => {
    s.idx = i;
    s.order = i;
  });
}

/**
 * Select Data's OK for a classic chart: point each series at the given refs
 * and refill the caches from `read`, so the saved file shows the new data
 * before Excel recalculates. Refs must already be validated.
 */
export function setSeriesRefs(chart: ChartReference, rows: readonly SeriesRefs[], categories: string, read: (ref: string) => readonly CalcScalar[] | undefined): void {
  const c = chart.space?.plotArea.chart;
  if (!c || rows.length === 0) return;
  const nums = (ref: string) => ({ ref, cache: (read(ref) ?? []).map((v) => (typeof v === 'number' ? v : 0)) });
  const tx = (name: string): { kind: 'literal'; value: string } | { kind: 'ref'; ref: string } | undefined =>
    name.startsWith('=') ? { kind: 'ref', ref: name.slice(1) } : name ? { kind: 'literal', value: name } : undefined;
  resizeSeries<(typeof c.series)[number]>(c.series, rows.length);
  c.series.forEach((s, i) => {
    const row = rows[i];
    if (!row) return;
    const name = tx(row.name);
    if (name) s.tx = name;
    else delete s.tx;
    if ('yVal' in s) {
      s.yVal = nums(row.values);
      if (row.x) s.xVal = nums(row.x);
      else delete s.xVal;
      if ('bubbleSize' in s) s.bubbleSize = nums(row.size || row.values);
      return;
    }
    s.val = nums(row.values);
    if (categories) s.cat = { ref: categories, cacheKind: 'str', cache: (read(categories) ?? []).map((v) => (v === null || typeof v === 'object' ? '' : String(v))) };
    else delete s.cat;
  });
}
