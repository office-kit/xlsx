// Reading a chart model into plain numbers and colours for the SVG renderer.
//
// Series refs (`Sheet1!$B$2:$B$9`) are resolved live through the formula
// engine so a chart follows edits; the caches written into the file are only
// a fallback for refs the editor cannot resolve (other workbooks, unions).

import type { BarSeries, ChartSpace, ChartTitle, DataLabelList, DataLabelPosition, DataPoint, LegendPosition, LineSeries } from '@office-kit/xlsx/chart';
import type { CxChartSpace, CxDataLabels } from '@office-kit/xlsx/chart';
import type { ChartReference, DmlColorWithMods, Fill, LineProperties, ShapeProperties, TextBody } from '@office-kit/xlsx/drawing';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { coerceToText, type CalcEngine, type CalcScalar } from '../calc/index.ts';
import { parseRangeAddress } from './address.ts';
import { applyTint, type ThemePalette } from './theme.ts';

export type RenderKind =
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'bubble'
  | 'radar'
  | 'stock'
  | 'surface'
  | 'waterfall'
  | 'funnel'
  | 'treemap'
  | 'sunburst'
  | 'histogram'
  | 'pareto'
  | 'boxWhisker';

/** Which parts of a data label are shown, and where. */
export interface LabelSpec {
  readonly pos: DataLabelPosition | undefined;
  readonly value: boolean;
  readonly percent: boolean;
  readonly category: boolean;
  readonly series: boolean;
}

export interface RenderSeries {
  readonly name: string;
  readonly values: ReadonlyArray<number | null>;
  /** Scatter / bubble X values. */
  readonly xs?: ReadonlyArray<number | null> | undefined;
  /** Bubble sizes. */
  readonly sizes?: ReadonlyArray<number | null> | undefined;
  readonly color: string;
  /** Per-point colour overrides (`c:dPt`), by point index. */
  readonly pointColors?: ReadonlyArray<string | undefined> | undefined;
  readonly markers: boolean;
  /** Connect the points (line, radar, scatter); false when the series line is set to No line. */
  readonly line: boolean;
  readonly smooth: boolean;
  readonly labels: LabelSpec | undefined;
}

export interface Stroke {
  readonly color: string;
  readonly width: number;
}

export interface ChartData {
  readonly kind: RenderKind;
  readonly horizontal: boolean;
  readonly grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  readonly categories: readonly string[];
  readonly series: readonly RenderSeries[];
  readonly title: string | undefined;
  readonly legend: LegendPosition | undefined;
  readonly varyColors: boolean;
  readonly holeSize: number;
  readonly gapWidth: number;
  readonly overlap: number;
  readonly valMin: number | undefined;
  readonly valMax: number | undefined;
  readonly catTitle: string | undefined;
  readonly valTitle: string | undefined;
  readonly catAxis: boolean;
  readonly valAxis: boolean;
  readonly majorGridlines: boolean;
  readonly minorGridlines: boolean;
  readonly catGridlines: boolean;
  /** Chart area fill; null when set to No fill. */
  readonly chartFill: string | null;
  readonly chartLine: Stroke | null;
  readonly plotFill: string | null;
  readonly radarFilled: boolean;
  readonly upDownBars: boolean;
  /** Waterfall points drawn as totals from the axis rather than as steps. */
  readonly subtotals: readonly number[];
}

/** Office theme order Excel cycles series colours through: accent1–6, then darker, then lighter variants. */
export function seriesColor(i: number, palette: ThemePalette): string {
  const base = palette[4 + (i % 6)] ?? '4472C4';
  const round = Math.floor(i / 6);
  const tint = round === 0 ? 0 : round % 2 === 1 ? -0.25 * Math.ceil(round / 2) : 0.4 * (round / 2);
  return `#${applyTint(base, tint)}`;
}

const SCHEME_INDEX: Readonly<Record<string, number>> = {
  lt1: 0,
  bg1: 0,
  dk1: 1,
  tx1: 1,
  lt2: 2,
  bg2: 2,
  dk2: 3,
  tx2: 3,
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9,
  hlink: 10,
  folHlink: 11,
};

export function dmlColor(c: DmlColorWithMods, palette: ThemePalette): string | undefined {
  let hex: string | undefined;
  const base = c.base;
  if (base.kind === 'srgb') hex = base.value;
  else if (base.kind === 'schemeClr') hex = palette[SCHEME_INDEX[base.value] ?? 4];
  else if (base.kind === 'sysClr') hex = base.lastClr ?? (base.value === 'window' ? 'FFFFFF' : '000000');
  if (!hex) return undefined;
  // lumMod/lumOff are how theme shades are written; approximate them as a tint.
  let lumMod = 1;
  let lumOff = 0;
  for (const m of c.mods) {
    if (m.kind === 'lumMod') lumMod = m.val / 100_000;
    if (m.kind === 'lumOff') lumOff = m.val / 100_000;
  }
  const tint = lumOff > 0 ? lumOff : lumMod < 1 ? lumMod - 1 : 0;
  return `#${applyTint(hex.toUpperCase(), tint)}`;
}

/** A fill's colour: a hex string, null for No fill, undefined when unset or not a flat colour. */
export function fillColor(fill: Fill | undefined, palette: ThemePalette): string | null | undefined {
  if (fill?.kind === 'noFill') return null;
  if (fill?.kind === 'solidFill') return dmlColor(fill.color, palette);
  if (fill?.kind === 'gradFill') {
    const stop = fill.stops[0];
    return stop ? dmlColor(stop.color, palette) : undefined;
  }
  return undefined;
}

const EMU_PER_PT = 12_700;

function stroke(ln: LineProperties | undefined, palette: ThemePalette, fallback: Stroke | null): Stroke | null {
  if (!ln) return fallback;
  const color = fillColor(ln.fill, palette);
  if (color === null) return null;
  return { color: color ?? fallback?.color ?? '#D9D9D9', width: ln.w !== undefined ? (ln.w / EMU_PER_PT) * (96 / 72) : (fallback?.width ?? 1) };
}

function bodyText(body: TextBody | undefined): string {
  if (!body) return '';
  return body.paragraphs.map((p) => p.runs.map((r) => (r.kind === 'r' ? r.t : r.kind === 'fld' ? (r.t ?? '') : '\n')).join('')).join('\n');
}

export function titleText(title: ChartTitle | undefined): string | undefined {
  if (!title) return undefined;
  return title.text ?? bodyText(title.tx);
}

/** The part of the calc engine a chart reads its values through. */
export type ChartCalc = Pick<CalcEngine, 'sheetName' | 'cellValue'>;

/** Largest number of points read from one ref; guards whole-column refs. */
const MAX_POINTS = 10_000;

/** Values of a ref like `'My Sheet'!$B$2:$B$9`, or undefined when it is not a single area on a known sheet. */
export function resolveRef(ref: string, calc: ChartCalc, worksheets: ReadonlyMap<string, Worksheet>): CalcScalar[] | undefined {
  const parsed = parseRangeAddress(ref.replace(/^\(|\)$/g, ''));
  if (!parsed?.sheet) return undefined;
  const sheet = calc.sheetName(parsed.sheet);
  if (sheet === undefined) return undefined;
  const ws = worksheets.get(sheet);
  if (!ws) return undefined;
  const { r1, c1, c2 } = parsed.range;
  let lastRow = 0;
  for (const r of ws.rows.keys()) if (r > lastRow) lastRow = r;
  const r2 = Math.min(parsed.range.r2, Math.max(r1, lastRow));
  const out: CalcScalar[] = [];
  for (let r = r1; r <= r2 && out.length < MAX_POINTS; r++) {
    for (let c = c1; c <= c2 && out.length < MAX_POINTS; c++) out.push(calc.cellValue(sheet, r, c));
  }
  return out;
}

const toNum = (v: CalcScalar): number | null => (typeof v === 'number' ? v : null);
const toText = (v: CalcScalar): string => {
  if (v === null) return '';
  const t = coerceToText(v);
  return typeof t === 'string' ? t : t.code;
};

function labelSpec(d: DataLabelList | undefined): LabelSpec | undefined {
  if (!d || d.delete) return undefined;
  const spec = { pos: d.dLblPos, value: d.showVal === true, percent: d.showPercent === true, category: d.showCatName === true, series: d.showSerName === true };
  return spec.value || spec.percent || spec.category || spec.series ? spec : undefined;
}

function cxLabelSpec(d: CxDataLabels | undefined): LabelSpec | undefined {
  if (!d) return undefined;
  const v = d.visibility;
  return { pos: d.pos, value: v?.value !== false, percent: false, category: v?.categoryName === true, series: v?.seriesName === true };
}

function pointColors(points: readonly DataPoint[] | undefined, palette: ThemePalette): Array<string | undefined> | undefined {
  if (!points || points.length === 0) return undefined;
  const out: Array<string | undefined> = [];
  for (const p of points) out[p.idx] = fillColor(p.spPr?.fill, palette) ?? undefined;
  return out;
}

/** Excel's default chart frame: white with a light grey outline. */
const DEFAULT_LINE: Stroke = { color: '#D9D9D9', width: 1 };

function frame(spPr: ShapeProperties | undefined, palette: ThemePalette): Pick<ChartData, 'chartFill' | 'chartLine'> {
  const fill = fillColor(spPr?.fill, palette);
  return { chartFill: fill === undefined ? '#FFFFFF' : fill, chartLine: stroke(spPr?.ln, palette, DEFAULT_LINE) };
}

/** Everything the renderer needs, with refs resolved; undefined for chart kinds it does not draw. */
export function chartData(chart: ChartReference, calc: ChartCalc, worksheets: ReadonlyMap<string, Worksheet>, palette: ThemePalette): ChartData | undefined {
  if (chart.space) return classicData(chart.space, calc, worksheets, palette);
  if (chart.cxSpace) return cxData(chart.cxSpace, calc, worksheets, palette);
  return undefined;
}

function classicData(space: ChartSpace, calc: ChartCalc, worksheets: ReadonlyMap<string, Worksheet>, palette: ThemePalette): ChartData | undefined {
  const chart = space.plotArea.chart;
  const resolveNums = (ref: { ref: string; cache?: number[] } | undefined): Array<number | null> => {
    if (!ref) return [];
    const live = resolveRef(ref.ref, calc, worksheets);
    return live ? live.map(toNum) : (ref.cache ?? []);
  };
  const resolveTexts = (ref: { ref: string; cache?: ReadonlyArray<string | number> } | undefined): string[] => {
    if (!ref) return [];
    const live = resolveRef(ref.ref, calc, worksheets);
    return live ? live.map(toText) : (ref.cache ?? []).map(String);
  };
  const nameOf = (tx: BarSeries['tx'], i: number): string => {
    if (tx?.kind === 'literal') return tx.value;
    if (tx?.kind === 'ref') {
      const live = resolveRef(tx.ref, calc, worksheets);
      if (live) return live.map(toText).join(' ');
    }
    return `Series${i + 1}`;
  };
  const varyColors = 'varyColors' in chart && chart.varyColors === true;
  const { catAx, valAx } = space.plotArea;
  const gridOn = (g: boolean | object | undefined): boolean => g !== undefined && g !== false;
  const plotFill = fillColor(space.plotArea.spPr?.fill, palette);
  const common = {
    title: titleText(space.title),
    legend: space.legend?.position,
    varyColors,
    holeSize: chart.kind === 'doughnut' ? (chart.holeSize ?? 50) : 0,
    gapWidth: 'gapWidth' in chart && chart.gapWidth !== undefined ? chart.gapWidth : 150,
    overlap: 'overlap' in chart && chart.overlap !== undefined ? chart.overlap : 0,
    valMin: valAx?.scaling?.min,
    valMax: valAx?.scaling?.max,
    catTitle: titleText(catAx?.title),
    valTitle: titleText(valAx?.title),
    catAxis: catAx?.delete !== true,
    valAxis: valAx?.delete !== true,
    majorGridlines: valAx ? gridOn(valAx.majorGridlines) : true,
    minorGridlines: gridOn(valAx?.minorGridlines),
    catGridlines: gridOn(catAx?.majorGridlines),
    ...frame(space.spPr, palette),
    plotFill: plotFill ?? null,
    radarFilled: chart.kind === 'radar' && chart.radarStyle === 'filled',
    upDownBars: chart.kind === 'stock' && chart.upDownBars !== undefined && chart.upDownBars !== false,
    subtotals: [],
  };
  const radarMarkers = chart.kind === 'radar' && chart.radarStyle === 'marker';
  const barLike = (kind: RenderKind, series: readonly LineSeries[], horizontal: boolean, grouping: ChartData['grouping'], lineish: boolean, chartLabels?: DataLabelList): ChartData => {
    const sorted = [...series].sort((a, b) => a.order - b.order);
    const first = sorted[0];
    const categories = first?.cat ? resolveTexts(first.cat) : [];
    return {
      kind,
      horizontal,
      grouping,
      categories,
      series: sorted.map((s, i) => {
        const colors = pointColors(s.dPt, palette);
        return {
          name: nameOf(s.tx, s.idx),
          values: resolveNums(s.val),
          color: (lineish ? fillColor(s.spPr?.ln?.fill, palette) : fillColor(s.spPr?.fill, palette)) ?? seriesColor(i, palette),
          pointColors: colors,
          // Radar series carry no marker of their own; the chart's style decides.
          markers: kind === 'radar' ? radarMarkers : lineish && s.marker?.symbol !== 'none',
          line: lineish && s.spPr?.ln?.fill?.kind !== 'noFill',
          smooth: s.smooth === true,
          labels: labelSpec(s.dLbls ?? chartLabels),
        };
      }),
      ...common,
    };
  };
  switch (chart.kind) {
    case 'bar':
    case 'bar3D':
      return barLike('bar', chart.series, chart.barDir === 'bar', chart.grouping, false);
    case 'line':
    case 'line3D':
      return barLike('line', chart.series, false, chart.grouping, true);
    case 'area':
    case 'area3D':
      return barLike('area', chart.series, false, chart.grouping, false);
    case 'pie':
    case 'pie3D':
    case 'ofPie':
      return barLike('pie', chart.series, false, 'standard', false);
    case 'doughnut':
      return barLike('doughnut', chart.series, false, 'standard', false);
    case 'radar':
      return barLike('radar', chart.series, false, 'standard', chart.radarStyle !== 'filled');
    case 'stock':
      return barLike('stock', chart.series, false, 'standard', false);
    case 'surface':
    case 'surface3D':
      return barLike('surface', chart.series, false, 'standard', false);
    case 'scatter':
    case 'bubble': {
      const scatterStyle = chart.kind === 'scatter' ? chart.scatterStyle : 'marker';
      return {
        kind: chart.kind,
        horizontal: false,
        grouping: 'standard',
        categories: [],
        series: [...chart.series]
          .sort((a, b) => a.order - b.order)
          .map((s, i) => {
            const ys = resolveNums(s.yVal);
            const rawXs = s.xVal ? resolveNums(s.xVal) : [];
            // Text X values plot at 1, 2, 3…, as in Excel.
            const xs = rawXs.some((x) => x !== null) ? rawXs : ys.map((_, j) => j + 1);
            const marker = 'marker' in s ? s.marker : undefined;
            const colors = pointColors(s.dPt, palette);
            return {
              name: nameOf(s.tx, s.idx),
              values: ys,
              xs,
              sizes: 'bubbleSize' in s ? resolveNums(s.bubbleSize) : undefined,
              color:
                (chart.kind === 'bubble' ? fillColor(s.spPr?.fill, palette) : (fillColor(s.spPr?.ln?.fill, palette) ?? fillColor(marker?.spPr?.fill, palette))) ??
                seriesColor(i, palette),
              pointColors: colors,
              markers: chart.kind === 'bubble' || (marker?.symbol !== 'none' && scatterStyle !== 'line' && scatterStyle !== 'smooth'),
              line: chart.kind === 'scatter' && s.spPr?.ln?.fill?.kind !== 'noFill' && scatterStyle !== 'marker' && scatterStyle !== 'none',
              smooth: ('smooth' in s && s.smooth === true) || scatterStyle.startsWith('smooth'),
              labels: labelSpec(s.dLbls),
            };
          }),
        ...common,
      };
    }
  }
}

const CX_KIND: Readonly<Record<string, RenderKind>> = {
  waterfall: 'waterfall',
  funnel: 'funnel',
  treemap: 'treemap',
  sunburst: 'sunburst',
  boxWhisker: 'boxWhisker',
};

function cxData(space: CxChartSpace, calc: ChartCalc, worksheets: ReadonlyMap<string, Worksheet>, palette: ThemePalette): ChartData | undefined {
  const plot = space.chart.plotArea;
  const main = plot.series.find((s) => s.layoutId !== 'paretoLine');
  if (!main) return undefined;
  const kind: RenderKind | undefined = main.layoutId === 'clusteredColumn' ? (plot.series.some((s) => s.layoutId === 'paretoLine') ? 'pareto' : 'histogram') : CX_KIND[main.layoutId];
  if (!kind) return undefined;
  const data = space.chartData.data.find((d) => d.id === (main.dataId ?? 0));
  let categories: string[] = [];
  let values: Array<number | null> = [];
  for (const dim of data?.dims ?? []) {
    const live = dim.f ? resolveRef(dim.f, calc, worksheets) : undefined;
    if (dim.kind === 'str' && dim.type === 'cat') categories = live ? live.map(toText) : dim.pts.map((p) => p.v);
    if (dim.kind === 'num' && dim.type === 'val') values = live ? live.map(toNum) : dim.pts.map((p) => (Number.isFinite(Number(p.v)) ? Number(p.v) : null));
  }
  const name = main.tx?.f ? (resolveRef(main.tx.f, calc, worksheets)?.map(toText).join(' ') ?? main.tx.v ?? '') : (main.tx?.v ?? 'Series1');
  const valAxis = plot.axes.find((a) => a.id === 1);
  const subtotals = main.layoutPr?.kind === 'waterfall' ? main.layoutPr.subtotalIdx : [];
  const legendPos = space.chart.legend?.pos;
  return {
    kind,
    horizontal: false,
    grouping: 'standard',
    categories,
    series: [
      {
        name,
        values,
        color: fillColor(main.spPr?.fill, palette) ?? seriesColor(0, palette),
        markers: false,
        line: false,
        smooth: false,
        labels: cxLabelSpec(main.dataLabels),
      },
    ],
    title: space.chart.title?.text,
    legend: legendPos,
    varyColors: kind === 'treemap' || kind === 'sunburst',
    holeSize: 0,
    gapWidth: kind === 'histogram' || kind === 'pareto' ? 0 : 50,
    overlap: 0,
    valMin: valAxis?.valScaling?.min,
    valMax: valAxis?.valScaling?.max,
    catTitle: plot.axes.find((a) => a.id === 0)?.title?.text,
    valTitle: valAxis?.title?.text,
    catAxis: plot.axes.find((a) => a.id === 0)?.hidden !== true,
    valAxis: valAxis?.hidden !== true,
    majorGridlines: valAxis ? valAxis.majorGridlines === true : kind !== 'funnel' && kind !== 'treemap' && kind !== 'sunburst',
    minorGridlines: false,
    catGridlines: false,
    ...frame(space.spPr, palette),
    plotFill: fillColor(plot.spPr?.fill, palette) ?? null,
    radarFilled: false,
    upDownBars: false,
    subtotals,
  };
}
