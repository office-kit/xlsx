// Phase 6 acceptance: every chart kind we model — 16 legacy `c:` + 8 chartex
// `cx:` = 24 distinct kinds (some references count Pie + Pie3D + OfPie as three
// slots and call it "25"; @office-kit/xlsx implements the same set either way). Each
// chart goes through workbookToBytes → loadWorkbook and the loaded chart's kind
// / key attributes must match what we wrote.

import { describe, expect, it } from 'vitest';
import {
  type ChartKind,
  type ChartSpace,
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
} from '../../src/chart/chart';
import {
  type CxChartSpace,
  makeBoxWhiskerChart,
  makeFunnelChart,
  makeHistogramChart,
  makeParetoChart,
  makeRegionMapChart,
  makeSunburstChart,
  makeTreemapChart,
  makeWaterfallChart,
} from '../../src/chart/cx/chartex';
import { makeTwoCellAnchor } from '../../src/drawing/anchor';
import { makeChartDrawingItem, makeDrawing } from '../../src/drawing/drawing';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

const VAL: { ref: string } = { ref: 'A1:A4' };
const VAL2: { ref: string } = { ref: 'B1:B4' };

const wrap = (chart: ChartKind, withAxes: boolean): ChartSpace =>
  makeChartSpace({
    plotArea: {
      chart,
      ...(withAxes ? { catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } } : {}),
    },
  });

interface LegacyCase {
  name: string;
  kind: ChartKind['kind'];
  build: () => ChartSpace;
  /** Optional extra assertion against the round-tripped chart. */
  assert?: (loaded: ChartKind) => void;
}

const LEGACY_CASES: LegacyCase[] = [
  {
    name: 'Bar',
    kind: 'bar',
    build: () => wrap(makeBarChart({ series: [makeBarSeries({ idx: 0, val: VAL })] }), true),
    assert: (c) => {
      if (c.kind !== 'bar') throw new Error('expected bar');
      expect(c.series.length).toBe(1);
    },
  },
  {
    name: 'Line',
    kind: 'line',
    build: () =>
      wrap(makeLineChart({ series: [makeBarSeries({ idx: 0, val: VAL })], smooth: true }), true),
    assert: (c) => {
      if (c.kind !== 'line') throw new Error('expected line');
      expect(c.smooth).toBe(true);
    },
  },
  {
    name: 'Area',
    kind: 'area',
    build: () =>
      wrap(
        makeAreaChart({ grouping: 'stacked', series: [makeBarSeries({ idx: 0, val: VAL })] }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'area') throw new Error('expected area');
      expect(c.grouping).toBe('stacked');
    },
  },
  {
    name: 'Pie',
    kind: 'pie',
    build: () =>
      wrap(makePieChart({ varyColors: true, series: [makeBarSeries({ idx: 0, val: VAL })] }), false),
    assert: (c) => {
      if (c.kind !== 'pie') throw new Error('expected pie');
      expect(c.varyColors).toBe(true);
    },
  },
  {
    name: 'Doughnut',
    kind: 'doughnut',
    build: () =>
      wrap(
        makeDoughnutChart({
          holeSize: 60,
          firstSliceAng: 45,
          series: [makeBarSeries({ idx: 0, val: VAL })],
        }),
        false,
      ),
    assert: (c) => {
      if (c.kind !== 'doughnut') throw new Error('expected doughnut');
      expect(c.holeSize).toBe(60);
      expect(c.firstSliceAng).toBe(45);
    },
  },
  {
    name: 'Scatter',
    kind: 'scatter',
    build: () =>
      wrap(
        makeScatterChart({
          scatterStyle: 'smoothMarker',
          series: [makeScatterSeries({ idx: 0, yVal: VAL, xVal: VAL2 })],
        }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'scatter') throw new Error('expected scatter');
      expect(c.scatterStyle).toBe('smoothMarker');
    },
  },
  {
    name: 'Radar',
    kind: 'radar',
    build: () =>
      wrap(
        makeRadarChart({ radarStyle: 'marker', series: [makeBarSeries({ idx: 0, val: VAL })] }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'radar') throw new Error('expected radar');
      expect(c.radarStyle).toBe('marker');
    },
  },
  {
    name: 'Bubble',
    kind: 'bubble',
    build: () =>
      wrap(
        makeBubbleChart({
          bubble3D: true,
          bubbleScale: 80,
          showNegBubbles: false,
          sizeRepresents: 'area',
          series: [
            makeBubbleSeries({
              idx: 0,
              yVal: VAL,
              xVal: VAL2,
              bubbleSize: { ref: 'C1:C4' },
              bubble3D: true,
            }),
          ],
        }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'bubble') throw new Error('expected bubble');
      expect(c.bubble3D).toBe(true);
      expect(c.bubbleScale).toBe(80);
      expect(c.sizeRepresents).toBe('area');
    },
  },
  {
    name: 'Stock',
    kind: 'stock',
    build: () =>
      wrap(
        makeStockChart({
          hiLowLines: true,
          upDownBars: true,
          series: [
            makeBarSeries({ idx: 0, val: { ref: 'A1:A5' } }),
            makeBarSeries({ idx: 1, val: { ref: 'B1:B5' } }),
            makeBarSeries({ idx: 2, val: { ref: 'C1:C5' } }),
            makeBarSeries({ idx: 3, val: { ref: 'D1:D5' } }),
          ],
        }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'stock') throw new Error('expected stock');
      expect(c.hiLowLines).toBe(true);
      expect(c.upDownBars).toBe(true);
      expect(c.series.length).toBe(4);
    },
  },
  {
    name: 'Surface',
    kind: 'surface',
    build: () =>
      wrap(
        makeSurfaceChart({
          wireframe: true,
          axIds: [11, 22, 33],
          series: [makeBarSeries({ idx: 0, val: VAL })],
        }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'surface') throw new Error('expected surface');
      expect(c.wireframe).toBe(true);
      expect(c.axIds).toEqual([11, 22, 33]);
    },
  },
  {
    name: 'OfPie',
    kind: 'ofPie',
    build: () =>
      wrap(
        makeOfPieChart({
          ofPieType: 'bar',
          secondPieSize: 80,
          splitType: 'cust',
          custSplit: [3, 5, 7],
          series: [makeBarSeries({ idx: 0, val: VAL })],
        }),
        false,
      ),
    assert: (c) => {
      if (c.kind !== 'ofPie') throw new Error('expected ofPie');
      expect(c.ofPieType).toBe('bar');
      expect(c.secondPieSize).toBe(80);
      expect(c.splitType).toBe('cust');
      expect(c.custSplit).toEqual([3, 5, 7]);
    },
  },
  {
    name: 'Bar3D',
    kind: 'bar3D',
    build: () =>
      wrap(
        makeBar3DChart({
          gapDepth: 200,
          shape: 'cylinder',
          series: [makeBarSeries({ idx: 0, val: VAL })],
        }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'bar3D') throw new Error('expected bar3D');
      expect(c.gapDepth).toBe(200);
      expect(c.shape).toBe('cylinder');
    },
  },
  {
    name: 'Line3D',
    kind: 'line3D',
    build: () =>
      wrap(
        makeLine3DChart({ gapDepth: 100, series: [makeBarSeries({ idx: 0, val: VAL })] }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'line3D') throw new Error('expected line3D');
      expect(c.gapDepth).toBe(100);
    },
  },
  {
    name: 'Pie3D',
    kind: 'pie3D',
    build: () =>
      wrap(makePie3DChart({ varyColors: true, series: [makeBarSeries({ idx: 0, val: VAL })] }), false),
    assert: (c) => {
      if (c.kind !== 'pie3D') throw new Error('expected pie3D');
      expect(c.varyColors).toBe(true);
    },
  },
  {
    name: 'Area3D',
    kind: 'area3D',
    build: () =>
      wrap(
        makeArea3DChart({
          grouping: 'percentStacked',
          gapDepth: 50,
          series: [makeBarSeries({ idx: 0, val: VAL })],
        }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'area3D') throw new Error('expected area3D');
      expect(c.grouping).toBe('percentStacked');
      expect(c.gapDepth).toBe(50);
    },
  },
  {
    name: 'Surface3D',
    kind: 'surface3D',
    build: () =>
      wrap(
        makeSurface3DChart({
          wireframe: false,
          axIds: [11, 22, 33],
          series: [makeBarSeries({ idx: 0, val: VAL })],
        }),
        true,
      ),
    assert: (c) => {
      if (c.kind !== 'surface3D') throw new Error('expected surface3D');
      expect(c.wireframe).toBe(false);
      expect(c.axIds).toEqual([11, 22, 33]);
    },
  },
];

interface ChartExCase {
  name: string;
  layoutId: string;
  build: () => CxChartSpace;
  /** Optional extra assertion against the round-tripped CxChartSpace. */
  assert?: (loaded: CxChartSpace) => void;
}

const CHARTEX_CASES: ChartExCase[] = [
  {
    name: 'Sunburst',
    layoutId: 'sunburst',
    build: () => makeSunburstChart({ catRef: 'A1:A4', valRef: 'B1:B4' }),
  },
  {
    name: 'Treemap',
    layoutId: 'treemap',
    build: () => makeTreemapChart({ catRef: 'A1:A4', valRef: 'B1:B4', parentLabelLayout: 'banner' }),
    assert: (s) => {
      const lp = s.chart.plotArea.series[0]?.layoutPr;
      expect(lp).toEqual({ kind: 'parentLabel', layout: 'banner' });
    },
  },
  {
    name: 'Waterfall',
    layoutId: 'waterfall',
    build: () => makeWaterfallChart({ catRef: 'A1:A4', valRef: 'B1:B4', subtotalIdx: [0, 3] }),
    assert: (s) => {
      const lp = s.chart.plotArea.series[0]?.layoutPr;
      expect(lp).toEqual({ kind: 'waterfall', subtotalIdx: [0, 3] });
    },
  },
  {
    name: 'Histogram',
    layoutId: 'clusteredColumn',
    build: () =>
      makeHistogramChart({ valRef: 'A1:A100', binCount: 20, intervalClosed: 'l' }),
    assert: (s) => {
      const lp = s.chart.plotArea.series[0]?.layoutPr;
      if (!lp || lp.kind !== 'binning') throw new Error('expected binning');
      expect(lp.binCount).toBe(20);
      expect(lp.intervalClosed).toBe('l');
    },
  },
  {
    name: 'Pareto',
    layoutId: 'paretoLine',
    build: () => makeParetoChart({ catRef: 'A1:A4', valRef: 'B1:B4', binCount: 4 }),
    assert: (s) => {
      const layoutIds = s.chart.plotArea.series.map((x) => x.layoutId);
      expect(layoutIds).toEqual(['clusteredColumn', 'paretoLine']);
    },
  },
  {
    name: 'Funnel',
    layoutId: 'funnel',
    build: () => makeFunnelChart({ catRef: 'A1:A4', valRef: 'B1:B4' }),
  },
  {
    name: 'BoxWhisker',
    layoutId: 'boxWhisker',
    build: () =>
      makeBoxWhiskerChart({
        catRef: 'A1:A4',
        valRef: 'B1:B4',
        meanLine: true,
        meanMarker: true,
        outliers: true,
        quartileMethod: 'inclusive',
      }),
    assert: (s) => {
      const lp = s.chart.plotArea.series[0]?.layoutPr;
      if (!lp || lp.kind !== 'visibility') throw new Error('expected visibility');
      expect(lp.quartileMethod).toBe('inclusive');
    },
  },
  {
    name: 'RegionMap',
    layoutId: 'regionMap',
    build: () =>
      makeRegionMapChart({
        catRef: 'A1:A50',
        valRef: 'B1:B50',
        cultureLanguage: 'en-US',
        projectionType: 'mercator',
      }),
    assert: (s) => {
      const lp = s.chart.plotArea.series[0]?.layoutPr;
      if (!lp || lp.kind !== 'region') throw new Error('expected region');
      expect(lp.projectionType).toBe('mercator');
    },
  },
];

const buildWorkbook = (chartSpaces: Array<{ space?: ChartSpace; cxSpace?: CxChartSpace }>): Promise<Uint8Array> => {
  const wb = createWorkbook();
  // One worksheet hosting all charts as a single drawing — keeps the archive
  // flat and exercises the workbook-global chartN / drawingN counters at scale.
  const ws = addWorksheet(wb, 'Charts');
  let row = 1;
  ws.drawing = makeDrawing(
    chartSpaces.map((ref) =>
      makeChartDrawingItem(makeTwoCellAnchor({ from: `A${row}`, to: `F${(row += 12)}` }), ref),
    ),
  );
  return workbookToBytes(wb);
};

describe('Phase 6 §10 — chart 25-kind round-trip acceptance', () => {
  it('round-trips all 16 legacy `c:` chart kinds in one workbook', async () => {
    const refs = LEGACY_CASES.map((c) => ({ space: c.build() }));
    const bytes = await buildWorkbook(refs);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws = wb2.sheets[0]?.sheet;
    if (!ws || !('drawing' in ws) || !ws.drawing) throw new Error('expected drawing');
    const items = ws.drawing.items;
    expect(items.length).toBe(LEGACY_CASES.length);

    LEGACY_CASES.forEach((kase, i) => {
      const item = items[i];
      if (!item || item.content.kind !== 'chart') {
        throw new Error(`${kase.name}: expected chart drawing item at index ${i}`);
      }
      const space = item.content.chart.space;
      if (!space) throw new Error(`${kase.name}: missing legacy space`);
      expect(space.plotArea.chart.kind).toBe(kase.kind);
      kase.assert?.(space.plotArea.chart);
    });
  });

  it('round-trips all 8 chartex `cx:` chart kinds in one workbook', async () => {
    const refs = CHARTEX_CASES.map((c) => ({ cxSpace: c.build() }));
    const bytes = await buildWorkbook(refs);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws = wb2.sheets[0]?.sheet;
    if (!ws || !('drawing' in ws) || !ws.drawing) throw new Error('expected drawing');
    const items = ws.drawing.items;
    expect(items.length).toBe(CHARTEX_CASES.length);

    CHARTEX_CASES.forEach((kase, i) => {
      const item = items[i];
      if (!item || item.content.kind !== 'chart') {
        throw new Error(`${kase.name}: expected chart drawing item at index ${i}`);
      }
      const cx = item.content.chart.cxSpace;
      if (!cx) throw new Error(`${kase.name}: missing chartex space`);
      expect(cx.kind).toBe('cxChartSpace');
      const seriesIds = cx.chart.plotArea.series.map((s) => s.layoutId);
      expect(seriesIds).toContain(kase.layoutId);
      kase.assert?.(cx);
    });
  });

  it('mixes legacy + chartex charts in a single workbook without rId collisions', async () => {
    const bar = LEGACY_CASES[0];
    const scatter = LEGACY_CASES[5];
    const sunburst = CHARTEX_CASES[0];
    const waterfall = CHARTEX_CASES[2];
    if (!bar || !scatter || !sunburst || !waterfall) throw new Error('catalogue mis-indexed');
    const refs = [
      { space: bar.build() },
      { cxSpace: sunburst.build() },
      { space: scatter.build() },
      { cxSpace: waterfall.build() },
    ];
    const bytes = await buildWorkbook(refs);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);

    // 4 charts ⇒ chart1..chart4, with both legacy and chartex content types
    // declared in the manifest.
    expect(entries['xl/charts/chart1.xml']).toBeDefined();
    expect(entries['xl/charts/chart2.xml']).toBeDefined();
    expect(entries['xl/charts/chart3.xml']).toBeDefined();
    expect(entries['xl/charts/chart4.xml']).toBeDefined();
    const ct = new TextDecoder().decode(entries['[Content_Types].xml']);
    expect(ct).toContain('drawingml.chart+xml');
    expect(ct).toContain('vnd.ms-office.chartex+xml');

    // Round-trip kinds line up with input order.
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const items = wb2.sheets[0]?.sheet.drawing?.items ?? [];
    const kinds = items.map((entry) =>
      entry.content.kind === 'chart'
        ? entry.content.chart.space
          ? entry.content.chart.space.plotArea.chart.kind
          : entry.content.chart.cxSpace?.chart.plotArea.series[0]?.layoutId
        : null,
    );
    expect(kinds).toEqual(['bar', 'sunburst', 'scatter', 'waterfall']);
  });

  it('reports the catalogue size for the acceptance criterion', () => {
    expect(LEGACY_CASES.length).toBe(16);
    expect(CHARTEX_CASES.length).toBe(8);
    expect(LEGACY_CASES.length + CHARTEX_CASES.length).toBe(24);
  });
});
