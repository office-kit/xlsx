import { describe, expect, it } from 'vitest';
import {
  type CxChartSpace,
  makeBoxWhiskerChart,
  makeCxChartSpace,
  makeCxData,
  makeCxNumDim,
  makeCxSeries,
  makeCxStrDim,
  makeFunnelChart,
  makeHistogramChart,
  makeParetoChart,
  makeRegionMapChart,
  makeSunburstChart,
  makeTreemapChart,
  makeWaterfallChart,
} from '../../src/chart/cx/chartex.js';
import { chartExToBytes, isChartExBytes, parseChartExXml, serializeChartExSpace } from '../../src/chart/cx/chartex-xml.js';

const roundTrip = (space: CxChartSpace): CxChartSpace => parseChartExXml(chartExToBytes(space));

describe('chartex sniff', () => {
  it('detects cx:chartSpace bytes', () => {
    const xml = serializeChartExSpace(makeSunburstChart({}));
    expect(isChartExBytes(xml)).toBe(true);
  });

  it('rejects c:chartSpace bytes', () => {
    expect(
      isChartExBytes(
        '<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
      ),
    ).toBe(false);
  });
});

describe('Sunburst round-trip', () => {
  it('preserves cat + val dims', () => {
    const back = roundTrip(makeSunburstChart({ catRef: 'Sheet1!$A$1:$A$4', valRef: 'Sheet1!$B$1:$B$4' }));
    expect(back.kind).toBe('cxChartSpace');
    expect(back.chart.plotArea.series[0]?.layoutId).toBe('sunburst');
    expect(back.chart.plotArea.series[0]?.dataId).toBe(0);
    const dims = back.chartData.data[0]?.dims;
    expect(dims?.[0]).toEqual(expect.objectContaining({ kind: 'str', type: 'cat', f: 'Sheet1!$A$1:$A$4' }));
    expect(dims?.[1]).toEqual(expect.objectContaining({ kind: 'num', type: 'val', f: 'Sheet1!$B$1:$B$4' }));
  });
});

describe('Treemap round-trip', () => {
  it('preserves parentLabelLayout', () => {
    const back = roundTrip(makeTreemapChart({ catRef: 'A1:A3', valRef: 'B1:B3', parentLabelLayout: 'banner' }));
    const series = back.chart.plotArea.series[0];
    expect(series?.layoutId).toBe('treemap');
    expect(series?.layoutPr).toEqual({ kind: 'parentLabel', layout: 'banner' });
  });
});

describe('Waterfall round-trip', () => {
  it('preserves subtotal indices', () => {
    const back = roundTrip(makeWaterfallChart({ catRef: 'A1:A5', valRef: 'B1:B5', subtotalIdx: [0, 4] }));
    const series = back.chart.plotArea.series[0];
    expect(series?.layoutId).toBe('waterfall');
    expect(series?.layoutPr).toEqual({ kind: 'waterfall', subtotalIdx: [0, 4] });
    expect(series?.axisIds).toEqual([0, 1]);
    expect(back.chart.plotArea.axes.map((a) => a.id)).toEqual([0, 1]);
  });
});

describe('Histogram round-trip', () => {
  // CT_Binning holds the bin count or size as a child element (a choice), and
  // Excel asks to repair a histogram that carries them as attributes.
  it('preserves the binning bounds and the bin size', () => {
    const back = roundTrip(
      makeHistogramChart({ valRef: 'A1:A100', binSize: 5, intervalClosed: 'l', underflow: 0, overflow: 100 }),
    );
    const series = back.chart.plotArea.series[0];
    expect(series?.layoutId).toBe('clusteredColumn');
    expect(series?.layoutPr).toEqual({ kind: 'binning', binSize: 5, intervalClosed: 'l', underflow: 0, overflow: 100 });
  });

  it('writes the bin count as a child element, and nothing for automatic binning', () => {
    const fixed = new TextDecoder().decode(chartExToBytes(makeHistogramChart({ binCount: 20 })));
    expect(fixed).toContain('<cx:binning><cx:binCount val="20"/></cx:binning>');
    expect(roundTrip(makeHistogramChart({ binCount: 20 })).chart.plotArea.series[0]?.layoutPr).toEqual({ kind: 'binning', binCount: 20 });
    const auto = new TextDecoder().decode(chartExToBytes(makeHistogramChart({ binCountAuto: true })));
    // Excel refuses an empty `<cx:binCount/>`; automatic is the default.
    expect(auto).toContain('<cx:binning/>');
  });
});

describe('Pareto round-trip', () => {
  it('preserves clusteredColumn + paretoLine pair', () => {
    const back = roundTrip(makeParetoChart({ catRef: 'A1:A4', valRef: 'B1:B4', binCount: 4 }));
    expect(back.chart.plotArea.series.map((s) => s.layoutId)).toEqual(['clusteredColumn', 'paretoLine']);
    expect(back.chart.plotArea.series[1]?.ownerIdx).toBe(0);
    expect(back.chart.plotArea.axes.map((a) => a.id)).toEqual([0, 1, 2]);
  });
});

describe('Funnel round-trip', () => {
  it('preserves funnel layoutId without layoutPr', () => {
    const back = roundTrip(makeFunnelChart({ catRef: 'A1:A5', valRef: 'B1:B5' }));
    const series = back.chart.plotArea.series[0];
    expect(series?.layoutId).toBe('funnel');
    expect(series?.layoutPr).toBeUndefined();
  });
});

describe('BoxWhisker round-trip', () => {
  it('preserves visibility flags + quartileMethod', () => {
    const back = roundTrip(
      makeBoxWhiskerChart({
        catRef: 'A1:A3',
        valRef: 'B1:B3',
        meanLine: true,
        meanMarker: true,
        outliers: false,
        nonoutliers: true,
        quartileMethod: 'exclusive',
      }),
    );
    const series = back.chart.plotArea.series[0];
    expect(series?.layoutId).toBe('boxWhisker');
    expect(series?.layoutPr).toEqual({
      kind: 'visibility',
      meanLine: true,
      meanMarker: true,
      outliers: false,
      nonoutliers: true,
      quartileMethod: 'exclusive',
    });
  });

  // Excel reads the quartile method from `<cx:statistics>`; a
  // `<cx:quartileMethod>` element makes it ask to repair the workbook.
  it('writes the quartile method on cx:statistics', () => {
    const xml = new TextDecoder().decode(chartExToBytes(makeBoxWhiskerChart({ quartileMethod: 'inclusive' })));
    expect(xml).toContain('<cx:statistics quartileMethod="inclusive"/>');
    expect(xml).not.toContain('<cx:quartileMethod');
  });
});

describe('RegionMap round-trip', () => {
  it('preserves geography + projection + region label layout', () => {
    const back = roundTrip(
      makeRegionMapChart({
        catRef: 'A1:A50',
        valRef: 'B1:B50',
        cultureLanguage: 'en-US',
        cultureRegion: 'US',
        projectionType: 'mercator',
        regionLabelLayout: 'bestFit',
      }),
    );
    const series = back.chart.plotArea.series[0];
    expect(series?.layoutId).toBe('regionMap');
    expect(series?.layoutPr).toEqual({
      kind: 'region',
      cultureLanguage: 'en-US',
      cultureRegion: 'US',
      projectionType: 'mercator',
      regionLabelLayout: 'bestFit',
    });
  });
});

describe('chartex chart-level metadata', () => {
  it('round-trips title text + legend pos + plotVisOnly + dispBlanksAs', () => {
    const space = makeCxChartSpace({
      data: [makeCxData(0, [makeCxStrDim({ type: 'cat', f: 'A1:A3' }), makeCxNumDim({ type: 'val', f: 'B1:B3' })])],
      series: [makeCxSeries({ layoutId: 'sunburst', dataId: 0 })],
      title: { pos: 't', align: 'ctr', overlay: false, text: 'My Chart' },
      legend: { pos: 'b', align: 'ctr', overlay: false },
      plotVisOnly: true,
      dispBlanksAs: 'gap',
    });
    const back = roundTrip(space);
    expect(back.chart.title?.text).toBe('My Chart');
    expect(back.chart.title?.pos).toBe('t');
    expect(back.chart.legend?.pos).toBe('b');
    expect(back.chart.plotVisOnly).toBe(true);
    expect(back.chart.dispBlanksAs).toBe('gap');
  });
});

describe('chartex point caches', () => {
  it('preserves numeric and string lvl/pt values', () => {
    const space = makeCxChartSpace({
      data: [
        makeCxData(0, [
          makeCxStrDim({
            type: 'cat',
            f: 'A1:A3',
            ptCount: 3,
            pts: [
              { idx: 0, v: 'A' },
              { idx: 1, v: 'B' },
              { idx: 2, v: 'C' },
            ],
          }),
          makeCxNumDim({
            type: 'val',
            f: 'B1:B3',
            ptCount: 3,
            pts: [
              { idx: 0, v: '10' },
              { idx: 1, v: '20' },
              { idx: 2, v: '30' },
            ],
            formatCode: '0.00',
          }),
        ]),
      ],
      series: [makeCxSeries({ layoutId: 'sunburst', dataId: 0 })],
    });
    const back = roundTrip(space);
    const dims = back.chartData.data[0]?.dims;
    expect(dims?.[0]?.pts.map((p) => p.v)).toEqual(['A', 'B', 'C']);
    expect(dims?.[1]?.pts.map((p) => p.v)).toEqual(['10', '20', '30']);
    expect(dims?.[1]?.formatCode).toBe('0.00');
  });
});
