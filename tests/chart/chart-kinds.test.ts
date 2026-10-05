import { describe, expect, it } from 'vitest';
import {
  type AreaChart,
  type DoughnutChart,
  type LineChart,
  makeAreaChart,
  makeBarSeries,
  makeChartSpace,
  makeDoughnutChart,
  makeLineChart,
  makePieChart,
  makeRadarChart,
  makeScatterChart,
  makeScatterSeries,
  type PieChart,
  type RadarChart,
  type ScatterChart,
} from '../../src/chart/chart.js';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml.js';

const wrap = <T>(chart: { kind: string }, withAxes = true) =>
  makeChartSpace({
    plotArea: {
      // biome-ignore lint/suspicious/noExplicitAny: helper for varied chart kinds
      chart: chart as any,
      ...(withAxes ? { catAx: { axId: 1, crossAx: 2 }, valAx: { axId: 2, crossAx: 1 } } : {}),
    },
  }) as T;

describe('LineChart round-trip', () => {
  it('preserves grouping / smooth / per-series smooth', () => {
    const chart = makeLineChart({
      grouping: 'standard',
      smooth: false,
      series: [
        makeBarSeries({ idx: 0, val: { ref: 'A1:A4', cache: [1, 2, 3, 4] } }),
        { ...makeBarSeries({ idx: 1, val: { ref: 'B1:B4' } }), smooth: true },
      ],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('line');
    const line = back.plotArea.chart as LineChart;
    expect(line.grouping).toBe('standard');
    expect(line.smooth).toBe(false);
    expect(line.series.length).toBe(2);
    expect(line.series[0]?.val.cache).toEqual([1, 2, 3, 4]);
    expect(line.series[1]?.smooth).toBe(true);
  });
});

describe('AreaChart round-trip', () => {
  it('preserves grouping + series', () => {
    const chart = makeAreaChart({
      grouping: 'stacked',
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A3' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('area');
    const area = back.plotArea.chart as AreaChart;
    expect(area.grouping).toBe('stacked');
    expect(area.series.length).toBe(1);
  });
});

describe('PieChart round-trip', () => {
  it('preserves single-series pie without axes', () => {
    const chart = makePieChart({
      varyColors: true,
      series: [
        makeBarSeries({
          idx: 0,
          val: { ref: 'A1:A4', cache: [40, 30, 20, 10] },
          cat: { ref: 'B1:B4', cacheKind: 'str', cache: ['a', 'b', 'c', 'd'] },
        }),
      ],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart, false);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('pie');
    const pie = back.plotArea.chart as PieChart;
    expect(pie.varyColors).toBe(true);
    expect(pie.series[0]?.val.cache).toEqual([40, 30, 20, 10]);
    expect(pie.series[0]?.cat?.cache).toEqual(['a', 'b', 'c', 'd']);
  });
});

describe('PieChart rotation', () => {
  it('preserves firstSliceAng', () => {
    const chart = makePieChart({ firstSliceAng: 90, series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A2' } })] });
    const back = parseChartXml(chartToBytes(wrap<ReturnType<typeof makeChartSpace>>(chart, false)));
    expect(back.plotArea.chart).toMatchObject({ kind: 'pie', firstSliceAng: 90 });
  });
});

describe('DoughnutChart round-trip', () => {
  it('preserves holeSize + firstSliceAng', () => {
    const chart = makeDoughnutChart({
      holeSize: 60,
      firstSliceAng: 45,
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A2' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart, false);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('doughnut');
    const donut = back.plotArea.chart as DoughnutChart;
    expect(donut.holeSize).toBe(60);
    expect(donut.firstSliceAng).toBe(45);
  });
});

describe('ScatterChart round-trip', () => {
  it('preserves scatterStyle + xVal/yVal pair', () => {
    const chart = makeScatterChart({
      scatterStyle: 'smoothMarker',
      series: [
        makeScatterSeries({
          idx: 0,
          xVal: { ref: 'A1:A5', cache: [1, 2, 3, 4, 5] },
          yVal: { ref: 'B1:B5', cache: [10, 8, 6, 4, 2] },
          smooth: true,
        }),
      ],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('scatter');
    const scatter = back.plotArea.chart as ScatterChart;
    expect(scatter.scatterStyle).toBe('smoothMarker');
    const s = scatter.series[0];
    expect(s?.xVal?.cache).toEqual([1, 2, 3, 4, 5]);
    expect(s?.yVal.cache).toEqual([10, 8, 6, 4, 2]);
    expect(s?.smooth).toBe(true);
  });
});

describe('RadarChart round-trip', () => {
  it('preserves radarStyle + series', () => {
    const chart = makeRadarChart({
      radarStyle: 'marker',
      series: [makeBarSeries({ idx: 0, val: { ref: 'A1:A4' } })],
    });
    const space = wrap<ReturnType<typeof makeChartSpace>>(chart);
    const back = parseChartXml(chartToBytes(space));
    expect(back.plotArea.chart.kind).toBe('radar');
    const radar = back.plotArea.chart as RadarChart;
    expect(radar.radarStyle).toBe('marker');
    expect(radar.series.length).toBe(1);
  });
});

describe('Unknown chart kind rejection', () => {
  it('throws when no supported chart child is found', () => {
    // chartex namespace kinds (sunburst etc) are not yet wired into the c:
    // namespace dispatcher — we use a fictitious tag to verify the failure
    // path stays clean.
    const xml =
      '<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout/><c:fancyChart/></c:plotArea></c:chart></c:chartSpace>';
    expect(() => parseChartXml(xml)).toThrowError(/no supported chart kind/);
  });
});
