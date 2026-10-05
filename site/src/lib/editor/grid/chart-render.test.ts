import { describe, expect, it } from 'vitest';
import type { ChartData, RenderSeries } from '../core/chart-data.ts';
import { OFFICE_PALETTE } from '../core/theme.ts';
import { histogramBins, niceScale, renderChartSvg, squarify } from './chart-render.ts';

const EAST: RenderSeries = { name: 'East', values: [10, 20], color: '#4472C4', markers: false, line: false, smooth: false, labels: undefined };

const base: ChartData = {
  kind: 'bar',
  horizontal: false,
  grouping: 'clustered',
  categories: ['Q1', 'Q2'],
  series: [EAST, { ...EAST, name: 'W<est>', values: [5, null], color: '#ED7D31' }],
  title: 'Sales',
  legend: 'b',
  varyColors: false,
  holeSize: 0,
  gapWidth: 150,
  overlap: 0,
  valMin: undefined,
  valMax: undefined,
  catTitle: undefined,
  valTitle: undefined,
  catAxis: true,
  valAxis: true,
  majorGridlines: true,
  minorGridlines: false,
  catGridlines: false,
  chartFill: '#FFFFFF',
  chartLine: { color: '#D9D9D9', width: 1 },
  plotFill: null,
  radarFilled: false,
  upDownBars: false,
  subtotals: [],
};

describe('niceScale', () => {
  it('rounds to 1/2/5 steps', () => {
    expect(niceScale(0, 23, 5)).toEqual({ min: 0, max: 25, step: 5 });
    expect(niceScale(-3, 7, 5)).toEqual({ min: -4, max: 8, step: 2 });
  });
});

describe('renderChartSvg', () => {
  it('draws one bar per value and escapes text', () => {
    const svg = renderChartSvg(base, 400, 300, OFFICE_PALETTE);
    expect(svg.match(/<rect [^>]*fill="#4472C4"/g)).toHaveLength(3); // two bars + legend key
    expect(svg).toContain('W&lt;est&gt;');
    expect(svg).toContain('>Sales<');
  });

  it('draws pie slices for each positive value', () => {
    const svg = renderChartSvg({ ...base, kind: 'pie', series: [{ ...EAST, values: [1, 1, 2] }], categories: ['a', 'b', 'c'] }, 300, 300, OFFICE_PALETTE);
    expect(svg.match(/<path d="M[^"]*A/g)).toHaveLength(3);
  });

  it('draws data labels, axis titles and the chart frame', () => {
    const labels = { pos: 'outEnd' as const, value: true, percent: false, category: false, series: false };
    const svg = renderChartSvg({ ...base, series: [{ ...EAST, values: [1234.5, 2], labels }], valTitle: 'Revenue', chartFill: '#404040' }, 400, 300, OFFICE_PALETTE);
    expect(svg).toContain('>1,234.5<');
    expect(svg).toContain('>Revenue<');
    expect(svg).toContain('fill="#404040" stroke="#D9D9D9"');
    // Light text on a dark chart area.
    expect(svg).toContain('fill="#F2F2F2"');
  });

  it('hides axis labels for a deleted axis', () => {
    const svg = renderChartSvg({ ...base, catAxis: false }, 400, 300, OFFICE_PALETTE);
    expect(svg).not.toContain('>Q1<');
  });
});

describe('squarify', () => {
  it('tiles the box with areas proportional to the values', () => {
    const tiles = squarify([6, 6, 4, 3, 2, 2, 1], { x: 0, y: 0, w: 6, h: 4 });
    const area = tiles.reduce((n, t) => n + t.w * t.h, 0);
    expect(area).toBeCloseTo(24);
    expect((tiles[0]?.w ?? 0) * (tiles[0]?.h ?? 0)).toBeCloseTo(6);
    for (const t of tiles) {
      expect(t.x + t.w).toBeLessThanOrEqual(6 + 1e-9);
      expect(t.y + t.h).toBeLessThanOrEqual(4 + 1e-9);
    }
  });
});

describe('histogramBins', () => {
  it('counts every value exactly once', () => {
    const values = [1, 2, 2, 3, 3, 3, 4, 4, 5, 9];
    const bins = histogramBins(values);
    expect(bins.reduce((n, b) => n + b.count, 0)).toBe(values.length);
    expect(bins[0]?.lo).toBe(1);
    expect(bins[bins.length - 1]?.hi).toBeGreaterThanOrEqual(9);
  });
});
