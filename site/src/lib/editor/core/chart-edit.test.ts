import { describe, expect, it } from 'vitest';
import type { CalcScalar } from '../calc/index.ts';
import {
  applyChartStyle,
  applyColorSet,
  applyQuickLayout,
  chartElements,
  COLOR_SETS,
  setAxisTitle,
  setChartFill,
  setChartOutline,
  setChartTitle,
  setDataLabels,
  setGridline,
  setLegend,
} from './chart-edit.ts';
import { buildChart } from './charts.ts';

const ROWS: CalcScalar[][] = [
  [null, 'East', 'West'],
  ['Q1', 10, 20],
  ['Q2', 11, 21],
];
const valueAt = (r: number, c: number): CalcScalar => ROWS[r - 1]?.[c - 1] ?? null;
const RANGE = { r1: 1, c1: 1, r2: 3, c2: 3 };
const column = () => buildChart('columnClustered', 'S', RANGE, valueAt);

describe('Add Chart Element', () => {
  it('reports what a freshly inserted chart shows', () => {
    const e = chartElements(column());
    expect(e).toMatchObject({ title: 'none', legend: 'b', labels: 'none', catTitle: false, valTitle: false, hasAxes: true });
    expect(e.gridlines).toEqual({ valMajor: true, valMinor: false, catMajor: false });
  });

  it('turns elements on and off', () => {
    const chart = column();
    setChartTitle(chart, 'overlay', 'Revenue');
    setAxisTitle(chart, 'val', true);
    setLegend(chart, 'r');
    setDataLabels(chart, 'center');
    setGridline(chart, 'valMajor', false);
    setGridline(chart, 'catMajor', true);
    const e = chartElements(chart);
    expect(e).toMatchObject({ title: 'overlay', titleText: 'Revenue', valTitle: true, legend: 'r', labels: 'center' });
    expect(e.gridlines).toEqual({ valMajor: false, valMinor: false, catMajor: true });
    expect(chart.space?.plotArea.chart.series.every((s) => s.dLbls?.dLblPos === 'ctr')).toBe(true);

    setChartTitle(chart, 'none');
    setLegend(chart, 'none');
    setDataLabels(chart, 'none');
    expect(chartElements(chart)).toMatchObject({ title: 'none', legend: 'none', labels: 'none' });
  });

  it('keeps the typed title text when moving the title', () => {
    const chart = column();
    setChartTitle(chart, 'above', 'Kept');
    setChartTitle(chart, 'overlay');
    expect(chartElements(chart).titleText).toBe('Kept');
  });

  it('edits chartex titles and legends', () => {
    const chart = buildChart('treemap', 'S', RANGE, valueAt);
    setChartTitle(chart, 'above', 'Tree');
    setLegend(chart, 't');
    expect(chartElements(chart)).toMatchObject({ title: 'above', titleText: 'Tree', legend: 't', hasAxes: false });
  });
});

describe('layouts, colours and styles', () => {
  it('applies a quick layout', () => {
    const chart = column();
    applyQuickLayout(chart, 5);
    expect(chartElements(chart)).toMatchObject({ title: 'above', legend: 'none', catTitle: true, valTitle: true });
  });

  it('colours each series from the set', () => {
    const chart = column();
    const set = COLOR_SETS.find((s) => s.id === 'mono2');
    if (!set) throw new Error('missing set');
    applyColorSet(chart, set);
    const fills = chart.space?.plotArea.chart.series.map((s) => s.spPr?.fill);
    expect(fills?.[0]).toEqual({ kind: 'solidFill', color: set.colors[0] });
    expect(fills?.[1]).toEqual({ kind: 'solidFill', color: set.colors[1] });
  });

  it('colours each slice of a single-series pie', () => {
    const chart = buildChart('pie', 'S', { r1: 1, c1: 1, r2: 3, c2: 2 }, valueAt);
    applyColorSet(chart, COLOR_SETS[0] ?? { id: '', colors: [] });
    expect(chart.space?.plotArea.chart.series[0]?.dPt?.map((p) => p.idx)).toEqual([0, 1]);
  });

  it('sets the frame for a chart style', () => {
    const chart = column();
    applyChartStyle(chart, 4);
    expect(chart.space?.spPr?.fill).toEqual({ kind: 'solidFill', color: { base: { kind: 'srgb', value: '404040' }, mods: [] } });
  });

  it('sets the chart area fill and outline', () => {
    const chart = column();
    setChartFill(chart, '#ffeeaa');
    setChartOutline(chart, null);
    expect(chart.space?.spPr?.fill).toEqual({ kind: 'solidFill', color: { base: { kind: 'srgb', value: 'FFEEAA' }, mods: [] } });
    expect(chart.space?.spPr?.ln).toEqual({ fill: { kind: 'noFill' } });
    setChartOutline(chart, '000000', 2);
    expect(chart.space?.spPr?.ln?.w).toBe(25_400);
  });
});
