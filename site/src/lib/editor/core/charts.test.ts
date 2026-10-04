import type { BarChart, ChartSpace, ScatterChart } from '@office-kit/xlsx/chart';
import type { ChartReference } from '@office-kit/xlsx/drawing';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { setCell } from '@office-kit/xlsx/worksheet';
import { describe, expect, it } from 'vitest';
import { CalcEngine, type CalcScalar } from '../calc/index.ts';
import { chartData } from './chart-data.ts';
import { analyzeSource, buildChart, CHART_CHOICES, chartChoiceOf, chartSource, rebuildChart, type ValueAt } from './charts.ts';
import { OFFICE_PALETTE } from './theme.ts';

function grid(rows: CalcScalar[][]): (r: number, c: number) => CalcScalar {
  return (r, c) => rows[r - 1]?.[c - 1] ?? null;
}

const SALES = grid([
  [null, 'East', 'West'],
  ['Q1', 10, 20],
  ['Q2', 11, 21],
  ['Q3', 12, 22],
]);
const ALL = { r1: 1, c1: 1, r2: 4, c2: 3 };

function classic(chart: ChartReference): ChartSpace {
  if (!chart.space) throw new Error('expected a c: chart');
  return chart.space;
}

const buildChartSpace = (choice: Parameters<typeof buildChart>[0], sheet: string, range: typeof ALL, valueAt: ValueAt): ChartSpace => classic(buildChart(choice, sheet, range, valueAt));

describe('analyzeSource', () => {
  it('detects a header row and a category column', () => {
    expect(analyzeSource(ALL, SALES)).toEqual({ headerRow: true, headerCol: true, byColumns: true });
  });

  it('treats an all-number block as plain data', () => {
    const v = grid([
      [1, 2],
      [3, 4],
    ]);
    expect(analyzeSource({ r1: 1, c1: 1, r2: 2, c2: 2 }, v)).toEqual({ headerRow: false, headerCol: false, byColumns: true });
  });

  it('plots series by rows when there are more columns than rows', () => {
    const v = grid([
      [null, 'Jan', 'Feb', 'Mar'],
      ['A', 1, 2, 3],
    ]);
    expect(analyzeSource({ r1: 1, c1: 1, r2: 2, c2: 4 }, v).byColumns).toBe(false);
  });
});

describe('buildChart', () => {
  it('builds one clustered column series per data column', () => {
    const space = buildChartSpace('columnClustered', 'Sheet 1', ALL, SALES);
    const chart = space.plotArea.chart as BarChart;
    expect(chart.kind).toBe('bar');
    expect(chart.barDir).toBe('col');
    expect(chart.series.map((s) => s.val.ref)).toEqual(["'Sheet 1'!$B$2:$B$4", "'Sheet 1'!$C$2:$C$4"]);
    expect(chart.series[0]?.cat?.ref).toBe("'Sheet 1'!$A$2:$A$4");
    expect(chart.series[1]?.tx).toEqual({ kind: 'ref', ref: "'Sheet 1'!$C$1" });
    expect(space.legend?.position).toBe('b');
  });

  it('titles a single-series chart with the series name', () => {
    const space = buildChartSpace('line', 'S', { r1: 1, c1: 1, r2: 4, c2: 2 }, SALES);
    expect(space.title?.text).toBe('East');
    expect(space.legend).toBeUndefined();
  });

  it('uses the first numeric column as X for scatter', () => {
    const v = grid([
      ['x', 'y'],
      [1, 2],
      [2, 4],
    ]);
    const chart = buildChartSpace('scatter', 'S', { r1: 1, c1: 1, r2: 3, c2: 2 }, v).plotArea.chart as ScatterChart;
    expect(chart.series).toHaveLength(1);
    expect(chart.series[0]?.xVal?.ref).toBe('S!$A$2:$A$3');
    expect(chart.series[0]?.yVal.ref).toBe('S!$B$2:$B$3');
  });
});

describe('chartData', () => {
  it('resolves refs live from the sheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet 1');
    const rows: CalcScalar[][] = [
      [null, 'East', 'West'],
      ['Q1', 10, 20],
      ['Q2', 11, 21],
    ];
    rows.forEach((row, r) => row.forEach((v, c) => typeof v !== 'object' && setCell(ws, r + 1, c + 1, v)));
    const calc = new CalcEngine(wb);
    calc.recalculateAll();
    const chart = buildChart('columnClustered', 'Sheet 1', { r1: 1, c1: 1, r2: 3, c2: 3 }, (r, c) => calc.cellValue('Sheet 1', r, c));
    setCell(ws, 2, 2, 99);
    const sheets = new Map<string, Worksheet>([['Sheet 1', ws]]);
    const data = chartData(chart, calc, sheets, OFFICE_PALETTE);
    expect(data?.categories).toEqual(['Q1', 'Q2']);
    expect(data?.series.map((s) => s.name)).toEqual(['East', 'West']);
    expect(data?.series[0]?.values).toEqual([99, 11]);
    expect(data?.series[0]?.color).toBe('#4472C4');
  });
});

describe('chart round trips', () => {
  it('recognises every inserted chart type again', () => {
    for (const choice of CHART_CHOICES) expect(chartChoiceOf(buildChart(choice, 'S', ALL, SALES)), choice).toBe(choice);
  });

  it('recovers the source block and orientation from the series refs', () => {
    expect(chartSource(buildChart('columnClustered', 'Sheet 1', ALL, SALES))).toEqual({ sheet: 'Sheet 1', range: ALL, byColumns: true });
    expect(chartSource(buildChart('columnClustered', 'S', ALL, SALES, false))).toEqual({ sheet: 'S', range: ALL, byColumns: false });
  });

  it('keeps the title, legend and series formatting when changing type', () => {
    const old = buildChart('columnClustered', 'S', ALL, SALES);
    const space = classic(old);
    space.title = { text: 'Kept' };
    const first = space.plotArea.chart.series[0];
    if (first) first.dLbls = { showVal: true };
    const next = classic(rebuildChart(old, 'line', { sheet: 'S', range: ALL, byColumns: true }, SALES));
    expect(next.plotArea.chart.kind).toBe('line');
    expect(next.title?.text).toBe('Kept');
    expect(next.legend?.position).toBe('b');
    expect(next.plotArea.chart.series[0]?.dLbls?.showVal).toBe(true);
  });

  it('fills the chartex caches from the cells', () => {
    const cx = buildChart('treemap', 'S', ALL, SALES).cxSpace;
    const dims = cx?.chartData.data[0]?.dims ?? [];
    expect(dims.find((d) => d.kind === 'num')?.pts.map((p) => p.v)).toEqual(['10', '11', '12']);
  });
});
