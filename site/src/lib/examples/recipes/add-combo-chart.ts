// Draw revenue as columns and margin as a line on a secondary value axis.

import { makeBarChart, makeBarSeries, makeChartSpace, makeLineChart } from '@office-kit/xlsx/chart';
import { addChartAt } from '@office-kit/xlsx/drawing';
import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { setCell } from '@office-kit/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Sales');

const rows: Array<[string, number, number]> = [
  ['Q1', 12_400, 0.18],
  ['Q2', 9_800, 0.22],
  ['Q3', 7_300, 0.27],
];
setCell(ws, 1, 1, 'Quarter');
setCell(ws, 1, 2, 'Revenue');
setCell(ws, 1, 3, 'Margin');
rows.forEach(([quarter, revenue, margin], i) => {
  setCell(ws, i + 2, 1, quarter);
  setCell(ws, i + 2, 2, revenue);
  setCell(ws, i + 2, 3, margin);
});

// Both groups share the category axis; the line gets a value axis of its own.
const CATEGORY_AXIS = 10;
const REVENUE_AXIS = 100;
const MARGIN_AXIS = 200;
const categories = { ref: 'Sales!$A$2:$A$4' };

const space = makeChartSpace({
  plotArea: {
    chart: makeBarChart({
      axIds: [CATEGORY_AXIS, REVENUE_AXIS],
      series: [makeBarSeries({ idx: 0, tx: { kind: 'ref', ref: 'Sales!$B$1' }, cat: categories, val: { ref: 'Sales!$B$2:$B$4' } })],
    }),
    additionalCharts: [
      makeLineChart({
        axIds: [CATEGORY_AXIS, MARGIN_AXIS],
        series: [makeBarSeries({ idx: 1, tx: { kind: 'ref', ref: 'Sales!$C$1' }, cat: categories, val: { ref: 'Sales!$C$2:$C$4' } })],
      }),
    ],
    catAx: { axId: CATEGORY_AXIS, crossAx: REVENUE_AXIS },
    valAx: { axId: REVENUE_AXIS, crossAx: CATEGORY_AXIS, majorGridlines: true },
    additionalValAx: [
      {
        axId: MARGIN_AXIS,
        crossAx: CATEGORY_AXIS,
        crosses: 'max',
        position: 'r',
        numFmt: { formatCode: '0%', sourceLinked: false },
      },
    ],
  },
  legend: { position: 't' },
});

addChartAt(ws, 'E2', { space }, { widthPx: 480, heightPx: 320 });

await saveWorkbook(wb, toFile('combo-chart.xlsx'));
