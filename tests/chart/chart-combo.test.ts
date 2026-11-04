import { describe, expect, it } from 'vitest';
import {
  type ChartSpace,
  makeBarChart,
  makeBarSeries,
  makeChartSpace,
  makeLineChart,
  makePieChart,
} from '../../src/chart/chart.js';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml.js';
import { makeTwoCellAnchor } from '../../src/drawing/anchor.js';
import { makeChartDrawingItem, makeDrawing } from '../../src/drawing/drawing.js';
import { makeSimpleTextBody } from '../../src/drawing/dml/text.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { validateXlsx } from '../conformance/validate.js';

const CAT_AX_ID = 10;
const PRIMARY_VAL_AX_ID = 100;
const SECONDARY_VAL_AX_ID = 200;

// Columns on the primary axis plus a line on a secondary value axis that shares the category axis.
const comboSpace = (): ChartSpace =>
  makeChartSpace({
    plotArea: {
      chart: makeBarChart({
        axIds: [CAT_AX_ID, PRIMARY_VAL_AX_ID],
        series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$4' }, cat: { ref: 'Sheet1!$A$2:$A$4' } })],
      }),
      additionalCharts: [
        makeLineChart({
          axIds: [CAT_AX_ID, SECONDARY_VAL_AX_ID],
          series: [{ ...makeBarSeries({ idx: 1, val: { ref: 'Sheet1!$C$2:$C$4' }, cat: { ref: 'Sheet1!$A$2:$A$4' } }), smooth: false }],
        }),
      ],
      catAx: { axId: CAT_AX_ID, crossAx: PRIMARY_VAL_AX_ID },
      valAx: { axId: PRIMARY_VAL_AX_ID, crossAx: CAT_AX_ID, majorGridlines: true },
      additionalValAx: [{ axId: SECONDARY_VAL_AX_ID, crossAx: CAT_AX_ID, crosses: 'max', position: 'r' }],
    },
    legend: { position: 't', legendEntry: [{ idx: 1, delete: true }] },
  });

const xmlOf = (space: ChartSpace): string => new TextDecoder().decode(chartToBytes(space));

describe('combo charts (several chart groups in one plot area)', () => {
  it('writes every chart group before the axes, and the extra value axis after the primary ones', () => {
    const xml = xmlOf(comboSpace());
    const order = ['<c:barChart>', '<c:lineChart>', '<c:catAx>', `<c:valAx><c:axId val="${PRIMARY_VAL_AX_ID}"/>`, `<c:valAx><c:axId val="${SECONDARY_VAL_AX_ID}"/>`];
    const positions = order.map((tag) => xml.indexOf(tag));
    expect(positions.every((p) => p >= 0)).toBe(true);
    expect([...positions].sort((a, b) => a - b)).toEqual(positions);
  });

  it('round-trips additionalCharts and additionalValAx', () => {
    const back = parseChartXml(chartToBytes(comboSpace()));
    expect(back.plotArea.chart.kind).toBe('bar');
    expect(back.plotArea.additionalCharts?.map((c) => c.kind)).toEqual(['line']);
    expect(back.plotArea.catAx?.axId).toBe(CAT_AX_ID);
    expect(back.plotArea.valAx?.axId).toBe(PRIMARY_VAL_AX_ID);
    expect(back.plotArea.additionalValAx).toHaveLength(1);
    expect(back.plotArea.additionalValAx?.[0]).toMatchObject({ axId: SECONDARY_VAL_AX_ID, crossAx: CAT_AX_ID, crosses: 'max', position: 'r' });
  });

  it('picks the primary value axis by the first chart group’s axIds, not by document order', () => {
    const space = comboSpace();
    const swapped: ChartSpace = {
      ...space,
      plotArea: {
        ...space.plotArea,
        valAx: { axId: SECONDARY_VAL_AX_ID, crossAx: CAT_AX_ID, crosses: 'max' },
        additionalValAx: [{ axId: PRIMARY_VAL_AX_ID, crossAx: CAT_AX_ID }],
      },
    };
    const back = parseChartXml(chartToBytes(swapped));
    expect(back.plotArea.valAx?.axId).toBe(PRIMARY_VAL_AX_ID);
    expect(back.plotArea.additionalValAx?.map((a) => a.axId)).toEqual([SECONDARY_VAL_AX_ID]);
  });

  it('keeps a second category axis, as Excel writes for a secondary axis pair', () => {
    const SECONDARY_CAT_AX_ID = 300;
    const space = comboSpace();
    const [line] = space.plotArea.additionalCharts ?? [];
    if (line?.kind !== 'line') throw new Error('expected line');
    const withSecondaryCat: ChartSpace = {
      ...space,
      plotArea: {
        ...space.plotArea,
        additionalCharts: [{ ...line, axIds: [SECONDARY_CAT_AX_ID, SECONDARY_VAL_AX_ID] }],
        additionalCatAx: [{ axId: SECONDARY_CAT_AX_ID, crossAx: SECONDARY_VAL_AX_ID, delete: true }],
        additionalValAx: [{ axId: SECONDARY_VAL_AX_ID, crossAx: SECONDARY_CAT_AX_ID, crosses: 'max' }],
      },
    };
    const back = parseChartXml(chartToBytes(withSecondaryCat));
    expect(back.plotArea.catAx?.axId).toBe(CAT_AX_ID);
    expect(back.plotArea.additionalCatAx).toHaveLength(1);
    expect(back.plotArea.additionalCatAx?.[0]).toMatchObject({ axId: SECONDARY_CAT_AX_ID, crossAx: SECONDARY_VAL_AX_ID, delete: true });
  });

  it('leaves a single-group chart without the additional fields', () => {
    const space = comboSpace();
    const single = makeChartSpace({ plotArea: { chart: space.plotArea.chart } });
    const back = parseChartXml(chartToBytes(single));
    expect(back.plotArea).not.toHaveProperty('additionalCharts');
    expect(back.plotArea).not.toHaveProperty('additionalCatAx');
    expect(back.plotArea).not.toHaveProperty('additionalValAx');
  });

  it('refuses an additional chart group whose axis the plot area does not define', () => {
    const space = comboSpace();
    const dangling: ChartSpace = { ...space, plotArea: { ...space.plotArea, additionalValAx: [] } };
    expect(() => chartToBytes(dangling)).toThrow(OpenXmlSchemaError);
    expect(() => chartToBytes(dangling)).toThrow(/axis 200/);
  });

  it('refuses a pie as an additional chart group, since it has no axes', () => {
    const space = comboSpace();
    const withPie: ChartSpace = { ...space, plotArea: { ...space.plotArea, additionalCharts: [makePieChart({})] } };
    expect(() => chartToBytes(withPie)).toThrow(OpenXmlSchemaError);
  });

  it('produces a schema-valid workbook that loads back as a combo chart', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    ws.drawing = makeDrawing([
      makeChartDrawingItem(makeTwoCellAnchor({ from: 'E2', to: 'L18' }), { space: comboSpace() }),
    ]);
    const bytes = await workbookToBytes(wb);
    expect((await validateXlsx(bytes)).issues).toEqual([]);

    const loaded = await loadWorkbook(fromBuffer(bytes));
    const ref = loaded.sheets[0];
    if (ref?.kind !== 'worksheet') throw new Error('expected worksheet');
    const content = ref.sheet.drawing?.items[0]?.content;
    if (content?.kind !== 'chart') throw new Error('expected chart');
    expect(content.chart.space?.plotArea.additionalCharts?.map((c) => c.kind)).toEqual(['line']);
    expect(content.chart.space?.legend?.legendEntry).toEqual([{ idx: 1, delete: true }]);
  });
});

describe('legend entries', () => {
  it('writes legendEntry right after legendPos, as ECMA-376 orders it', () => {
    const xml = xmlOf(comboSpace());
    expect(xml).toContain('<c:legend><c:legendPos val="t"/><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry>');
  });

  it('round-trips a deleted entry and a restyled entry', () => {
    const txPr = makeSimpleTextBody('');
    const space: ChartSpace = {
      ...comboSpace(),
      legend: { position: 'b', legendEntry: [{ idx: 0, txPr }, { idx: 1, delete: true }] },
    };
    const back = parseChartXml(chartToBytes(space));
    expect(back.legend?.legendEntry?.map((e) => ('delete' in e ? { idx: e.idx, delete: true } : { idx: e.idx }))).toEqual([
      { idx: 0 },
      { idx: 1, delete: true },
    ]);
  });

  it('drops an entry that neither hides nor restyles its row', () => {
    const xml = xmlOf(comboSpace()).replace('<c:delete val="1"/></c:legendEntry>', '<c:delete val="0"/></c:legendEntry>');
    expect(parseChartXml(xml).legend).not.toHaveProperty('legendEntry');
  });
});

describe('series names', () => {
  it('keeps an empty literal name through a round-trip', () => {
    const space = comboSpace();
    if (space.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    const [first] = space.plotArea.chart.series;
    if (!first) throw new Error('expected a series');
    const named: ChartSpace = {
      ...space,
      plotArea: { ...space.plotArea, chart: { ...space.plotArea.chart, series: [{ ...first, tx: { kind: 'literal', value: '' } }] } },
    };
    const back = parseChartXml(chartToBytes(named));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.chart.series[0]?.tx).toEqual({ kind: 'literal', value: '' });
  });
});
