import { describe, expect, expectTypeOf, it } from 'vitest';
import type { ChartNumberFormat, ValueAxis } from '../../src/chart';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart';
import { chartToBytes } from '../../src/chart/chart-xml';

describe('issue #58 — ChartNumberFormat is publicly importable from @office-kit/xlsx/chart', () => {
  it('lets a caller name the type used by axis numFmt', () => {
    const fmt: ChartNumberFormat = { formatCode: '#,##0', sourceLinked: false };
    const valAx: Partial<ValueAxis> = { numFmt: fmt };
    expectTypeOf<ChartNumberFormat>().toEqualTypeOf<{ formatCode: string; sourceLinked?: boolean }>();
    expect(valAx.numFmt?.formatCode).toBe('#,##0');
  });

  it('round-trips numFmt through the chart serializer', () => {
    const chart = makeBarChart({
      series: [makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$5' } })],
    });
    const space = makeChartSpace({
      plotArea: {
        chart,
        catAx: { axId: 1, crossAx: 2 },
        valAx: {
          axId: 2,
          crossAx: 1,
          numFmt: { formatCode: '#,##0', sourceLinked: false } satisfies ChartNumberFormat,
        },
      },
    });
    const xml = new TextDecoder().decode(chartToBytes(space));
    expect(xml).toContain('<c:numFmt formatCode="#,##0" sourceLinked="0"/>');
  });
});
