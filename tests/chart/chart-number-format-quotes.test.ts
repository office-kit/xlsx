import { describe, expect, it } from 'vitest';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart.js';
import { chartToBytes, parseChartXml } from '../../src/chart/chart-xml.js';

// Number formats put literal text in double quotes ("%", "万"), which must survive inside the attribute.
const QUOTED_FORMAT = '[>=10000]0\\.0,"万";[<=-10000]0\\.0,"万";0.0"%"';

describe('chart number formats with quoted text', () => {
  it('round-trips a quoted literal on an axis and on data labels', () => {
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({
          series: [
            {
              ...makeBarSeries({ idx: 0, val: { ref: 'Sheet1!$B$2:$B$4' } }),
              dLbls: { showVal: true, numFmt: { formatCode: QUOTED_FORMAT, sourceLinked: false } },
            },
          ],
        }),
        valAx: { axId: 2, crossAx: 1, numFmt: { formatCode: QUOTED_FORMAT, sourceLinked: false } },
      },
    });
    const back = parseChartXml(chartToBytes(space));
    if (back.plotArea.chart.kind !== 'bar') throw new Error('expected bar');
    expect(back.plotArea.valAx?.numFmt?.formatCode).toBe(QUOTED_FORMAT);
    expect(back.plotArea.chart.series[0]?.dLbls?.numFmt?.formatCode).toBe(QUOTED_FORMAT);
  });
});
