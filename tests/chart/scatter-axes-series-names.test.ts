import { describe, expect, it } from 'vitest';
import { parseChartXml, chartToBytes } from '../../src/chart/chart-xml.js';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

const SCATTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<c:chart><c:plotArea><c:layout/>
<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>
<c:ser><c:idx val="0"/><c:order val="0"/>
<c:tx><c:strRef><c:f>data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>y</c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:xVal><c:numRef><c:f>data!$A$2:$A$4</c:f></c:numRef></c:xVal>
<c:yVal><c:numRef><c:f>data!$B$2:$B$4</c:f></c:numRef></c:yVal>
</c:ser>
<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>Typed name</c:v></c:tx>
<c:yVal><c:numRef><c:f>data!$C$2:$C$4</c:f></c:numRef></c:yVal>
</c:ser>
<c:axId val="11"/><c:axId val="22"/></c:scatterChart>
<c:valAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/><c:max val="50"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="22"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>
<c:valAx><c:axId val="22"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="11"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>
</c:plotArea></c:chart></c:chartSpace>`;

describe('scatter chart axes and series names', () => {
  it('keeps both value axes with their own ids', () => {
    const space = parseChartXml(SCATTER);
    expect(space.plotArea.xValAx?.axId).toBe(11);
    expect(space.plotArea.xValAx?.scaling?.max).toBe(50);
    expect(space.plotArea.valAx?.axId).toBe(22);
    const xml = decode(chartToBytes(space));
    const ids = [...xml.matchAll(/<c:valAx><c:axId val="(\d+)"\/>/g)].map((m) => m[1]);
    expect(ids).toEqual(['11', '22']);
    expect(xml).toContain('<c:max val="50"/>');
  });

  it('reads and writes series names from a reference or typed text', () => {
    const space = parseChartXml(SCATTER);
    const chart = space.plotArea.chart;
    if (chart.kind !== 'scatter') throw new Error('expected scatter');
    expect(chart.series.map((s) => s.tx)).toEqual([
      { kind: 'ref', ref: 'data!$B$1' },
      { kind: 'literal', value: 'Typed name' },
    ]);
    const back = parseChartXml(chartToBytes(space)).plotArea.chart;
    if (back.kind !== 'scatter') throw new Error('expected scatter');
    expect(back.series.map((s) => s.tx)).toEqual(chart.series.map((s) => s.tx));
  });
});
