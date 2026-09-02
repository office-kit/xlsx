import { describe, expect, it } from 'vitest';
import type { FormulaValue } from '../../src/cell/cell';
import { parseWorksheetXml } from '../../src/worksheet/reader';
import { getCell } from '../../src/worksheet/worksheet';

const wrap = (sheetData: string): string =>
  `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheetData}
</worksheet>`;

describe('parseWorksheetXml — value kinds', () => {
  it('reads number cells (t="n" and unspecified default)', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1"><v>42</v></c><c r="B1" t="n"><v>3.14</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    expect(getCell(ws, 1, 1)?.value).toBe(42);
    expect(getCell(ws, 1, 2)?.value).toBeCloseTo(3.14);
  });

  it('reads boolean cells (t="b")', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    expect(getCell(ws, 1, 1)?.value).toBe(true);
    expect(getCell(ws, 1, 2)?.value).toBe(false);
  });

  it('reads error cells (t="e")', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1" t="e"><v>#REF!</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'error', code: '#REF!' });
  });

  it('reads inline strings (t="inlineStr")', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    expect(getCell(ws, 1, 1)?.value).toBe('Hello');
  });

  it('reads inline rich strings as rich text, one value per <r> run', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1" t="inlineStr">
        <is>
          <r><rPr><b/></rPr><t>foo</t></r>
          <r><t>bar</t></r>
        </is>
      </c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    expect(getCell(ws, 1, 1)?.value).toEqual({
      kind: 'rich-text',
      runs: [{ text: 'foo', font: { b: true } }, { text: 'bar' }],
    });
  });

  it('reads shared-string cells (t="s") via the sst lookup', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>2</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: ['alpha', 'beta', 'gamma'] });
    expect(getCell(ws, 1, 1)?.value).toBe('alpha');
    expect(getCell(ws, 1, 2)?.value).toBe('gamma');
  });

  it('reads "str" formula-result cells', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1" t="str"><v>computed</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    expect(getCell(ws, 1, 1)?.value).toBe('computed');
  });

  it('throws on shared-string index out of range', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1" t="s"><v>5</v></c></row>
    </sheetData>`);
    expect(() => parseWorksheetXml(xml, 'S', { sharedStrings: ['a', 'b'] })).toThrowError(
      /shared-string index 5 out of range/,
    );
  });
});

describe('parseWorksheetXml — formulas', () => {
  it('reads normal formulas with cached value', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1"><f>1+2</f><v>3</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    const v = getCell(ws, 1, 1)?.value as FormulaValue;
    expect(v.kind).toBe('formula');
    expect(v.t).toBe('normal');
    expect(v.formula).toBe('1+2');
    expect(v.cachedValue).toBe(3);
  });

  it('reads array formulas with @ref', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1"><f t="array" ref="A1:A3">SUM(B:B)</f><v>10</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    const v = getCell(ws, 1, 1)?.value as FormulaValue;
    expect(v.t).toBe('array');
    expect(v.ref).toBe('A1:A3');
    expect(v.formula).toBe('SUM(B:B)');
    expect(v.cachedValue).toBe(10);
  });

  it('reads shared formulas — origin keeps text, references translate', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A3">B1+C1</f><v>10</v></c></row>
      <row r="2"><c r="A2"><f t="shared" si="0"/><v>20</v></c></row>
      <row r="3"><c r="A3"><f t="shared" si="0"/><v>30</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    const v1 = getCell(ws, 1, 1)?.value as FormulaValue;
    const v2 = getCell(ws, 2, 1)?.value as FormulaValue;
    const v3 = getCell(ws, 3, 1)?.value as FormulaValue;
    expect(v1.t).toBe('shared');
    expect(v1.formula).toBe('B1+C1');
    expect(v2.t).toBe('shared');
    expect(v2.formula).toBe('B2+C2');
    expect(v3.formula).toBe('B3+C3');
    expect(v2.si).toBe(0);
  });

  it('rejects orphaned shared-formula reference', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1"><f t="shared" si="0"/></c></row>
    </sheetData>`);
    expect(() => parseWorksheetXml(xml, 'S', { sharedStrings: [] })).toThrowError(
      /shared.*si="0".*no preceding origin/,
    );
  });
});

describe('parseWorksheetXml — coords + style', () => {
  it('preserves styleId from <c s="N">', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1" s="3"><v>42</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    expect(getCell(ws, 1, 1)?.styleId).toBe(3);
  });

  it('falls back to next-column slot when @r is missing', () => {
    const xml = wrap(`<sheetData>
      <row r="2"><c><v>1</v></c><c><v>2</v></c><c r="E2"><v>5</v></c></row>
    </sheetData>`);
    const ws = parseWorksheetXml(xml, 'S', { sharedStrings: [] });
    expect(getCell(ws, 2, 1)?.value).toBe(1);
    expect(getCell(ws, 2, 2)?.value).toBe(2);
    expect(getCell(ws, 2, 5)?.value).toBe(5);
  });

  it('returns an empty worksheet when <sheetData> is empty', () => {
    const xml = wrap('<sheetData/>');
    const ws = parseWorksheetXml(xml, 'EmptyShape', { sharedStrings: [] });
    expect(ws.title).toBe('EmptyShape');
    expect(ws.rows.size).toBe(0);
  });
});

describe('parseWorksheetXml — error paths', () => {
  it('rejects a non-worksheet root', () => {
    expect(() => parseWorksheetXml('<foo/>', 'S', { sharedStrings: [] })).toThrowError(/root is .*expected worksheet/);
  });

  it('rejects unknown cell type t="???"', () => {
    const xml = wrap(`<sheetData>
      <row r="1"><c r="A1" t="bogus"><v>1</v></c></row>
    </sheetData>`);
    expect(() => parseWorksheetXml(xml, 'S', { sharedStrings: [] })).toThrowError(/unknown cell type/);
  });
});
