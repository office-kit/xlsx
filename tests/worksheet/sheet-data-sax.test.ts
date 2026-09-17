// `<sheetData>` is cut out of the worksheet text and walked with SAX while the
// rest of the part keeps its node tree. These cover the shapes that split can
// get wrong: a namespace prefix, an empty or absent element, a `<v>` that is
// not the cell's own, and a body the split could mislocate.

import { describe, expect, it } from 'vitest';
import { getFormulaText } from '../../src/cell/cell.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { parseWorksheetXml } from '../../src/worksheet/reader.js';
import { getCell } from '../../src/worksheet/worksheet.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const read = (xml: string, sharedStrings: string[] = []) =>
  parseWorksheetXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`, 'Data', { sharedStrings });

const sheet = (body: string): string => `<worksheet xmlns="${MAIN_NS}"><sheetData>${body}</sheetData></worksheet>`;

describe('a worksheet written with a namespace prefix', () => {
  it('reads its cells, rows and row dimensions', () => {
    const ws = read(
      `<x:worksheet xmlns:x="${MAIN_NS}">` +
        '<x:sheetData>' +
        '<x:row r="1" ht="30" customHeight="1">' +
        '<x:c r="A1"><x:v>42</x:v></x:c>' +
        '<x:c r="B1" t="str"><x:v>hi</x:v></x:c>' +
        '</x:row>' +
        '</x:sheetData>' +
        '</x:worksheet>',
    );
    expect(getCell(ws, 1, 1)?.value).toBe(42);
    expect(getCell(ws, 1, 2)?.value).toBe('hi');
    expect(ws.rowDimensions.get(1)?.height).toBe(30);
  });

  it('keeps reading the parts outside sheetData', () => {
    const ws = read(
      `<x:worksheet xmlns:x="${MAIN_NS}">` +
        '<x:cols><x:col min="1" max="1" width="18" customWidth="1"/></x:cols>' +
        '<x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData>' +
        '<x:mergeCells count="1"><x:mergeCell ref="C1:D1"/></x:mergeCells>' +
        '</x:worksheet>',
    );
    expect(ws.columnDimensions.get(1)?.width).toBe(18);
    expect(ws.mergedCells).toHaveLength(1);
    expect(getCell(ws, 1, 1)?.value).toBe(1);
  });
});

describe('a sheet with nothing to walk', () => {
  it('accepts a self-closing <sheetData/>', () => {
    const ws = read(
      `<worksheet xmlns="${MAIN_NS}"><sheetData/>` +
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>',
    );
    expect(ws.mergedCells).toHaveLength(1);
    expect(ws.rows.size).toBe(0);
  });

  it('accepts a worksheet with no <sheetData> at all', () => {
    const ws = read(`<worksheet xmlns="${MAIN_NS}"><dimension ref="A1"/></worksheet>`);
    expect(ws.rowDimensions.size).toBe(0);
  });

  it('accepts an empty <sheetData></sheetData>', () => {
    const ws = read(sheet(''));
    expect(ws.rowDimensions.size).toBe(0);
  });
});

describe('a body the split cannot read', () => {
  it('rejects a <sheetData> that is never closed', () => {
    expect(() => read(`<worksheet xmlns="${MAIN_NS}"><sheetData><row r="1"/></worksheet>`)).toThrow(
      OpenXmlSchemaError,
    );
  });

  it('rejects an unclosed element inside the body', () => {
    expect(() => read(sheet('<row r="1"><c r="A1"><v>1</v></row>'))).toThrow(OpenXmlSchemaError);
  });

  it('rejects a stray close tag inside the body', () => {
    expect(() => read(sheet('<row r="1"></c></row>'))).toThrow(OpenXmlSchemaError);
  });

  it('rejects a DTD declaration hidden in the body', () => {
    expect(() => read(sheet('<!DOCTYPE row><row r="1"/>'))).toThrow(OpenXmlSchemaError);
  });
});

describe('the body walk and the node walk stay on their own side of the cut', () => {
  it('ignores a comment inside sheetData', () => {
    const ws = read(sheet('<!-- rows below --><row r="1"><c r="A1"><v>7</v></c></row>'));
    expect(getCell(ws, 1, 1)?.value).toBe(7);
  });

  it('does not take a <v> nested in an <extLst> as the cell value', () => {
    // CT_Cell has no extLst, but a producer that invents one must not be able
    // to overwrite the cell: only a direct child of <c> is its value.
    const ws = read(sheet('<row r="1"><c r="A1"><v>7</v><extLst><ext><v>999</v></ext></extLst></c></row>'));
    expect(getCell(ws, 1, 1)?.value).toBe(7);
  });

  it('does not take a <row> nested below sheetData as a row', () => {
    const ws = read(
      sheet('<row r="1"><c r="A1"><v>1</v></c></row><ignored><row r="9"><c r="A9"><v>9</v></c></row></ignored>'),
    );
    expect(getCell(ws, 1, 1)?.value).toBe(1);
    expect(getCell(ws, 9, 1)).toBeUndefined();
  });

  it('keeps element text that arrives as CDATA', () => {
    const ws = read(sheet('<row r="1"><c r="A1" t="str"><v><![CDATA[a&b]]></v></c></row>'));
    expect(getCell(ws, 1, 1)?.value).toBe('a&b');
  });

  it('decodes entity references in cell text', () => {
    const ws = read(sheet('<row r="1"><c r="A1" t="str"><v>a &amp;&lt;&#65;</v></c></row>'));
    expect(getCell(ws, 1, 1)?.value).toBe('a &<A');
  });

  it('ignores the whitespace of a pretty-printed body', () => {
    const ws = read(sheet('\n  <row r="1">\n    <c r="A1">\n      <v>3</v>\n    </c>\n  </row>\n'));
    expect(getCell(ws, 1, 1)?.value).toBe(3);
  });
});

describe('cell shapes the walk has to keep apart', () => {
  it('reads an inline string with runs as rich text', () => {
    const ws = read(
      sheet(
        '<row r="1"><c r="A1" t="inlineStr"><is>' +
          '<r><rPr><b/><sz val="11"/></rPr><t>bold</t></r><r><t> plain</t></r>' +
          '</is></c></row>',
      ),
    );
    expect(getCell(ws, 1, 1)?.value).toEqual({
      kind: 'rich-text',
      runs: [{ text: 'bold', font: { b: true, sz: 11 } }, { text: ' plain' }],
    });
  });

  it('reads a plain inline string', () => {
    const ws = read(sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>plain</t></is></c></row>'));
    expect(getCell(ws, 1, 1)?.value).toBe('plain');
  });

  it('tells an empty <v/> apart from no <v> on a formula cell', () => {
    // Excel writes `<v/>` for a formula whose cached result is the empty
    // string, and the cut must not turn that into an uncached formula.
    const ws = read(
      sheet('<row r="1"><c r="A1" t="str"><f>B1</f><v/></c><c r="B1" t="str"><f>C1</f></c></row>'),
    );
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'formula', t: 'normal', formula: 'B1', cachedValue: '' });
    expect(getCell(ws, 1, 2)?.value).toEqual({ kind: 'formula', t: 'normal', formula: 'C1' });
  });

  it('expands a shared formula against the origin it recorded', () => {
    const ws = read(
      sheet(
        '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>2</v></c></row>' +
          '<row r="2"><c r="A2"><f t="shared" si="0"/><v>3</v></c></row>',
      ),
    );
    expect(getFormulaText(getCell(ws, 2, 1) ?? { row: 0, col: 0, value: null, styleId: 0 })).toBe('B2+1');
  });

  it('resolves a shared-string index', () => {
    const ws = read(sheet('<row r="1"><c r="A1" t="s"><v>1</v></c></row>'), ['zero', 'one']);
    expect(getCell(ws, 1, 1)?.value).toBe('one');
  });

  it('places a row and a cell that carry no @r', () => {
    const ws = read(sheet('<row><c><v>1</v></c><c r="C7"><v>2</v></c></row>'));
    expect(getCell(ws, 7, 1)?.value).toBe(1);
    expect(getCell(ws, 7, 3)?.value).toBe(2);
  });

  it('keeps a self-closing cell as an empty cell', () => {
    const ws = read(sheet('<row r="1"><c r="A1" s="3"/></row>'));
    const cell = getCell(ws, 1, 1);
    expect(cell?.value).toBeNull();
    expect(cell?.styleId).toBe(3);
  });
});
