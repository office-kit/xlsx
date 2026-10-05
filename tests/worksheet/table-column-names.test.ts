import { describe, expect, it } from 'vitest';
import { parseTableXml, tableToBytes } from '../../src/worksheet/table-xml.js';

const XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:D4" totalsRowCount="1"><autoFilter ref="A1:D3"/><tableColumns count="4"><tableColumn id="1" name="Amount " totalsRowLabel="Sum_x000a_all"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Line_x000a_Break"/><tableColumn id="4" name="Total" totalsRowFunction="custom"><calculatedColumnFormula>Table1[[#This Row],[Amount]]*2</calculatedColumnFormula><totalsRowFormula>SUM(Table1[Total])/2</totalsRowFormula></tableColumn></tableColumns></table>`;

describe('table column names and formulas', () => {
  it('decodes _xHHHH_ in names and keeps column formulas', () => {
    const def = parseTableXml(XML);
    expect(def.columns.map((c) => c.name)).toEqual(['Amount ', 'Amount', 'Line\nBreak', 'Total']);
    expect(def.columns[0]?.totalsRowLabel).toBe('Sum\nall');
    expect(def.columns[3]).toMatchObject({ calculatedColumnFormula: 'Table1[[#This Row],[Amount]]*2', totalsRowFormula: 'SUM(Table1[Total])/2' });
  });

  it('writes them back the way Excel does', () => {
    const xml = new TextDecoder().decode(tableToBytes(parseTableXml(XML)));
    expect(xml).toContain('name="Line_x000A_Break"');
    expect(xml).toContain('name="Amount "');
    expect(xml).toContain('<calculatedColumnFormula>Table1[[#This Row],[Amount]]*2</calculatedColumnFormula><totalsRowFormula>SUM(Table1[Total])/2</totalsRowFormula>');
    expect(parseTableXml(xml).columns).toEqual(parseTableXml(XML).columns);
  });
});
