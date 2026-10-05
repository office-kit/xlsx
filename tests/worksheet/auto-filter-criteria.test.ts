import { describe, expect, it } from 'vitest';
import { parseAutoFilterNode, serializeAutoFilterXml } from '../../src/worksheet/auto-filter-xml.js';
import { parseTableXml, tableToBytes } from '../../src/worksheet/table-xml.js';
import { parseXml } from '../../src/xml/parser.js';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const parse = (inner: string) => parseAutoFilterNode(parseXml(`<autoFilter ${NS} ref="A1:F9">${inner}</autoFilter>`));

describe('AutoFilter criteria', () => {
  it('reads custom, Top 10 and dynamic filters and writes them back', () => {
    const filter = parse(
      '<filterColumn colId="0"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="10"/><customFilter operator="lessThanOrEqual" val="20"/></customFilters></filterColumn>' +
        '<filterColumn colId="1"><customFilters><customFilter val="ab*"/></customFilters></filterColumn>' +
        '<filterColumn colId="2"><top10 top="0" percent="1" val="5" filterVal="3"/></filterColumn>' +
        '<filterColumn colId="3"><dynamicFilter type="aboveAverage" val="12.5"/></filterColumn>',
    );
    expect(filter?.filterColumns).toEqual([
      { kind: 'custom', colId: 0, and: true, conditions: [{ operator: 'greaterThanOrEqual', val: '10' }, { operator: 'lessThanOrEqual', val: '20' }] },
      { kind: 'custom', colId: 1, conditions: [{ val: 'ab*' }] },
      { kind: 'top10', colId: 2, top: false, percent: true, val: 5, filterVal: 3 },
      { kind: 'dynamic', colId: 3, type: 'aboveAverage', val: 12.5 },
    ]);
    if (!filter) throw new Error('no filter');
    const again = parseAutoFilterNode(parseXml(serializeAutoFilterXml(filter).replace('<autoFilter', `<autoFilter ${NS}`)));
    expect(again).toEqual(filter);
  });

  it('keeps colour filters verbatim', () => {
    const filter = parse('<filterColumn colId="4"><colorFilter dxfId="2"/></filterColumn>');
    expect(filter?.filterColumns[0]).toMatchObject({ kind: 'raw', colId: 4 });
    if (!filter) throw new Error('no filter');
    const xml = serializeAutoFilterXml(filter);
    expect(xml).toMatch(/<filterColumn colId="4"><(?:\w+:)?colorFilter[^>]*dxfId="2"/);
    const again = parseAutoFilterNode(parseXml(xml.replace('<autoFilter', `<autoFilter ${NS}`)));
    expect(again?.filterColumns[0]).toMatchObject({ kind: 'raw', colId: 4 });
  });

  it('keeps the sort Excel nests inside <autoFilter>', () => {
    const filter = parse('<filterColumn colId="0"><filters><filter val="a"/></filters></filterColumn><sortState ref="A2:F9"><sortCondition descending="1" ref="B2:B9"/></sortState>');
    expect(filter?.sortState).toEqual({ ref: 'A2:F9', conditions: [{ ref: 'B2:B9', descending: true }] });
    if (!filter) throw new Error('no filter');
    expect(serializeAutoFilterXml(filter)).toBe(
      '<autoFilter ref="A1:F9"><filterColumn colId="0"><filters><filter val="a"/></filters></filterColumn><sortState ref="A2:F9"><sortCondition descending="1" ref="B2:B9"/></sortState></autoFilter>',
    );
  });
});

describe('table sortState', () => {
  it('survives a table part round-trip', () => {
    const xml = (
      `<table ${NS} id="1" name="T" displayName="T" ref="A1:B5"><autoFilter ref="A1:B5"/><sortState ref="A2:B5"><sortCondition ref="B2:B5"/></sortState><tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns></table>`
    );
    const table = parseTableXml(xml);
    expect(table.sortState).toEqual({ ref: 'A2:B5', conditions: [{ ref: 'B2:B5' }] });
    expect(new TextDecoder().decode(tableToBytes(table))).toContain('<autoFilter ref="A1:B5"/><sortState ref="A2:B5"><sortCondition ref="B2:B5"/></sortState><tableColumns');
  });
});
