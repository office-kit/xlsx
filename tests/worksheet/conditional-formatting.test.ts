import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook, type Workbook } from '../../src/workbook/workbook.js';
import { parseMultiCellRange } from '../../src/worksheet/cell-range.js';
import { makeCfRule, makeConditionalFormatting } from '../../src/worksheet/conditional-formatting.js';
import { addConditionalFormatting, getConditionalFormatting, type Worksheet } from '../../src/worksheet/worksheet.js';
import { addDxf, makeDifferentialStyle } from '../../src/styles/differential.js';

/** Differential formats for the dxfIds below to point at; a save refuses dangling ones. */
function poolDxfs(wb: Workbook, n: number): void {
  for (let i = 0; i < n; i++) addDxf(wb.styles, makeDifferentialStyle({ font: { size: 10 + i } }));
}

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('addConditionalFormatting / getConditionalFormatting', () => {
  it('starts empty and accepts a block', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    expect(getConditionalFormatting(ws).length).toBe(0);
    addConditionalFormatting(
      ws,
      makeConditionalFormatting({
        sqref: parseMultiCellRange('A1:A10'),
        rules: [makeCfRule({ type: 'cellIs', priority: 1, dxfId: 0, operator: 'greaterThan', formulas: ['100'] })],
      }),
    );
    expect(getConditionalFormatting(ws).length).toBe(1);
  });
});

describe('conditional formatting round-trip', () => {
  it('preserves cellIs / expression / containsText / duplicateValues / aboveAverage / top10', async () => {
    const wb = createWorkbook();
    poolDxfs(wb, 6);
    const ws = addWorksheet(wb, 'C');
    addConditionalFormatting(
      ws,
      makeConditionalFormatting({
        sqref: parseMultiCellRange('A1:A10'),
        rules: [
          makeCfRule({ type: 'cellIs', priority: 1, dxfId: 0, operator: 'greaterThan', formulas: ['100'] }),
          makeCfRule({ type: 'expression', priority: 2, dxfId: 1, formulas: ['MOD(ROW(),2)=0'] }),
          makeCfRule({
            type: 'containsText',
            priority: 3,
            dxfId: 2,
            operator: 'containsText',
            text: 'urgent',
            formulas: ['NOT(ISERROR(SEARCH("urgent",A1)))'],
          }),
          makeCfRule({ type: 'duplicateValues', priority: 4, dxfId: 3 }),
          makeCfRule({ type: 'aboveAverage', priority: 5, dxfId: 4, aboveAverage: false, equalAverage: true }),
          makeCfRule({ type: 'top10', priority: 6, dxfId: 5, rank: 5, percent: true, bottom: true }),
        ],
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const cf = getConditionalFormatting(ws2)[0];
    expect(cf?.rules.length).toBe(6);
    expect(cf?.rules[0]?.type).toBe('cellIs');
    expect(cf?.rules[0]?.operator).toBe('greaterThan');
    expect(cf?.rules[0]?.formulas).toEqual(['100']);
    expect(cf?.rules[1]?.type).toBe('expression');
    expect(cf?.rules[1]?.formulas).toEqual(['MOD(ROW(),2)=0']);
    expect(cf?.rules[2]?.text).toBe('urgent');
    expect(cf?.rules[3]?.type).toBe('duplicateValues');
    expect(cf?.rules[4]?.aboveAverage).toBe(false);
    expect(cf?.rules[4]?.equalAverage).toBe(true);
    expect(cf?.rules[5]?.rank).toBe(5);
    expect(cf?.rules[5]?.percent).toBe(true);
    expect(cf?.rules[5]?.bottom).toBe(true);
  });

  it('preserves multi-range sqref + multiple blocks', async () => {
    const wb = createWorkbook();
    poolDxfs(wb, 6);
    const ws = addWorksheet(wb, 'C');
    addConditionalFormatting(
      ws,
      makeConditionalFormatting({
        sqref: parseMultiCellRange('A1:A5 D1:D5'),
        rules: [makeCfRule({ type: 'cellIs', priority: 1, dxfId: 0, operator: 'equal', formulas: ['"YES"'] })],
      }),
    );
    addConditionalFormatting(
      ws,
      makeConditionalFormatting({
        sqref: parseMultiCellRange('B2:C3'),
        rules: [makeCfRule({ type: 'expression', priority: 2, dxfId: 1, formulas: ['B2>0'] })],
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const cfs = getConditionalFormatting(ws2);
    expect(cfs.length).toBe(2);
    expect(cfs[0]?.sqref.ranges.length).toBe(2);
    expect(cfs[1]?.sqref.ranges.length).toBe(1);
  });

  it('round-trips visual rule innerXml verbatim (colorScale)', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'C');
    addConditionalFormatting(
      ws,
      makeConditionalFormatting({
        sqref: parseMultiCellRange('A1:A10'),
        rules: [
          makeCfRule({
            type: 'colorScale',
            priority: 1,
            innerXml:
              '<colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale>',
          }),
        ],
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    const rule = getConditionalFormatting(ws2)[0]?.rules[0];
    expect(rule?.type).toBe('colorScale');
    // The reader collapses nested namespace decls but keeps the structure;
    // assert the key landmarks.
    expect(rule?.innerXml).toContain('cfvo');
    expect(rule?.innerXml).toContain('FFFF0000');
    expect(rule?.innerXml).toContain('FF00FF00');
  });

  it('omits <conditionalFormatting> when none are set', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'NoCF');
    const bytes = await workbookToBytes(wb);
    const txt = new TextDecoder().decode(bytes);
    expect(txt).not.toContain('<conditionalFormatting');
    expect(txt).not.toContain('<cfRule ');
  });

  it('escapes special chars in formula text', async () => {
    const wb = createWorkbook();
    poolDxfs(wb, 6);
    const ws = addWorksheet(wb, 'C');
    addConditionalFormatting(
      ws,
      makeConditionalFormatting({
        sqref: parseMultiCellRange('A1:A5'),
        rules: [makeCfRule({ type: 'expression', priority: 1, dxfId: 0, formulas: ['IF(A1<10,"low & slow","fast")'] })],
      }),
    );
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const cf = getConditionalFormatting(expectSheet(wb2.sheets[0]?.sheet))[0];
    expect(cf?.rules[0]?.formulas[0]).toBe('IF(A1<10,"low & slow","fast")');
  });
});
