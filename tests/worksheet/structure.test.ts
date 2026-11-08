import { describe, expect, it } from 'vitest';
import { makeArrayFormula, makeFormula, makeSharedFormula } from '../../src/cell/cell.js';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart.js';
import { addChartAt } from '../../src/drawing/drawing.js';
import { loadWorkbook } from '../../src/io/load.js';
import { registerCellStyle } from '../../src/styles/cell-style.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addDefinedName } from '../../src/workbook/defined-names.js';
import { addWorksheet, createWorkbook, type Workbook } from '../../src/workbook/workbook.js';
import { makeConditionalFormatting, makeCfRule } from '../../src/worksheet/conditional-formatting.js';
import { makeDataValidation } from '../../src/worksheet/data-validations.js';
import { addPivotTable } from '../../src/worksheet/pivot-table.js';
import { makeSparklineGroup } from '../../src/worksheet/sparklines.js';
import { deleteRows, insertRows } from '../../src/worksheet/structure.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import {
  addConditionalFormatting,
  addDataValidation,
  getCell,
  mergeCells,
  setCell,
  setComment,
  setHyperlink,
  setRowHeight,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';
import { validateXlsx } from '../conformance/validate.js';

const formulaAt = (ws: Worksheet, ref: string): string | undefined => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) throw new Error(ref);
  const col = [...(m[1] as string)].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const v = getCell(ws, Number(m[2]), col)?.value;
  return typeof v === 'object' && v !== null && 'kind' in v && v.kind === 'formula' ? v.formula : undefined;
};

const valueAt = (ws: Worksheet, row: number, col: number): unknown => getCell(ws, row, col)?.value;

/** The issue's template: a block of experience rows, a SUM under it, notes and names below. */
const cvTemplate = (): { wb: Workbook; ws: Worksheet } => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'CV');
  const text = registerCellStyle(wb, { font: { bold: true } });
  const days = registerCellStyle(wb, { numberFormat: '0.0' });
  setCell(ws, 9, 1, 'Experience');
  for (let r = 10; r <= 13; r++) {
    setCell(ws, r, 1, `job ${r}`, text);
    setCell(ws, r, 5, r, days);
    mergeCells(ws, `A${r}:D${r}`);
    setRowHeight(ws, r, 30);
  }
  setCell(ws, 14, 5, makeFormula('SUM(E10:E13)'));
  setCell(ws, 16, 1, 'Notes');
  mergeCells(ws, 'A16:E18');
  mergeCells(ws, 'F10:F13');
  addDefinedName(wb, { name: '_xlnm.Print_Area', value: "CV!$A$1:$F$20", scope: 0 });
  addDefinedName(wb, { name: 'Total', value: 'CV!$E$14', scope: 0 });
  return { wb, ws };
};

describe('insertRows', () => {
  it('grows the SUM, merges, print area and names around a block (issue #313)', () => {
    const { wb, ws } = cvTemplate();
    insertRows(wb, ws, 12, 2);

    expect(formulaAt(ws, 'E16')).toBe('SUM(E10:E15)');
    expect(valueAt(ws, 14, 1)).toBe('job 12');
    expect(valueAt(ws, 15, 5)).toBe(13);
    expect(getCell(ws, 12, 1)).toMatchObject({ value: null, styleId: getCell(ws, 11, 1)?.styleId });
    expect(getCell(ws, 13, 5)).toMatchObject({ value: null, styleId: getCell(ws, 11, 5)?.styleId });
    expect(getCell(ws, 12, 1)?.styleId).not.toBe(0);
    expect(ws.rowDimensions.get(12)?.height).toBe(30);
    expect(ws.rowDimensions.get(15)?.height).toBe(30);

    const merges = ws.mergedCells.map((m) => [m.minRow, m.minCol, m.maxRow, m.maxCol]);
    expect(merges).toEqual(
      expect.arrayContaining([
        [10, 1, 10, 4],
        [11, 1, 11, 4],
        [14, 1, 14, 4],
        [15, 1, 15, 4],
        [18, 1, 20, 5],
        [10, 6, 15, 6],
      ]),
    );
    // The inserted rows get no merge of their own.
    expect(merges.some(([r]) => r === 12 || r === 13)).toBe(false);

    expect(wb.definedNames.find((d) => d.name === '_xlnm.Print_Area')?.value).toBe('CV!$A$1:$F$22');
    expect(wb.definedNames.find((d) => d.name === 'Total')?.value).toBe('CV!$E$16');
  });

  it('rewrites references from other sheets, and leaves their own cells alone', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    const other = addWorksheet(wb, 'Summary');
    setCell(ws, 5, 1, 1);
    setCell(other, 5, 1, makeFormula('Data!A5+A5+SUM(Data!$A$1:$A$10)'));
    insertRows(wb, ws, 3);
    expect(formulaAt(other, 'A5')).toBe('Data!A6+A5+SUM(Data!$A$1:$A$11)');
    expect(valueAt(other, 5, 1)).toMatchObject({ kind: 'formula' });
  });

  it('moves validations, formats, hyperlinks, notes and their formulas', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    addDataValidation(ws, makeDataValidation({ type: 'list', sqref: 'B5:B8', formula1: '$H$1:$H$9' }));
    addConditionalFormatting(
      ws,
      makeConditionalFormatting({ sqref: 'C5:C8', rules: [makeCfRule({ type: 'expression', priority: 1, formulas: ['$C5>$D$9'] })] }),
    );
    setHyperlink(ws, 'A7', { target: 'https://example.com' });
    setComment(ws, {
      ref: 'A7',
      author: 'me',
      text: 'note',
      anchor: { fromCol: 1, fromColOffset: 0, fromRow: 5, fromRowOffset: 0, toCol: 3, toColOffset: 0, toRow: 9, toRowOffset: 0 },
    });
    insertRows(wb, ws, 6, 2);

    expect(ws.dataValidations[0]?.sqref.ranges[0]).toMatchObject({ minRow: 5, maxRow: 10 });
    expect(ws.dataValidations[0]?.formula1).toBe('$H$1:$H$11');
    expect(ws.conditionalFormatting[0]?.sqref.ranges[0]).toMatchObject({ minRow: 5, maxRow: 10 });
    expect(ws.conditionalFormatting[0]?.rules[0]?.formulas).toEqual(['$C5>$D$11']);
    expect(ws.hyperlinks[0]?.ref).toBe('A9');
    expect(ws.legacyComments[0]).toMatchObject({ ref: 'A9', anchor: { fromRow: 7, toRow: 11 } });
  });

  it('grows a table the rows land in, and its column formulas follow', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'Name');
    setCell(ws, 1, 2, 'Days');
    for (let r = 2; r <= 4; r++) {
      setCell(ws, r, 1, `n${r}`);
      setCell(ws, r, 2, r);
    }
    addExcelTable(wb, ws, { name: 'Jobs', ref: 'A1:B4', columns: ['Name', 'Days'] });
    insertRows(wb, ws, 3);
    expect(ws.tables[0]?.ref).toBe('A1:B5');
    insertRows(wb, ws, 1);
    expect(ws.tables[0]?.ref).toBe('A2:B6');
  });

  it('moves charts, drawing anchors, sparklines and pivot tables', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const report = addWorksheet(wb, 'R');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 1, 2, 'v');
    for (let r = 2; r <= 5; r++) {
      setCell(ws, r, 1, `k${r}`);
      setCell(ws, r, 2, r);
    }
    const space = makeChartSpace({
      plotArea: {
        chart: makeBarChart({
          series: [makeBarSeries({ idx: 0, val: { ref: 'S!$B$2:$B$5' }, cat: { ref: 'S!$A$2:$A$5' }, tx: { kind: 'ref', ref: 'S!$B$1' } })],
        }),
      },
    });
    const item = addChartAt(ws, 'D8', { space });
    ws.sparklineGroups = [makeSparklineGroup({ sparklines: [{ formula: 'S!B2:B5', location: 'C7' }] })];
    addPivotTable(wb, report, { source: { sheet: 'S', ref: 'A1:B5' }, anchor: 'A3', rows: [0], values: [{ field: 1, aggregate: 'sum' }] });

    insertRows(wb, ws, 3, 2);

    const series = space.plotArea.chart.series[0];
    expect(series && 'val' in series ? series.val.ref : undefined).toBe('S!$B$2:$B$7');
    expect(series && 'cat' in series ? series.cat?.ref : undefined).toBe('S!$A$2:$A$7');
    expect(series?.tx).toEqual({ kind: 'ref', ref: 'S!$B$1' });
    expect(item.anchor.kind === 'oneCell' ? item.anchor.from.row : -1).toBe(9);
    expect(ws.sparklineGroups[0]?.sparklines[0]).toEqual({ formula: 'S!B2:B7', location: 'C9' });
    expect(report.pivotTables?.[0]?.source).toEqual({ sheet: 'S', ref: 'A1:B7' });
  });

  it('writes a shared-formula group on the sheet back as plain formulas', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 2, makeSharedFormula(0, 'A1*2', 'B1:B3'));
    setCell(ws, 2, 2, makeSharedFormula(0));
    setCell(ws, 3, 2, makeSharedFormula(0));
    insertRows(wb, ws, 2);
    expect([1, 3, 4].map((r) => valueAt(ws, r, 2))).toEqual([
      expect.objectContaining({ t: 'normal', formula: 'A1*2' }),
      expect.objectContaining({ t: 'normal', formula: 'A3*2' }),
      expect.objectContaining({ t: 'normal', formula: 'A4*2' }),
    ]);
    for (const r of [1, 3, 4]) expect(valueAt(ws, r, 2)).not.toHaveProperty('si');
  });

  it('moves an array formula that lies wholly below the insertion', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 5, 1, makeArrayFormula('A5:A6', 'B5:B6*2'));
    insertRows(wb, ws, 2);
    expect(valueAt(ws, 6, 1)).toMatchObject({ t: 'array', ref: 'A6:A7', formula: 'B6:B7*2' });
  });

  it('keeps the append cursor after the last row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 3, 1, 'x');
    insertRows(wb, ws, 2, 2);
    expect(ws._appendRowCursor).toBe(5);
  });

  it('round-trips and passes conformance', async () => {
    const { wb } = cvTemplate();
    const ws = wb.sheets[0]?.sheet as Worksheet;
    insertRows(wb, ws, 12, 3);
    const bytes = await workbookToBytes(wb);
    expect(await validateXlsx(bytes)).toMatchObject({ status: 'valid' });
    const back = await loadWorkbook(fromBuffer(bytes));
    const sheet = back.sheets[0]?.sheet as Worksheet;
    expect(formulaAt(sheet, 'E17')).toBe('SUM(E10:E16)');
  });
});

describe('deleteRows', () => {
  it('shrinks the SUM, merges and names, and drops what lay in the deleted rows', () => {
    const { wb, ws } = cvTemplate();
    setCell(ws, 20, 1, makeFormula('E11+E12+E14'));
    deleteRows(wb, ws, 11, 2);
    expect(formulaAt(ws, 'E12')).toBe('SUM(E10:E11)');
    expect(formulaAt(ws, 'A18')).toBe('#REF!+#REF!+E12');
    expect(valueAt(ws, 11, 1)).toBe('job 13');
    const merges = ws.mergedCells.map((m) => [m.minRow, m.minCol, m.maxRow, m.maxCol]);
    expect(merges).toEqual(
      expect.arrayContaining([
        [10, 1, 10, 4],
        [11, 1, 11, 4],
        [10, 6, 11, 6],
        [14, 1, 16, 5],
      ]),
    );
    expect(merges).toHaveLength(4);
    expect(wb.definedNames.find((d) => d.name === 'Total')?.value).toBe('CV!$E$12');
  });

  it('turns a defined name on deleted cells into #REF!', () => {
    const { wb, ws } = cvTemplate();
    deleteRows(wb, ws, 14);
    expect(wb.definedNames.find((d) => d.name === 'Total')?.value).toBe('CV!#REF!');
  });

  it('collapses a merge left with one cell, and drops validations wholly deleted', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    mergeCells(ws, 'A1:A2');
    addDataValidation(ws, makeDataValidation({ type: 'whole', sqref: 'B2 B5', formula1: '0' }));
    addDataValidation(ws, makeDataValidation({ type: 'whole', sqref: 'C2', formula1: '0' }));
    deleteRows(wb, ws, 2);
    expect(ws.mergedCells).toEqual([]);
    expect(ws.dataValidations.map((dv) => dv.sqref.ranges.map((r) => r.minRow))).toEqual([[4]]);
  });
});

describe('rejected edits leave the workbook unchanged', () => {
  const tableBook = (): { wb: Workbook; ws: Worksheet } => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'H');
    setCell(ws, 2, 1, 1);
    setCell(ws, 3, 1, 2);
    setCell(ws, 10, 1, makeFormula('A2'));
    addExcelTable(wb, ws, { name: 'T', ref: 'A1:A3', columns: ['H'] });
    return { wb, ws };
  };

  it.each([
    ['the header row', 1, 1],
    ['every data row', 2, 2],
    ['the whole table', 1, 3],
  ])('deleting %s of a table', (_what, at, count) => {
    const { wb, ws } = tableBook();
    expect(() => deleteRows(wb, ws, at, count)).toThrow(/table "T"/);
    expect(ws.tables[0]?.ref).toBe('A1:A3');
    expect(formulaAt(ws, 'A10')).toBe('A2');
    expect(valueAt(ws, 2, 1)).toBe(1);
  });

  it('cutting through an array formula', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, makeArrayFormula('A1:A3', 'B1:B3'));
    setCell(ws, 5, 2, makeFormula('A1'));
    expect(() => insertRows(wb, ws, 2)).toThrow(/array formula at A1/);
    expect(() => deleteRows(wb, ws, 3)).toThrow(/array formula at A1/);
    expect(formulaAt(ws, 'B5')).toBe('A1');
  });

  it('pushing a populated cell off the grid', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1_048_575, 1, 'x');
    expect(() => insertRows(wb, ws, 5, 2)).toThrow(/pushed past row/);
    expect(valueAt(ws, 1_048_575, 1)).toBe('x');
  });

  it('deleting a pivot table source', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const report = addWorksheet(wb, 'R');
    setCell(ws, 1, 1, 'k');
    setCell(ws, 2, 1, 'a');
    addPivotTable(wb, report, { source: { sheet: 'S', ref: 'A1:A2' }, anchor: 'A3', rows: [0] });
    expect(() => deleteRows(wb, ws, 1, 2)).toThrow(/removePivotTable/);
  });

  it('a formula the tokenizer rejects, before any other change', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 5, 1, 'x');
    setCell(ws, 6, 1, makeFormula('A5'));
    ws.dataValidations.push({ type: 'custom', sqref: { ranges: [{ minRow: 5, minCol: 2, maxRow: 5, maxCol: 2 }] }, formula1: 'A5)' });
    expect(() => insertRows(wb, ws, 1)).toThrow(/data validation/);
    expect(valueAt(ws, 5, 1)).toBe('x');
    expect(formulaAt(ws, 'A6')).toBe('A5');
  });

  it.each([
    ['at', 0, 1],
    ['at', 1.5, 1],
    ['count', 1, 0],
    ['count', 1, -1],
  ])('bad %s', (_what, at, count) => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    expect(() => insertRows(wb, ws, at, count)).toThrow();
    expect(() => deleteRows(wb, ws, at, count)).toThrow();
  });

  it('a sheet from another workbook', () => {
    const ws = addWorksheet(createWorkbook(), 'S');
    expect(() => insertRows(createWorkbook(), ws, 1)).toThrow(/does not belong/);
  });
});

describe('formulas anchored at a range', () => {
  it('re-anchors a rule formula when the top-left row of its range is deleted', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    addConditionalFormatting(
      ws,
      makeConditionalFormatting({ sqref: 'A5:A10', rules: [makeCfRule({ type: 'expression', priority: 1, formulas: ['A5>$B$1'] })] }),
    );
    addDataValidation(ws, makeDataValidation({ type: 'custom', sqref: 'C5:C10', formula1: 'C5<>C4' }));
    deleteRows(wb, ws, 5);
    expect(ws.conditionalFormatting[0]?.sqref.ranges[0]).toMatchObject({ minRow: 5, maxRow: 9 });
    expect(ws.conditionalFormatting[0]?.rules[0]?.formulas).toEqual(['A5>$B$1']);
    expect(ws.dataValidations[0]?.formula1).toBe('C5<>#REF!');
  });

  it('keeps a picture deleted with its rows at the top of the next row', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const item = addChartAt(ws, 'B5', { space: makeChartSpace({ plotArea: { chart: makeBarChart({}) } }) });
    if (item.anchor.kind !== 'oneCell') throw new Error('expected a one-cell anchor');
    item.anchor.from.rowOff = 1000;
    deleteRows(wb, ws, 3, 4);
    expect(item.anchor.from).toMatchObject({ row: 2, rowOff: 0 });
  });
});
