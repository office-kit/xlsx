import { fromArrayBuffer, loadWorkbook } from '@office-kit/xlsx/io';
import { addExcelTable, getCell, setHyperlink, setRowDimension } from '@office-kit/xlsx/worksheet';
import { describe, expect, it } from 'vitest';
import * as A from './actions.ts';
import { compileCriteria, criterionTest, matches, runAdvancedFilter } from './advanced-filter.ts';
import { consolidateGrids } from './consolidate.ts';
import { EditorController } from './controller.svelte.ts';
import { removeDuplicates } from './data.ts';
import { setTotalRow } from './tables.ts';
import { applySubtotals, removeSubtotals } from './subtotal.ts';
import { changingCells, fillDataTable, goalSeek, saveScenario, scenarioSummary, showScenario } from './what-if.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

function grid(ctl: EditorController, rows: string[][], top = 1, left = 1): void {
  rows.forEach((r, i) => r.forEach((text, j) => text !== '' && type(ctl, top + i, left + j, text)));
}

function value(ctl: EditorController, row: number, col: number): unknown {
  const v = getCell(ctl.doc.ws, row, col)?.value;
  if (v && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula') return v.cachedValue;
  return v;
}

function formula(ctl: EditorController, row: number, col: number): string | undefined {
  const v = getCell(ctl.doc.ws, row, col)?.value;
  return v && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.formula : undefined;
}

const SALES = [
  ['Region', 'Sales'],
  ['East', '10'],
  ['East', '20'],
  ['West', '5'],
  ['West', '7'],
  ['West', '1'],
];

describe('subtotals', () => {
  const labels = { group: (v: string) => `${v} Total`, grand: 'Grand Total' };

  it('inserts a SUBTOTAL row per group, a grand total and the outline', () => {
    const ctl = new EditorController();
    grid(ctl, SALES);
    const list = applySubtotals(ctl, { r1: 1, c1: 1, r2: 6, c2: 2 }, { groupBy: 1, fn: 'sum', columns: [2], replace: true, pageBreaks: false, summaryBelow: true }, labels);
    expect(list).toEqual({ r1: 1, c1: 1, r2: 9, c2: 2 });
    expect(value(ctl, 4, 1)).toBe('East Total');
    expect(formula(ctl, 4, 2)).toBe('SUBTOTAL(9,B2:B3)');
    expect(value(ctl, 4, 2)).toBe(30);
    expect(formula(ctl, 8, 2)).toBe('SUBTOTAL(9,B5:B7)');
    expect(value(ctl, 9, 1)).toBe('Grand Total');
    expect(value(ctl, 9, 2)).toBe(43);
    expect(ctl.doc.ws.rowDimensions.get(2)?.outlineLevel).toBe(2);
    expect(ctl.doc.ws.rowDimensions.get(4)?.outlineLevel).toBe(1);
    expect(ctl.doc.ws.rowDimensions.get(9)?.outlineLevel).toBeUndefined();
    expect(ctl.doc.styles.get(getCell(ctl.doc.ws, 4, 1)?.styleId ?? 0).bold).toBe(true);

    const back = removeSubtotals(ctl, list);
    expect(back.r2).toBe(6);
    expect(value(ctl, 4, 1)).toBe('West');
    expect(ctl.doc.ws.rowDimensions.get(2)?.outlineLevel).toBeUndefined();
    A.undo(ctl);
    expect(value(ctl, 4, 1)).toBe('East Total');
    A.undo(ctl);
    expect(value(ctl, 4, 1)).toBe('West');
  });

  it('puts summaries above the data when asked', () => {
    const ctl = new EditorController();
    grid(ctl, SALES);
    applySubtotals(ctl, { r1: 1, c1: 1, r2: 6, c2: 2 }, { groupBy: 1, fn: 'count', columns: [2], replace: true, pageBreaks: false, summaryBelow: false }, labels);
    expect(value(ctl, 2, 1)).toBe('Grand Total');
    expect(value(ctl, 3, 1)).toBe('East Total');
    expect(formula(ctl, 3, 2)).toBe('SUBTOTAL(3,B4:B5)');
    expect(ctl.doc.ws.sheetProperties?.outlinePr?.summaryBelow).toBe(false);
  });
});

describe('consolidate', () => {
  it('matches by labels across sources', () => {
    const a = [
      [null, 'Q1', 'Q2'],
      ['East', 1, 2],
      ['West', 3, 4],
    ];
    const b = [
      [null, 'Q2', 'Q3'],
      ['west', 10, 20],
      ['North', 5, 6],
    ];
    expect(consolidateGrids([a, b], 'sum', true, true)).toEqual([
      [null, 'Q1', 'Q2', 'Q3'],
      ['East', 1, 2, null],
      ['West', 3, 14, 20],
      ['North', null, 5, 6],
    ]);
  });

  it('combines by position without labels', () => {
    expect(consolidateGrids([[[1, 2]], [[3, 'x', 5]]], 'average', false, false)).toEqual([[2, 2, 5]]);
    expect(consolidateGrids([[[1, 2]], [[3, 'x']]], 'count', false, false)).toEqual([[2, 2]]);
  });
});

describe('advanced filter criteria', () => {
  it('reads Excel criteria forms', () => {
    expect(criterionTest('East')?.('Eastern')).toBe(true);
    expect(criterionTest('=East')?.('Eastern')).toBe(false);
    expect(criterionTest('=east')?.('East')).toBe(true);
    expect(criterionTest('>10')?.(11)).toBe(true);
    expect(criterionTest('>10')?.('11')).toBe(false);
    expect(criterionTest('<>West')?.('East')).toBe(true);
    expect(criterionTest('=')?.(null)).toBe(true);
    expect(criterionTest('<>')?.(null)).toBe(false);
    expect(criterionTest('E?st*')?.('Eastern')).toBe(true);
    expect(criterionTest(5)?.(5)).toBe(true);
  });

  it('ORs criteria rows and ANDs their cells', () => {
    const compiled = compileCriteria(['Region', 'Sales'], [
      ['Region', 'Sales'],
      ['East', '>15'],
      ['West', null],
    ]);
    expect(compiled).toBeDefined();
    if (!compiled) return;
    expect(matches(['East', 20], compiled)).toBe(true);
    expect(matches(['East', 10], compiled)).toBe(false);
    expect(matches(['West', 1], compiled)).toBe(true);
    expect(compileCriteria(['Region'], [['Nope'], ['x']])).toBeUndefined();
  });

  it('filters in place and copies unique records elsewhere', () => {
    const ctl = new EditorController();
    grid(ctl, SALES);
    grid(ctl, [['Sales'], ['>6']], 1, 5);
    expect(runAdvancedFilter(ctl, { list: { r1: 1, c1: 1, r2: 6, c2: 2 }, criteria: { r1: 1, c1: 5, r2: 2, c2: 5 }, copyTo: undefined, unique: false })).toBeUndefined();
    expect(ctl.doc.rows.isHidden(4)).toBe(true);
    expect(ctl.doc.rows.isHidden(5)).toBe(false);
    expect(runAdvancedFilter(ctl, { list: { r1: 1, c1: 1, r2: 6, c2: 2 }, criteria: undefined, copyTo: { r1: 10, c1: 1, r2: 10, c2: 1 }, unique: true })).toBeUndefined();
    type(ctl, 20, 1, 'x');
    expect(value(ctl, 10, 1)).toBe('Region');
    expect(value(ctl, 15, 2)).toBe(1);
  });
});

describe('what-if analysis', () => {
  it('goal-seeks a formula input and leaves the model unchanged', () => {
    const ctl = new EditorController();
    grid(ctl, [['3'], ['=A1*A1+1']]);
    const res = goalSeek(ctl, { row: 2, col: 1 }, 10, { row: 1, col: 1 });
    expect(res.found).toBe(true);
    expect(res.value).toBeCloseTo(3, 6);
    expect(value(ctl, 1, 1)).toBe(3);
    const neg = goalSeek(ctl, { row: 2, col: 1 }, 26, { row: 1, col: 1 });
    expect(Math.abs(neg.value)).toBeCloseTo(5, 6);
    expect(goalSeek(ctl, { row: 2, col: 1 }, -5, { row: 1, col: 1 }).found).toBe(false);
  });

  it('fills a two-variable data table with a dataTable formula', () => {
    const ctl = new EditorController();
    // Inputs A1 (rate) and A2 (qty); the table's corner formula multiplies them.
    grid(ctl, [['1'], ['1']]);
    grid(ctl, [['=A1*A2', '2', '3'], ['10', '', ''], ['20', '', '']], 4, 1);
    expect(fillDataTable(ctl, { r1: 4, c1: 1, r2: 6, c2: 3 }, { row: 1, col: 1 }, { row: 2, col: 1 })).toBeUndefined();
    expect(value(ctl, 5, 2)).toBe(20);
    expect(value(ctl, 6, 3)).toBe(60);
    const top = getCell(ctl.doc.ws, 5, 2)?.value;
    expect(top).toMatchObject({ kind: 'formula', t: 'dataTable', ref: 'B5:C6', r1: 'A1', r2: 'A2', dt2D: true });
    expect(value(ctl, 4, 1)).toBe(1);
  });

  it('fills a one-variable column-oriented data table', () => {
    const ctl = new EditorController();
    grid(ctl, [['2']]);
    grid(ctl, [['', '=A1*10'], ['1', ''], ['5', '']], 3, 1);
    expect(fillDataTable(ctl, { r1: 3, c1: 1, r2: 5, c2: 2 }, undefined, { row: 1, col: 1 })).toBeUndefined();
    expect(value(ctl, 4, 2)).toBe(10);
    expect(value(ctl, 5, 2)).toBe(50);
  });

  it('stores, shows and summarises scenarios', () => {
    const ctl = new EditorController();
    grid(ctl, [['1'], ['2'], ['=A1+A2']]);
    expect(changingCells('A1:A2')).toEqual(['A1', 'A2']);
    expect(changingCells('nope')).toBe('invalidReference');
    saveScenario(ctl, { name: 'Best', cells: 'A1:A2', comment: '', locked: true, hidden: false }, ['A1', 'A2'], ['10', '20']);
    expect(ctl.doc.ws.scenarios?.scenarios[0]?.inputCells).toEqual([
      { ref: 'A1', val: '10' },
      { ref: 'A2', val: '20' },
    ]);
    showScenario(ctl, 0);
    expect(value(ctl, 3, 1)).toBe(30);
    A.undo(ctl);
    expect(value(ctl, 3, 1)).toBe(3);
    scenarioSummary(ctl, ['A3'], { title: 'Scenario Summary', current: 'Current Values:', changing: 'Changing Cells:', result: 'Result Cells:', sheetName: 'Scenario Summary' });
    const sheet = ctl.doc.wb.sheets[0];
    expect(sheet?.sheet.title).toBe('Scenario Summary');
    if (sheet?.kind !== 'worksheet') return;
    expect(getCell(sheet.sheet, 3, 5)?.value).toBe('Best');
    expect(getCell(sheet.sheet, 8, 4)?.value).toBe(3);
    expect(getCell(sheet.sheet, 8, 5)?.value).toBe(30);
  });
});

describe('Remove Duplicates', () => {
  // Region / Qty with East and West repeated: rows 4 and 5 are the duplicates on Region.
  function sales(ctl: EditorController): void {
    grid(ctl, [['Region', 'Qty'], ['East', '1'], ['West', '2'], ['East', '3'], ['West', '2'], ['North', '5']]);
  }
  const column = (ctl: EditorController, col: number, r1: number, r2: number) =>
    Array.from({ length: r2 - r1 + 1 }, (_, i) => value(ctl, r1 + i, col) ?? null);

  it('keeps the first occurrence, removes whole rows of the range, and moves nothing outside it', () => {
    const ctl = new EditorController();
    sales(ctl);
    type(ctl, 3, 3, 'right');
    type(ctl, 8, 1, 'below');
    expect(removeDuplicates(ctl, { r1: 1, c1: 1, r2: 6, c2: 2 }, [1], true)).toBe(2);
    expect(column(ctl, 1, 1, 6)).toEqual(['Region', 'East', 'West', 'North', null, null]);
    // Qty of the removed rows goes too, although only Region was compared.
    expect(column(ctl, 2, 1, 6)).toEqual(['Qty', 1, 2, 5, null, null]);
    expect(value(ctl, 3, 3)).toBe('right');
    expect(value(ctl, 8, 1)).toBe('below');
    A.undo(ctl);
    expect(column(ctl, 1, 2, 6)).toEqual(['East', 'West', 'East', 'West', 'North']);
    A.redo(ctl);
    expect(column(ctl, 1, 2, 6)).toEqual(['East', 'West', 'North', null, null]);
  });

  it('compares what the cell shows, so one date in two formats is two values', () => {
    const ctl = new EditorController();
    grid(ctl, [['When'], ['2006-03-08'], ['2006-03-08'], ['2006-03-08']]);
    ctl.selectCell({ row: 3, col: 1 });
    A.setNumberFormat(ctl, 'mmm d, yyyy');
    expect(removeDuplicates(ctl, { r1: 1, c1: 1, r2: 4, c2: 1 }, [1], true)).toBe(1);
    expect(getCell(ctl.doc.ws, 3, 1)?.value).toEqual(getCell(ctl.doc.ws, 2, 1)?.value);
    expect(getCell(ctl.doc.ws, 4, 1)).toBeUndefined();
  });

  it('shrinks a table, keeping its filter and total row in step, and leaves cells below it alone', () => {
    const ctl = new EditorController();
    sales(ctl);
    ctl.doc.transact('Create Table', (tx) => {
      tx.sheet(ctl.doc.ws, 'tables');
      addExcelTable(ctl.doc.wb, ctl.doc.ws, { name: 'Sales', ref: 'A1:B6', columns: ['Region', 'Qty'], autoFilter: { ref: 'A1:B6', filterColumns: [] } });
    });
    const def = ctl.doc.ws.tables[0];
    if (!def) return expect.unreachable('no table');
    setTotalRow(ctl, def, true, 'Total');
    type(ctl, 9, 1, 'below');
    // The whole table is the target even when the selection covered its total row.
    expect(removeDuplicates(ctl, { r1: 1, c1: 1, r2: 7, c2: 2 }, [1], true)).toBe(2);
    const after = ctl.doc.ws.tables[0];
    expect(after?.ref).toBe('A1:B5');
    expect(after?.autoFilter?.ref).toBe('A1:B4');
    expect(value(ctl, 5, 1)).toBe('Total');
    expect(column(ctl, 1, 6, 7)).toEqual([null, null]);
    expect(value(ctl, 9, 1)).toBe('below');
    A.undo(ctl);
    expect(ctl.doc.ws.tables[0]?.ref).toBe('A1:B7');
    expect(value(ctl, 7, 1)).toBe('Total');
  });

  it('moves links with their rows and drops the ones on removed rows', () => {
    const ctl = new EditorController();
    sales(ctl);
    setHyperlink(ctl.doc.ws, 'A4', { target: 'https://example.com/removed' });
    setHyperlink(ctl.doc.ws, 'A6', { target: 'https://example.com/north' });
    setHyperlink(ctl.doc.ws, 'C6', { target: 'https://example.com/outside' });
    removeDuplicates(ctl, { r1: 1, c1: 1, r2: 6, c2: 2 }, [1], true);
    expect(ctl.doc.ws.hyperlinks.map((h) => [h.ref, h.target]).sort()).toEqual([
      ['A4', 'https://example.com/north'],
      ['C6', 'https://example.com/outside'],
    ]);
  });

  it('refuses outlined data before changing anything', () => {
    const ctl = new EditorController();
    sales(ctl);
    setRowDimension(ctl.doc.ws, 3, { outlineLevel: 1 });
    const undoBefore = ctl.doc.history.undoLabel;
    expect(removeDuplicates(ctl, { r1: 1, c1: 1, r2: 6, c2: 2 }, [1], true)).toBe('outlined');
    expect(column(ctl, 1, 2, 6)).toEqual(['East', 'West', 'East', 'West', 'North']);
    expect(ctl.doc.history.undoLabel).toBe(undoBefore);
  });

  it('is kept by save and reopen', async () => {
    const ctl = new EditorController();
    sales(ctl);
    removeDuplicates(ctl, { r1: 1, c1: 1, r2: 6, c2: 2 }, [1], true);
    const wb = await loadWorkbook(fromArrayBuffer(await ctl.doc.toBytes()));
    const ref = wb.sheets[0];
    if (ref?.kind !== 'worksheet') return expect.unreachable('expected a worksheet');
    expect([2, 3, 4, 5].map((r) => getCell(ref.sheet, r, 1)?.value ?? null)).toEqual(['East', 'West', 'North', null]);
  });
});
