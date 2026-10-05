import { addExcelTable, getCell } from '@office-kit/xlsx/worksheet';
import { describe, expect, it } from 'vitest';
import * as A from './actions.ts';
import { captureSelection, pasteSpecial } from './clipboard.ts';
import { EditorController } from './controller.svelte.ts';
import { filterBySelectedValue } from './filter.ts';
import { addName } from './names.ts';

const PASTE_ALL = { what: 'all', operation: 'none', skipBlanks: false, transpose: false } as const;

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
  return v ?? null;
}

function formula(ctl: EditorController, row: number, col: number): string | undefined {
  const v = getCell(ctl.doc.ws, row, col)?.value;
  return v && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.formula : undefined;
}

function copyTo(ctl: EditorController, from: { r1: number; c1: number; r2: number; c2: number }, row: number, col: number, cut = false): void {
  ctl.selectRange(from);
  expect(captureSelection(ctl, cut)).not.toBeNull();
  ctl.selectCell({ row, col });
  pasteSpecial(ctl, PASTE_ALL);
}

const SALES = [
  ['Region', 'Sales'],
  ['East', '10'],
  ['East', '20'],
  ['West', '5'],
  ['West', '7'],
];

describe('cut and paste', () => {
  it('moves a defined name with the cells it names', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '5');
    expect(addName(ctl, { name: 'Price', value: 'Sheet1!$A$1' })).toBeUndefined();
    type(ctl, 1, 2, '=Price*2');
    copyTo(ctl, { r1: 1, c1: 1, r2: 1, c2: 1 }, 3, 1, true);
    expect(ctl.doc.wb.definedNames[0]?.value).toBe('Sheet1!$A$3');
    expect(value(ctl, 1, 2)).toBe(10);
  });

  it('moves a merge inside the cut block instead of leaving it behind', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, 'x');
    ctl.selectRange({ r1: 1, c1: 1, r2: 1, c2: 2 });
    A.merge(ctl, 'merge');
    copyTo(ctl, { r1: 1, c1: 1, r2: 1, c2: 2 }, 4, 1, true);
    expect(ctl.doc.ws.mergedCells.map((m) => [m.minRow, m.minCol, m.maxRow, m.maxCol])).toEqual([[4, 1, 4, 2]]);
  });

  it('refuses to paste over part of a merge', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, 'a');
    type(ctl, 5, 2, 'keep');
    ctl.selectRange({ r1: 5, c1: 2, r2: 6, c2: 3 });
    A.merge(ctl, 'merge');
    copyTo(ctl, { r1: 1, c1: 1, r2: 1, c2: 2 }, 5, 1);
    expect(value(ctl, 5, 1)).toBeNull();
    expect(value(ctl, 5, 2)).toBe('keep');
  });
});

describe('filtered lists', () => {
  function filtered(): EditorController {
    const ctl = new EditorController();
    grid(ctl, SALES);
    ctl.selectCell({ row: 4, col: 1 });
    filterBySelectedValue(ctl);
    expect([2, 3].map((r) => ctl.doc.ws.rowDimensions.get(r)?.hidden)).toEqual([true, true]);
    return ctl;
  }

  it('copies only the rows on show', () => {
    const ctl = filtered();
    copyTo(ctl, { r1: 1, c1: 1, r2: 5, c2: 2 }, 1, 4);
    expect([1, 2, 3, 4].map((r) => value(ctl, r, 5))).toEqual(['Sales', 5, 7, null]);
  });

  it('translates copied formulas from their own source row', () => {
    const ctl = filtered();
    type(ctl, 5, 3, '=B5*2');
    copyTo(ctl, { r1: 2, c1: 3, r2: 5, c2: 3 }, 10, 3);
    expect(formula(ctl, 11, 3)).toBe('B11*2');
  });

  it('Ctrl+Enter fills only the rows on show', () => {
    const ctl = filtered();
    ctl.selectRange({ r1: 2, c1: 3, r2: 5, c2: 3 });
    ctl.startEdit('x');
    expect(ctl.commitEdit({ fillSelection: true })).toBe(true);
    expect([2, 3, 4, 5].map((r) => value(ctl, r, 3))).toEqual([null, null, 'x', 'x']);
  });

  it('Fill Down skips filtered-out rows', () => {
    const ctl = filtered();
    type(ctl, 1, 3, 'top');
    ctl.selectRange({ r1: 1, c1: 3, r2: 5, c2: 3 });
    A.fillFrom(ctl, 'down');
    expect([2, 3, 4, 5].map((r) => value(ctl, r, 3))).toEqual([null, null, 'top', 'top']);
  });
});

describe('table columns', () => {
  function table(): EditorController {
    const ctl = new EditorController();
    grid(ctl, [
      ['Item', 'Qty', 'Price'],
      ['a', '2', '3'],
    ]);
    addExcelTable(ctl.doc.wb, ctl.doc.ws, { name: 'T', ref: 'A1:C2', columns: ['Item', 'Qty', 'Price'] });
    type(ctl, 5, 1, '=SUM(T[Qty])');
    type(ctl, 5, 2, '=SUM(T[Price])');
    return ctl;
  }

  it('renaming a header rewrites references to the column', () => {
    const ctl = table();
    type(ctl, 1, 2, 'Count');
    expect(ctl.doc.ws.tables[0]?.columns.map((c) => c.name)).toEqual(['Item', 'Count', 'Price']);
    expect(formula(ctl, 5, 1)).toBe('SUM(T[Count])');
    expect(value(ctl, 5, 1)).toBe(2);
    ctl.doc.undo();
    expect(formula(ctl, 5, 1)).toBe('SUM(T[Qty])');
    expect(ctl.doc.ws.tables[0]?.columns.map((c) => c.name)).toEqual(['Item', 'Qty', 'Price']);
  });

  it('deleting a column turns references to it into #REF!', () => {
    const ctl = table();
    ctl.selectRange({ r1: 1, c1: 2, r2: 1_048_576, c2: 2 });
    A.deleteLines(ctl, 'col');
    expect(ctl.doc.ws.tables[0]?.columns.map((c) => c.name)).toEqual(['Item', 'Price']);
    expect(formula(ctl, 5, 1)).toBe('SUM(#REF!)');
  });

  it('inserting a column inside the table adds a ColumnN column', () => {
    const ctl = table();
    ctl.selectRange({ r1: 1, c1: 2, r2: 1_048_576, c2: 2 });
    A.insertLines(ctl, 'col');
    expect(ctl.doc.ws.tables[0]?.ref).toBe('A1:D2');
    expect(ctl.doc.ws.tables[0]?.columns.map((c) => c.name)).toEqual(['Item', 'Column1', 'Qty', 'Price']);
  });
});
