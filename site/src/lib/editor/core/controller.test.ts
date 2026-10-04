import { describe, expect, test } from 'vitest';
import { getCell, setHyperlink } from '@office-kit/xlsx/worksheet';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
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

describe('editing session', () => {
  test('typing values and formulas recalculates dependents incrementally', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '1');
    type(ctl, 2, 1, '2');
    type(ctl, 3, 1, '=A1+A2');
    expect(value(ctl, 3, 1)).toBe(3);
    type(ctl, 1, 1, '10');
    expect(value(ctl, 3, 1)).toBe(12);
  });

  test('undo and redo restore cells and recalculated results', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '5');
    type(ctl, 1, 2, '=A1*2');
    type(ctl, 1, 1, '7');
    expect(value(ctl, 1, 2)).toBe(14);
    A.undo(ctl);
    expect(value(ctl, 1, 1)).toBe(5);
    expect(value(ctl, 1, 2)).toBe(10);
    A.redo(ctl);
    expect(value(ctl, 1, 2)).toBe(14);
  });

  test('inserting a row shifts cells and rewrites references', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '1');
    type(ctl, 2, 1, '2');
    type(ctl, 3, 1, '=SUM(A1:A2)');
    ctl.selectCell({ row: 2, col: 1 });
    ctl.selectEntire('rows');
    A.insertLines(ctl, 'row');
    expect(value(ctl, 3, 1)).toBe(2);
    expect(formula(ctl, 4, 1)).toBe('SUM(A1:A3)');
    expect(value(ctl, 4, 1)).toBe(3);
    A.undo(ctl);
    expect(formula(ctl, 3, 1)).toBe('SUM(A1:A2)');
  });

  test('autofill extends a linear series and translates formulas', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '1');
    type(ctl, 2, 1, '3');
    type(ctl, 1, 2, '=A1*10');
    A.autoFill(ctl, { r1: 1, c1: 1, r2: 2, c2: 1 }, { r1: 1, c1: 1, r2: 5, c2: 1 });
    expect([3, 4, 5].map((r) => value(ctl, r, 1))).toEqual([5, 7, 9]);
    A.autoFill(ctl, { r1: 1, c1: 2, r2: 1, c2: 2 }, { r1: 1, c1: 2, r2: 3, c2: 2 });
    expect(formula(ctl, 3, 2)).toBe('A3*10');
    expect(value(ctl, 3, 2)).toBe(50);
  });

  test('dates typed in the locale order become serials with a date format', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '1/15/2024');
    expect(value(ctl, 1, 1)).toBe(45306);
    expect(ctl.doc.styles.get(getCell(ctl.doc.ws, 1, 1)?.styleId ?? 0).numFmt).not.toBe('General');
  });

  test('invalid formulas are rejected without changing the cell', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '4');
    ctl.selectCell({ row: 1, col: 1 });
    ctl.startEdit('=1+*2');
    expect(ctl.commitEdit()).toBe(false);
    expect(ctl.dialog?.kind).toBe('alert');
    expect(value(ctl, 1, 1)).toBe(4);
  });

  test('missing closing parentheses are added on commit', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '=SUM(1,2');
    expect(formula(ctl, 1, 1)).toBe('SUM(1,2)');
    expect(value(ctl, 1, 1)).toBe(3);
  });
});

describe('Enter after Tab', () => {
  test('returns to the column where the Tab run started', () => {
    const ctl = new EditorController();
    ctl.selectCell({ row: 1, col: 2 });
    for (const text of ['a', 'b']) {
      ctl.startEdit(text);
      ctl.commitEdit();
      ctl.advance('right');
    }
    ctl.startEdit('c');
    ctl.commitEdit();
    ctl.advance('down');
    expect(ctl.doc.selection.active).toEqual({ row: 2, col: 2 });
  });
});

describe('Create from Selection', () => {
  test('a label that reads as a cell reference gets a leading underscore', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, 'Q1');
    type(ctl, 1, 2, 'Total');
    type(ctl, 2, 1, '1');
    type(ctl, 2, 2, '2');
    ctl.selectRange({ r1: 1, c1: 1, r2: 2, c2: 2 });
    A.createNamesFromSelection(ctl);
    expect(ctl.doc.wb.definedNames.map((d) => d.name)).toEqual(['_Q1', 'Total']);
  });
});

describe('Remove Hyperlinks', () => {
  test('drops the link and the linked cell style; Clear Hyperlinks keeps the style', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, 'site');
    type(ctl, 1, 2, 'other');
    ctl.selectRange({ r1: 1, c1: 1, r2: 1, c2: 2 });
    A.format(ctl, { font: { bold: true } });
    const styled = getCell(ctl.doc.ws, 1, 1)?.styleId;
    setHyperlink(ctl.doc.ws, 'A1', { target: 'https://example.com' });

    A.clear(ctl, 'hyperlinks');
    expect(ctl.doc.ws.hyperlinks).toEqual([]);
    expect(getCell(ctl.doc.ws, 1, 1)?.styleId).toBe(styled);
    A.undo(ctl);

    A.clear(ctl, 'removeHyperlinks');
    expect(ctl.doc.ws.hyperlinks).toEqual([]);
    expect(getCell(ctl.doc.ws, 1, 1)?.styleId).toBe(0);
    // B1 had no link, so its formatting stays.
    expect(getCell(ctl.doc.ws, 1, 2)?.styleId).toBe(styled);
    A.undo(ctl);
    expect(ctl.doc.ws.hyperlinks.map((h) => h.ref)).toEqual(['A1']);
    expect(getCell(ctl.doc.ws, 1, 1)?.styleId).toBe(styled);
  });
});

describe('Shrink Text to Fit', () => {
  test('toggles the alignment flag as one undo step', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, 'a long label');
    ctl.selectCell({ row: 1, col: 1 });
    A.toggleShrink(ctl);
    expect(A.activeStyle(ctl).shrink).toBe(true);
    A.undo(ctl);
    expect(A.activeStyle(ctl).shrink).toBe(false);
  });
});

describe('Remove Page Break', () => {
  test('removes only the breaks on the active cell edges and undoes as one step', () => {
    const ctl = new EditorController();
    ctl.selectCell({ row: 5, col: 3 });
    A.insertPageBreak(ctl);
    ctl.selectCell({ row: 9, col: 1 });
    A.insertPageBreak(ctl);
    ctl.selectCell({ row: 5, col: 3 });
    A.removePageBreak(ctl);
    expect(ctl.doc.ws.rowBreaks.map((b) => b.id)).toEqual([8]);
    expect(ctl.doc.ws.colBreaks).toEqual([]);
    A.undo(ctl);
    expect(ctl.doc.ws.rowBreaks.map((b) => b.id)).toEqual([4, 8]);
    expect(ctl.doc.ws.colBreaks.map((b) => b.id)).toEqual([2]);
  });
});
