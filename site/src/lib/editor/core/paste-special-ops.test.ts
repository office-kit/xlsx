import { getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import { captureSelection, pasteSpecial } from './clipboard.ts';
import { EditorController } from './controller.svelte.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

function pasteOp(ctl: EditorController, from: { row: number; col: number }, to: { r1: number; c1: number; r2: number; c2: number }, operation: 'add' | 'multiply'): void {
  ctl.selectCell(from);
  expect(captureSelection(ctl, false)).not.toBeNull();
  ctl.selectRange(to);
  pasteSpecial(ctl, { what: 'all', operation, skipBlanks: false, transpose: false });
}

const value = (ctl: EditorController, row: number, col: number) => getCell(ctl.doc.ws, row, col)?.value;

// Each result was read back from Excel for Mac after the same Paste Special.
it('Add with a blank cell turns numeric text into numbers and leaves other text and formulas working', () => {
  const ctl = new EditorController();
  type(ctl, 1, 3, '3');
  type(ctl, 1, 2, "'5");
  type(ctl, 2, 2, 'abc');
  type(ctl, 4, 2, '=C1');
  pasteOp(ctl, { row: 1, col: 5 }, { r1: 1, c1: 2, r2: 4, c2: 2 }, 'add');
  expect(value(ctl, 1, 2)).toBe(5);
  expect(value(ctl, 2, 2)).toBe('abc');
  expect(value(ctl, 4, 2)).toMatchObject({ kind: 'formula', formula: '(C1)+0' });
});

it('Multiply wraps a formula in the paste area and converts numeric text', () => {
  const ctl = new EditorController();
  type(ctl, 1, 3, '3');
  type(ctl, 1, 6, '2');
  type(ctl, 1, 7, '=C1');
  type(ctl, 2, 7, 'abc');
  type(ctl, 3, 7, "'7");
  pasteOp(ctl, { row: 1, col: 6 }, { r1: 1, c1: 7, r2: 3, c2: 7 }, 'multiply');
  expect(value(ctl, 1, 7)).toMatchObject({ kind: 'formula', formula: '(C1)*2' });
  expect(value(ctl, 2, 7)).toBe('abc');
  expect(value(ctl, 3, 7)).toBe(14);
});

it('a copied formula is added as (formula), shifted like a normal paste', () => {
  const ctl = new EditorController();
  type(ctl, 1, 8, '=C1');
  type(ctl, 1, 9, '4');
  pasteOp(ctl, { row: 1, col: 8 }, { r1: 1, c1: 9, r2: 1, c2: 9 }, 'add');
  expect(value(ctl, 1, 9)).toMatchObject({ kind: 'formula', formula: '4+(D1)' });
});
