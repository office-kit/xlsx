import { describe, expect, test } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

const height = (ctl: EditorController, row: number) => ctl.doc.ws.rowDimensions.get(row)?.height;

describe('automatic row height', () => {
  test('a larger font grows the row, and undo shrinks it back with the font', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, 'Hello');
    expect(height(ctl, 1)).toBeUndefined();
    ctl.selectCell({ row: 1, col: 1 });
    A.setFontSize(ctl, 22);
    expect(height(ctl, 1)).toBe(30);
    expect(ctl.doc.rows.sizeOf(1)).toBe(40);
    A.undo(ctl);
    expect(height(ctl, 1)).toBeUndefined();
  });

  test('the tallest cell sets the row, and returning to the Normal font restores the default height', () => {
    const ctl = new EditorController();
    ctl.selectRange({ r1: 2, c1: 1, r2: 2, c2: 2 });
    A.setFontSize(ctl, 16);
    ctl.selectCell({ row: 2, col: 2 });
    A.setFontSize(ctl, 22);
    expect(height(ctl, 2)).toBe(30);
    ctl.selectRange({ r1: 2, c1: 1, r2: 2, c2: 2 });
    A.setFontSize(ctl, 11);
    expect(height(ctl, 2)).toBeUndefined();
  });

  test('a row whose height was set by hand keeps it', () => {
    const ctl = new EditorController();
    ctl.doc.transact('Row Height', (tx) => {
      tx.sheet(ctl.doc.ws, 'rowDimensions');
      ctl.doc.ws.rowDimensions.set(3, { height: 12, customHeight: true });
    });
    ctl.selectCell({ row: 3, col: 1 });
    A.setFontSize(ctl, 36);
    expect(height(ctl, 3)).toBe(12);
  });

  test('a smaller font never fits a row below the default height', () => {
    const ctl = new EditorController();
    type(ctl, 4, 1, 'x');
    ctl.selectCell({ row: 4, col: 1 });
    A.setFontSize(ctl, 6);
    expect(height(ctl, 4)).toBeUndefined();
  });

  test('typing a line break turns on Wrap Text', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, 'a\nb');
    expect(A.activeStyle(ctl).wrap).toBe(true);
  });
});
