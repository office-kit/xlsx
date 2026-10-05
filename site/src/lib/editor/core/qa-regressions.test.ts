import { describe, expect, test } from 'vitest';
import { getCell } from '@office-kit/xlsx/worksheet';
import * as A from './actions.ts';
import { MAX_COL, parseRangeAddress } from './address.ts';
import { EditorController } from './controller.svelte.ts';
import { currentRange } from './selection.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

const value = (ctl: EditorController, row: number, col: number) => getCell(ctl.doc.ws, row, col)?.value;

describe('editing regressions found against Excel', () => {
  test('Ctrl+D on one cell leaves that cell active, so the next entry does not overwrite the source', () => {
    const ctl = new EditorController();
    type(ctl, 1, 2, 'src');
    ctl.selectCell({ row: 2, col: 2 });
    A.fillFrom(ctl, 'down');
    expect(ctl.doc.selection.active).toEqual({ row: 2, col: 2 });
    type(ctl, ctl.doc.selection.active.row, ctl.doc.selection.active.col, 'Q');
    expect(value(ctl, 1, 2)).toBe('src');
    expect(value(ctl, 2, 2)).toBe('Q');
  });

  test('Enter on a merged cell leaves the merge instead of moving inside it', () => {
    const ctl = new EditorController();
    ctl.selectRange({ r1: 3, c1: 3, r2: 4, c2: 4 });
    A.merge(ctl, 'merge');
    ctl.selectCell({ row: 3, col: 3 });
    ctl.startEdit('m');
    ctl.commitEdit();
    ctl.advance('down');
    expect(ctl.doc.selection.active).toEqual({ row: 5, col: 3 });
    ctl.selectCell({ row: 3, col: 3 });
    ctl.advance('right');
    expect(ctl.doc.selection.active).toEqual({ row: 3, col: 5 });
  });

  test('Enter cycling through a selection skips the hidden cells of a merge', () => {
    const ctl = new EditorController();
    ctl.selectRange({ r1: 3, c1: 3, r2: 4, c2: 4 });
    A.merge(ctl, 'merge');
    ctl.selectRange({ r1: 2, c1: 2, r2: 5, c2: 5 });
    const visited: string[] = [];
    for (let i = 0; i < 16; i++) {
      ctl.advance('down');
      const { row, col } = ctl.doc.selection.active;
      visited.push(`${row},${col}`);
    }
    for (const inside of ['4,3', '3,4', '4,4']) expect(visited).not.toContain(inside);
    expect(visited).toContain('3,3');
  });

  test('Shift+Down after Shift+Space extends whole rows', () => {
    const ctl = new EditorController();
    ctl.selectCell({ row: 3, col: 3 });
    ctl.selectEntire('rows');
    ctl.move(1, 0, { extend: true });
    expect(currentRange(ctl.doc.selection)).toEqual({ r1: 3, r2: 4, c1: 1, c2: MAX_COL });
  });

  test('a column past XFD is not a reference', () => {
    expect(parseRangeAddress('XFE1')).toBeUndefined();
    expect(parseRangeAddress('XFD1')).toEqual({ range: { r1: 1, c1: MAX_COL, r2: 1, c2: MAX_COL } });
  });

  test('F4 repeats the last entry into the active cell', () => {
    const ctl = new EditorController();
    type(ctl, 3, 1, 'rep');
    ctl.selectCell({ row: 4, col: 1 });
    A.redo(ctl);
    expect(value(ctl, 4, 1)).toBe('rep');
  });
});
