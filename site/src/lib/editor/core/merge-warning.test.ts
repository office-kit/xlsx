import { getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

it('asks before merging cells that would lose values, then discards them', () => {
  const ctl = new EditorController();
  type(ctl, 1, 1, 'keep');
  type(ctl, 1, 2, 'lose');
  ctl.selectRange({ r1: 1, c1: 1, r2: 1, c2: 2 });
  A.merge(ctl, 'mergeCenter');
  expect(ctl.doc.ws.mergedCells).toHaveLength(0);
  const props = ctl.dialog?.props;
  expect(props?.['message']).toBe('mergeDiscardsValues');
  const confirm = props?.['onConfirm'];
  expect(typeof confirm).toBe('function');
  if (typeof confirm === 'function') confirm();
  expect(ctl.doc.ws.mergedCells).toHaveLength(1);
  expect(getCell(ctl.doc.ws, 1, 1)?.value).toBe('keep');
  expect(getCell(ctl.doc.ws, 1, 2)?.value ?? null).toBeNull();
  ctl.doc.undo();
  expect(getCell(ctl.doc.ws, 1, 2)?.value).toBe('lose');
});

it('merges without asking when only one cell has a value', () => {
  const ctl = new EditorController();
  type(ctl, 1, 2, 'only');
  ctl.selectRange({ r1: 1, c1: 1, r2: 2, c2: 2 });
  A.merge(ctl, 'merge');
  expect(ctl.dialog).toBeNull();
  expect(getCell(ctl.doc.ws, 1, 1)?.value).toBe('only');
  expect(getCell(ctl.doc.ws, 1, 2)?.value ?? null).toBeNull();
});

it('Merge Across asks only when one row would lose a value', () => {
  const ctl = new EditorController();
  type(ctl, 1, 1, 'a');
  type(ctl, 2, 1, 'b');
  ctl.selectRange({ r1: 1, c1: 1, r2: 2, c2: 2 });
  A.merge(ctl, 'mergeAcross');
  expect(ctl.dialog).toBeNull();
  expect(ctl.doc.ws.mergedCells).toHaveLength(2);
});
