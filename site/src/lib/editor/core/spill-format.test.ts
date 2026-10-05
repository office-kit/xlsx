import { getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

const shown = (ctl: EditorController, row: number): unknown => {
  const v = getCell(ctl.doc.ws, row, 5)?.value;
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.cachedValue : v;
};

it('formatting a spilled cell keeps the spill, before and after undo', () => {
  const ctl = new EditorController();
  ctl.selectCell({ row: 1, col: 5 });
  ctl.startEdit('=SEQUENCE(3)');
  expect(ctl.commitEdit()).toBe(true);
  expect([1, 2, 3].map((r) => shown(ctl, r))).toEqual([1, 2, 3]);
  ctl.selectCell({ row: 2, col: 5 });
  A.toggleBold(ctl);
  expect([1, 2, 3].map((r) => shown(ctl, r))).toEqual([1, 2, 3]);
  ctl.doc.undo();
  expect([1, 2, 3].map((r) => shown(ctl, r))).toEqual([1, 2, 3]);
  ctl.selectCell({ row: 1, col: 1 });
  ctl.startEdit('x');
  expect(ctl.commitEdit()).toBe(true);
  expect([1, 2, 3].map((r) => shown(ctl, r))).toEqual([1, 2, 3]);
});
