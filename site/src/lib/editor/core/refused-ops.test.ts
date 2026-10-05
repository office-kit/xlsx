import { addExcelTable, getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

function withTable(): EditorController {
  const ctl = new EditorController();
  [['Item', 'Qty'], ['a', '1'], ['b', '2']].forEach((r, i) => r.forEach((t, j) => type(ctl, i + 1, j + 1, t)));
  addExcelTable(ctl.doc.wb, ctl.doc.ws, { name: 'T', ref: 'A1:B3', columns: ['Item', 'Qty'] });
  return ctl;
}

it('refuses to merge cells in a table', () => {
  const ctl = withTable();
  ctl.selectRange({ r1: 2, c1: 1, r2: 2, c2: 2 });
  A.merge(ctl, 'merge');
  expect(ctl.dialog?.props?.['message']).toBe('mergeInTable');
  expect(ctl.doc.ws.mergedCells).toHaveLength(0);
});

it('refuses Insert Cells that would shift part of a table, allows a whole-width shift', () => {
  const ctl = withTable();
  ctl.selectRange({ r1: 1, c1: 1, r2: 1, c2: 1 });
  A.shiftCells(ctl, 'down');
  expect(ctl.dialog?.props?.['message']).toBe('shiftTable');
  expect(ctl.doc.ws.tables[0]?.ref).toBe('A1:B3');
  ctl.dialog = null;
  ctl.selectRange({ r1: 1, c1: 1, r2: 1, c2: 2 });
  A.shiftCells(ctl, 'down');
  expect(ctl.dialog).toBeNull();
  expect(ctl.doc.ws.tables[0]?.ref).toBe('A2:B4');
});

it('refuses Delete Cells that would shift part of a merge', () => {
  const ctl = new EditorController();
  type(ctl, 3, 1, 'm');
  ctl.selectRange({ r1: 3, c1: 1, r2: 3, c2: 2 });
  A.merge(ctl, 'merge');
  ctl.selectRange({ r1: 1, c1: 1, r2: 1, c2: 1 });
  A.shiftCells(ctl, 'up');
  expect(ctl.dialog?.props?.['message']).toBe('shiftMerge');
  expect(getCell(ctl.doc.ws, 3, 1)?.value).toBe('m');
});

it('Filter inside a table toggles the table filter buttons instead of adding a sheet filter', () => {
  const ctl = withTable();
  ctl.selectCell({ row: 2, col: 1 });
  const had = ctl.doc.ws.tables[0]?.autoFilter !== undefined;
  ctl.toggleFilter();
  expect(ctl.doc.ws.autoFilter).toBeUndefined();
  expect(ctl.doc.ws.tables[0]?.autoFilter !== undefined).toBe(!had);
});
