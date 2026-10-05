import { addExcelTable, getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { captureSelection, pasteSpecial } from './clipboard.ts';
import { EditorController } from './controller.svelte.ts';
import { goToReference } from './names.ts';
import { setTotalRow } from './tables.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

it('Total Row on a table with data right below it shifts the data and grows the table', () => {
  const ctl = new EditorController();
  [['Item', 'Qty'], ['a', '1'], ['b', '2'], ['below', '9']].forEach((r, i) => r.forEach((t, j) => type(ctl, i + 1, j + 1, t)));
  addExcelTable(ctl.doc.wb, ctl.doc.ws, { name: 'T', ref: 'A1:B3', columns: ['Item', 'Qty'] });
  const def = ctl.doc.ws.tables[0];
  if (!def) throw new Error('no table');
  setTotalRow(ctl, def, true, 'Total');
  const live = ctl.doc.ws.tables[0];
  expect(live?.ref).toBe('A1:B4');
  expect(live?.totalsRowCount).toBe(1);
  expect(getCell(ctl.doc.ws, 4, 1)?.value).toBe('Total');
  expect(getCell(ctl.doc.ws, 5, 1)?.value).toBe('below');
  ctl.doc.undo();
  expect(ctl.doc.ws.tables[0]?.ref).toBe('A1:B3');
  expect(getCell(ctl.doc.ws, 4, 1)?.value).toBe('below');
});

it('the Name Box selects a whole merge, and copying it pastes cleanly', () => {
  const ctl = new EditorController();
  type(ctl, 1, 1, 'm');
  ctl.selectRange({ r1: 1, c1: 1, r2: 1, c2: 2 });
  A.merge(ctl, 'merge');
  expect(goToReference(ctl, 'B1')).toBe(true);
  expect(ctl.doc.selection.ranges[0]).toEqual({ r1: 1, c1: 1, r2: 1, c2: 2 });
  expect(captureSelection(ctl, false)).not.toBeNull();
  ctl.selectCell({ row: 5, col: 1 });
  pasteSpecial(ctl, { what: 'all', operation: 'none', skipBlanks: false, transpose: false });
  expect(ctl.doc.ws.mergedCells).toHaveLength(2);
  expect(getCell(ctl.doc.ws, 5, 1)?.value).toBe('m');
});

it('a step that throws part-way leaves nothing changed', () => {
  const ctl = new EditorController();
  type(ctl, 1, 1, 'keep');
  const ws = ctl.doc.ws;
  expect(() =>
    ctl.doc.transact('Broken', (tx) => {
      tx.cells(ws, { r1: 1, c1: 1, r2: 1, c2: 1 });
      ws.rows.get(1)?.delete(1);
      throw new Error('boom');
    }),
  ).toThrow('boom');
  expect(getCell(ws, 1, 1)?.value).toBe('keep');
});
