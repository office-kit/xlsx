import { addExcelTable } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import { EditorController } from './controller.svelte.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

function sales(): EditorController {
  const ctl = new EditorController();
  [['Region', 'Sales'], ['East', '10'], ['West', '20'], ['North', '30']].forEach((r, i) => r.forEach((t, j) => type(ctl, i + 1, j + 1, t)));
  addExcelTable(ctl.doc.wb, ctl.doc.ws, { name: 'Sales', ref: 'A1:B4', columns: ['Region', 'Sales'] });
  return ctl;
}

const ref = (ctl: EditorController) => ctl.doc.ws.tables[0]?.ref;

it('typing just below a table adds a row to it; undo takes the row back out first', () => {
  const ctl = sales();
  type(ctl, 5, 1, 'South');
  expect(ref(ctl)).toBe('A1:B5');
  ctl.doc.undo();
  expect(ref(ctl)).toBe('A1:B4');
});

it('typing just right of a table adds a column named after the header typed', () => {
  const ctl = sales();
  type(ctl, 1, 3, 'Extra');
  expect(ref(ctl)).toBe('A1:C4');
  expect(ctl.doc.ws.tables[0]?.columns.map((c) => c.name)).toEqual(['Region', 'Sales', 'Extra']);
});

it('typing two rows below leaves the table alone', () => {
  const ctl = sales();
  type(ctl, 6, 1, 'Far');
  expect(ref(ctl)).toBe('A1:B4');
});
