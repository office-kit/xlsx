import { addExcelTable } from '@office-kit/xlsx/worksheet';
import { strFromU8, unzipSync } from 'fflate';
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

const formulas = (ctl: EditorController, col: number) =>
  [2, 3, 4].map((r) => {
    const v = ctl.doc.ws.rows.get(r)?.get(col)?.value;
    return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.formula : v;
  });

it('a formula typed into an empty table column fills the column; undo takes the fill back', () => {
  const ctl = sales();
  type(ctl, 1, 3, 'Double');
  type(ctl, 2, 3, '=[@Sales]*2');
  const stored = 'Sales[[#This Row],[Sales]]*2';
  expect(formulas(ctl, 3)).toEqual([stored, stored, stored]);
  expect(ctl.doc.ws.tables[0]?.columns[2]?.calculatedColumnFormula).toBe(stored);
  ctl.doc.undo();
  expect(formulas(ctl, 3)).toEqual([stored, undefined, undefined]);
});

it('relative A1 references follow each row', () => {
  const ctl = sales();
  type(ctl, 1, 3, 'Double');
  type(ctl, 3, 3, '=B3*2');
  expect(formulas(ctl, 3)).toEqual(['B2*2', 'B3*2', 'B4*2']);
});

it('a column that already has data is not filled', () => {
  const ctl = sales();
  type(ctl, 2, 2, '=1+1');
  expect(formulas(ctl, 2)).toEqual(['1+1', 20, 30]);
});

it('a this-row reference is saved in the form Excel opens, and edits as [@Col]', async () => {
  const ctl = sales();
  type(ctl, 2, 3, '=[@Sales]+Sales[@Sales]');
  ctl.selectCell({ row: 2, col: 3 });
  ctl.startEdit();
  expect(ctl.edit?.text).toBe('=[@Sales]+[@Sales]');
  ctl.cancelEdit();
  const xml = strFromU8(unzipSync(await ctl.doc.toBytes())['xl/worksheets/sheet1.xml'] ?? new Uint8Array());
  expect(xml).toContain('<f>Sales[[#This Row],[Sales]]+Sales[[#This Row],[Sales]]</f>');
  expect(xml).not.toContain('[@');
});

it('a row added by typing below picks up the calculated column', () => {
  const ctl = sales();
  type(ctl, 1, 3, 'Double');
  type(ctl, 2, 3, '=[@Sales]*2');
  type(ctl, 5, 1, 'South');
  const v = ctl.doc.ws.rows.get(5)?.get(3)?.value;
  expect(v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.formula : v).toBe('Sales[[#This Row],[Sales]]*2');
});
