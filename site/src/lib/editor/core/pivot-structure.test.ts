import { expect, test } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';
import { createPivotTable } from './pivot.ts';

const DATA = [
  ['Region', 'Sales'],
  ['East', '100'],
  ['West', '150'],
];

function setup(): EditorController {
  const ctl = new EditorController();
  DATA.forEach((row, r) =>
    row.forEach((text, c) => {
      ctl.selectCell({ row: r + 1, col: c + 1 });
      ctl.startEdit(text);
      ctl.commitEdit();
    }),
  );
  expect(createPivotTable(ctl, { source: { sheet: 'Sheet1', ref: 'A1:B3' } })).toBeUndefined();
  return ctl;
}

const source = (ctl: EditorController) => ctl.doc.wb.sheets.flatMap((s) => (s.kind === 'worksheet' ? (s.sheet.pivotTables ?? []) : [])).map((pt) => pt.source.ref);

test('inserting rows on the source sheet moves the source, saves, and undoes', async () => {
  const ctl = setup();
  ctl.doc.activateSheet(ctl.doc.wb.sheets.findIndex((s) => s.sheet.title === 'Sheet1'));
  ctl.selectRange({ r1: 1, c1: 1, r2: 1, c2: 16_384 });
  A.insertLines(ctl, 'row');
  expect(source(ctl)).toEqual(['A2:B4']);
  await expect(ctl.doc.toBytes()).resolves.toBeInstanceOf(Uint8Array);
  ctl.doc.undo();
  expect(source(ctl)).toEqual(['A1:B3']);
});

test('a row inserted inside the source grows it', () => {
  const ctl = setup();
  ctl.doc.activateSheet(ctl.doc.wb.sheets.findIndex((s) => s.sheet.title === 'Sheet1'));
  ctl.selectRange({ r1: 3, c1: 1, r2: 3, c2: 16_384 });
  A.insertLines(ctl, 'row');
  expect(source(ctl)).toEqual(['A1:B4']);
});

test('deleting the source sheet keeps the report as values and still saves', async () => {
  const ctl = setup();
  A.deleteSheet(ctl, ctl.doc.wb.sheets.findIndex((s) => s.sheet.title === 'Sheet1'));
  expect(source(ctl)).toEqual([]);
  await expect(ctl.doc.toBytes()).resolves.toBeInstanceOf(Uint8Array);
  ctl.doc.undo();
  expect(source(ctl)).toEqual(['A1:B3']);
});
