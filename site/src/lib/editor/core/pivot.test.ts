import { describe, expect, test } from 'vitest';
import { getCell } from '@office-kit/xlsx/worksheet';
import { EditorController } from './controller.svelte.ts';
import { createPivotTable, editPivot, placeField } from './pivot.ts';

const DATA = [
  ['Region', 'Product', 'Sales'],
  ['East', 'Apple', '100'],
  ['West', 'Apple', '150'],
  ['East', 'Banana', '80'],
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
  return ctl;
}

const text = (ctl: EditorController, row: number, col: number): unknown => getCell(ctl.doc.ws, row, col)?.value;

describe('PivotTables', () => {
  test('Insert PivotTable adds a sheet with an empty pivot and selects it', () => {
    const ctl = setup();
    expect(createPivotTable(ctl, { source: { sheet: 'Sheet1', ref: 'A1:C4' } })).toBeUndefined();
    expect(ctl.doc.ws.title).toBe('Sheet2');
    expect(ctl.doc.wb.sheets[0]?.sheet.title).toBe('Sheet2');
    expect(ctl.activePivot?.kind).toBe('model');
    expect(ctl.doc.ws.pivotTables).toHaveLength(1);
  });

  test('each field change rewrites the report as one undo step', () => {
    const ctl = setup();
    createPivotTable(ctl, { source: { sheet: 'Sheet1', ref: 'A1:C4' } });
    const ws = ctl.doc.ws;
    editPivot(ctl, ws, 0, 'Add', (pt) => placeField(pt, 0, 'rows'));
    editPivot(ctl, ws, 0, 'Add', (pt) => placeField(pt, 2, 'values'));
    expect([text(ctl, 3, 1), text(ctl, 3, 2)]).toEqual(['Row Labels', 'Sum of Sales']);
    expect([text(ctl, 4, 1), text(ctl, 4, 2), text(ctl, 5, 2), text(ctl, 6, 2)]).toEqual(['East', 180, 150, 330]);

    editPivot(ctl, ws, 0, 'Add', (pt) => placeField(pt, 1, 'rows'));
    expect(text(ctl, 5, 1)).toBe('Apple');
    expect(text(ctl, 8, 1)).toBe('Apple');

    ctl.doc.undo();
    expect(text(ctl, 5, 1)).toBe('West');
    expect(text(ctl, 7, 1)).toBeUndefined();
    expect(ctl.doc.ws.pivotTables?.[0]?.rows).toEqual([0]);
    ctl.doc.redo();
    expect(text(ctl, 5, 1)).toBe('Apple');
  });

  test('Refresh picks up edited source data', () => {
    const ctl = setup();
    createPivotTable(ctl, { source: { sheet: 'Sheet1', ref: 'A1:C4' }, destination: { sheet: 'Sheet1', cell: 'E1' } });
    const ws = ctl.doc.ws;
    editPivot(ctl, ws, 0, 'Add', (pt) => placeField(pt, 2, 'values'));
    expect(text(ctl, 2, 5)).toBe(330);
    ctl.selectCell({ row: 2, col: 3 });
    ctl.startEdit('1000');
    ctl.commitEdit();
    expect(text(ctl, 2, 5)).toBe(330);
    editPivot(ctl, ws, 0, 'Refresh', () => {});
    expect(text(ctl, 2, 5)).toBe(1230);
  });

  test('adding a report filter keeps the table where it is, as Excel does', () => {
    const ctl = setup();
    createPivotTable(ctl, { source: { sheet: 'Sheet1', ref: 'A1:C4' } });
    const ws = ctl.doc.ws;
    editPivot(ctl, ws, 0, 'Add', (pt) => placeField(pt, 1, 'rows'));
    editPivot(ctl, ws, 0, 'Add', (pt) => placeField(pt, 0, 'filters'));
    expect(ws.pivotTables?.[0]?.anchor).toBe('A1');
    expect([text(ctl, 1, 1), text(ctl, 1, 2), text(ctl, 3, 1)]).toEqual(['Region', '(All)', 'Row Labels']);
  });

  test('a source without a header row is refused', () => {
    const ctl = setup();
    expect(createPivotTable(ctl, { source: { sheet: 'Sheet1', ref: 'A1:C1' } })).toBe('pvInvalidSource');
    expect(createPivotTable(ctl, { source: { sheet: 'Sheet1', ref: 'A1:D4' } })).toBe('pvInvalidSource');
  });
});
