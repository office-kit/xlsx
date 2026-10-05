import { getCell, setCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

function grid(): EditorController {
  const ctl = new EditorController();
  setCell(ctl.doc.ws, 1, 1, 1);
  setCell(ctl.doc.ws, 1, 2, 2);
  setCell(ctl.doc.ws, 2, 1, 3);
  setCell(ctl.doc.ws, 2, 2, 4);
  return ctl;
}

const formula = (ctl: EditorController, row: number, col: number) => {
  const v = getCell(ctl.doc.ws, row, col)?.value;
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.formula : v;
};

it('fills an empty last row with the column totals', () => {
  const ctl = grid();
  ctl.selectRange({ r1: 1, c1: 1, r2: 3, c2: 2 });
  A.autoSum(ctl);
  expect([formula(ctl, 3, 1), formula(ctl, 3, 2), formula(ctl, 4, 1)]).toEqual(['SUM(A1:A2)', 'SUM(B1:B2)', undefined]);
});

it('fills an empty last column and row with row, column and grand totals', () => {
  const ctl = grid();
  ctl.selectRange({ r1: 1, c1: 1, r2: 3, c2: 3 });
  A.autoSum(ctl);
  expect([formula(ctl, 1, 3), formula(ctl, 2, 3), formula(ctl, 3, 1), formula(ctl, 3, 2), formula(ctl, 3, 3)]).toEqual(['SUM(A1:B1)', 'SUM(A2:B2)', 'SUM(A1:A2)', 'SUM(B1:B2)', 'SUM(C1:C2)']);
});

it('still totals below a selection with no empty edge', () => {
  const ctl = grid();
  ctl.selectRange({ r1: 1, c1: 1, r2: 2, c2: 2 });
  A.autoSum(ctl);
  expect([formula(ctl, 3, 1), formula(ctl, 3, 2)]).toEqual(['SUM(A1:A2)', 'SUM(B1:B2)']);
});
