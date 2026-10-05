import { getCell, makeDataValidation } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import { EditorController } from './controller.svelte.ts';

function enter(ctl: EditorController, row: number, col: number, text: string): boolean {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  return ctl.commitEdit();
}

it('a whole-number rule refuses error values, typed or computed', () => {
  const ctl = new EditorController();
  ctl.doc.ws.dataValidations.push(makeDataValidation({ type: 'whole', sqref: 'A1', operator: 'between', formula1: '1', formula2: '10' }));
  for (const text of ['#N/A', '=1/0', '=NA()']) {
    expect(enter(ctl, 1, 1, text), text).toBe(false);
    ctl.validationPrompt = null;
    ctl.edit = null;
  }
  expect(enter(ctl, 1, 1, '5')).toBe(true);
});

it('a list taken from a range of dates accepts a typed date', () => {
  const ctl = new EditorController();
  enter(ctl, 1, 2, '1/1/2024');
  enter(ctl, 2, 2, '2/1/2024');
  ctl.doc.ws.dataValidations.push(makeDataValidation({ type: 'list', sqref: 'A1', formula1: '$B$1:$B$2' }));
  expect(enter(ctl, 1, 1, '1/1/2024')).toBe(true);
  expect(getCell(ctl.doc.ws, 1, 1)?.value).toBe(45292);
  expect(enter(ctl, 2, 1, '3/1/2024')).toBe(true); // A2 has no rule
  ctl.doc.ws.dataValidations.push(makeDataValidation({ type: 'list', sqref: 'A3', formula1: '$B$1:$B$2' }));
  expect(enter(ctl, 3, 1, '3/1/2024')).toBe(false);
});
