import { getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';
import { addName } from './names.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

function cellValue(ctl: EditorController): unknown {
  return getCell(ctl.doc.ws, 1, 1)?.value;
}

it('turns references to a deleted sheet into #REF! and undo restores them', () => {
  const ctl = new EditorController();
  A.insertSheet(ctl);
  type(ctl, 1, 1, '7');
  expect(addName(ctl, { name: 'Seven', value: 'Sheet2!$A$1' })).toBeUndefined();
  ctl.doc.activateSheet(0);
  type(ctl, 1, 1, '=Sheet2!A1+1');
  type(ctl, 1, 2, '=Seven');
  A.deleteSheet(ctl, 1);
  expect(cellValue(ctl)).toMatchObject({ formula: '#REF!+1', cachedValue: '#REF!', cachedValueType: 'error' });
  expect(ctl.doc.wb.definedNames[0]?.value).toBe('#REF!');
  expect(getCell(ctl.doc.ws, 1, 2)?.value).toMatchObject({ cachedValue: '#REF!' });
  ctl.doc.undo();
  ctl.doc.activateSheet(0);
  expect(cellValue(ctl)).toMatchObject({ formula: 'Sheet2!A1+1', cachedValue: 8 });
  expect(ctl.doc.wb.definedNames[0]?.value).toBe('Sheet2!$A$1');
});
