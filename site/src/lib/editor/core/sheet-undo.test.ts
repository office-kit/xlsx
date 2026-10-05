import { getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { captureSelection, pasteSpecial } from './clipboard.ts';
import { EditorController } from './controller.svelte.ts';
import { addName } from './names.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

it('Hide Sheet and Unhide Sheet undo and redo', () => {
  const ctl = new EditorController();
  A.insertSheet(ctl);
  A.setSheetHidden(ctl, 0, true);
  expect(ctl.doc.wb.sheets[0]?.state).toBe('hidden');
  ctl.doc.undo();
  expect(ctl.doc.wb.sheets[0]?.state).toBe('visible');
  ctl.doc.redo();
  expect(ctl.doc.wb.sheets[0]?.state).toBe('hidden');
});

it('renaming a sheet rewrites defined names that point at it', () => {
  const ctl = new EditorController();
  type(ctl, 2, 2, '4');
  expect(addName(ctl, { name: 'Rate', value: 'Sheet1!$B$2' })).toBeUndefined();
  type(ctl, 1, 1, '=Rate*2');
  expect(A.renameSheetAt(ctl, 0, 'Data')).toBeUndefined();
  expect(ctl.doc.wb.definedNames[0]?.value).toBe('Data!$B$2');
  expect(getCell(ctl.doc.ws, 1, 1)?.value).toMatchObject({ cachedValue: 8 });
  ctl.doc.undo();
  expect(ctl.doc.wb.definedNames[0]?.value).toBe('Sheet1!$B$2');
});

it('a cut still moves from its own sheet after sheets are reordered', () => {
  const ctl = new EditorController();
  A.insertSheet(ctl);
  ctl.doc.activateSheet(0);
  type(ctl, 2, 1, 'moved');
  ctl.selectRange({ r1: 2, c1: 1, r2: 2, c2: 1 });
  expect(captureSelection(ctl, true)).not.toBeNull();
  const source = ctl.doc.ws;
  A.moveSheetTo(ctl, 0, 1);
  ctl.doc.activateSheet(ctl.doc.wb.sheets.findIndex((s) => s.sheet === source));
  ctl.selectCell({ row: 5, col: 1 });
  pasteSpecial(ctl, { what: 'all', operation: 'none', skipBlanks: false, transpose: false });
  expect(getCell(source, 5, 1)?.value).toBe('moved');
  expect(getCell(source, 2, 1)).toBeUndefined();
});
