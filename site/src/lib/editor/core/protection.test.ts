import { describe, expect, test } from 'vitest';
import { getCell } from '@office-kit/xlsx/worksheet';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

function protect(ctl: EditorController): void {
  ctl.doc.transact('Protect Sheet', (tx) => {
    tx.sheet(ctl.doc.ws, 'sheetProtection');
    ctl.doc.ws.sheetProtection = { sheet: true };
  });
}

describe('sheet protection', () => {
  test('locked cells refuse typing and the edit leaves no trace', () => {
    const ctl = new EditorController();
    ctl.selectCell({ row: 1, col: 1 });
    ctl.startEdit('5');
    ctl.commitEdit();
    protect(ctl);
    ctl.startEdit('9');
    expect(ctl.edit).toBeNull();
    expect(ctl.dialog?.props?.['message']).toBe('protectedCellAlert');
    ctl.closeDialog();
    A.clear(ctl, 'contents');
    expect(getCell(ctl.doc.ws, 1, 1)?.value).toBe(5);
  });

  test('cells unlocked before protecting stay editable', () => {
    const ctl = new EditorController();
    ctl.selectCell({ row: 2, col: 2 });
    A.toggleLocked(ctl);
    protect(ctl);
    ctl.startEdit('7');
    expect(ctl.commitEdit()).toBe(true);
    expect(getCell(ctl.doc.ws, 2, 2)?.value).toBe(7);
  });

  test('a protected sheet can still be renamed', () => {
    const ctl = new EditorController();
    protect(ctl);
    A.renameSheetAt(ctl, ctl.doc.activeSheetIndex, 'Renamed');
    expect(ctl.doc.ws.title).toBe('Renamed');
  });

  test('a protected structure refuses renaming sheets', () => {
    const ctl = new EditorController();
    ctl.doc.wb.workbookProtection = { lockStructure: true };
    const before = ctl.doc.ws.title;
    A.renameSheetAt(ctl, ctl.doc.activeSheetIndex, 'Other');
    expect(ctl.doc.ws.title).toBe(before);
    expect(ctl.dialog?.props?.['message']).toBe('protectedStructureAlert');
  });
});
