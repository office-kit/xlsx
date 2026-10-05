import { describe, expect, test } from 'vitest';
import type { CellValue } from '@office-kit/xlsx/cell';
import { fromArrayBuffer, loadWorkbook } from '@office-kit/xlsx/io';
import { getCell } from '@office-kit/xlsx/worksheet';
import * as A from './actions.ts';
import { setValueAt } from './cells.ts';
import { EditorController } from './controller.svelte.ts';
import { SpreadsheetEditor } from './editor.svelte.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

function storedFormula(v: CellValue | undefined): string | undefined {
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.formula : undefined;
}

async function reopen(ctl: EditorController): Promise<EditorController> {
  const bytes = await ctl.doc.toBytes();
  const wb = await loadWorkbook(fromArrayBuffer(bytes));
  return new EditorController(new SpreadsheetEditor(wb));
}

/** Every stored cell of every sheet, in a form that compares structurally. */
function snapshot(ctl: EditorController): unknown {
  return ctl.doc.wb.sheets.map((ref) => {
    if (ref.kind !== 'worksheet') return ref.kind;
    const cells: Array<[number, number, unknown]> = [];
    for (const [r, row] of ref.sheet.rows) for (const [c, cell] of row) cells.push([r, c, cell.value]);
    return { title: ref.sheet.title, cells };
  });
}

describe('load → edit → save → reopen', () => {
  // Excel shows #NAME? (or offers to repair the file) for a bare XLOOKUP in a file.
  test('functions newer than Excel 2007 are saved with their _xlfn. prefix', async () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, 'a');
    type(ctl, 1, 2, '1');
    type(ctl, 2, 1, '=XLOOKUP("a",A1:A1,B1:B1)');
    type(ctl, 3, 1, '=SORT(B1:B1)');
    type(ctl, 4, 1, '=LET(x,2,x*3)');
    expect(storedFormula(getCell(ctl.doc.ws, 2, 1)?.value)).toBe('_xlfn.XLOOKUP("a",A1:A1,B1:B1)');
    expect(storedFormula(getCell(ctl.doc.ws, 3, 1)?.value)).toBe('_xlfn._xlws.SORT(B1:B1)');
    expect(storedFormula(getCell(ctl.doc.ws, 4, 1)?.value)).toBe('_xlfn.LET(_xlpm.x,2,_xlpm.x*3)');
    expect(getCell(ctl.doc.ws, 2, 1)?.value).toMatchObject({ cachedValue: 1 });

    const back = await reopen(ctl);
    expect(storedFormula(getCell(back.doc.ws, 2, 1)?.value)).toBe('_xlfn.XLOOKUP("a",A1:A1,B1:B1)');
    // The formula bar shows the formula as typed.
    back.selectCell({ row: 2, col: 1 });
    back.startEdit();
    expect(back.edit?.text).toBe('=XLOOKUP("a",A1:A1,B1:B1)');
  });

  test('a formula that is already prefixed is not prefixed twice', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '=_xlfn.XLOOKUP(1,B1:B2,C1:C2)');
    expect(storedFormula(getCell(ctl.doc.ws, 1, 1)?.value)).toBe('_xlfn.XLOOKUP(1,B1:B2,C1:C2)');
  });

  test('saving a reopened workbook again changes nothing', async () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '2');
    type(ctl, 2, 1, '=A1*3');
    type(ctl, 3, 1, '=XLOOKUP(2,A1:A2,A1:A2)');
    A.renameSheetAt(ctl, 0, 'Data');
    const once = await reopen(ctl);
    const twice = await reopen(once);
    expect(snapshot(twice)).toEqual(snapshot(once));
  });
});

describe('a transaction refused by protection', () => {
  test('rolls back the parts it already changed and leaves no undo step', () => {
    const ctl = new EditorController();
    const doc = ctl.doc;
    ctl.selectCell({ row: 2, col: 2 });
    A.toggleLocked(ctl);
    doc.transact('Protect Sheet', (tx) => {
      tx.sheet(doc.ws, 'sheetProtection');
      doc.ws.sheetProtection = { sheet: true };
    });
    const ws = doc.ws;
    const result = doc.transact('Two cells', (tx) => {
      tx.cells(ws, { r1: 2, c1: 2, r2: 2, c2: 2 });
      setValueAt(ws, 2, 2, 7);
      // A1 is locked, so declaring it refuses the whole step.
      tx.cells(ws, { r1: 1, c1: 1, r2: 1, c2: 1 });
      return true;
    });
    expect(result).toBeUndefined();
    expect(getCell(ws, 2, 2)?.value ?? null).toBeNull();
    ctl.closeDialog();
    // The step before the refused one is still the one undo reverts.
    A.undo(ctl);
    expect(ws.sheetProtection).toBeUndefined();
  });
});
