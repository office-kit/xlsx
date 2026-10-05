import { getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import { EditorController } from './controller.svelte.ts';
import { replaceAll } from './find.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

const opts = (query: string) => ({ query, matchCase: false, wholeCell: false, byColumns: false, lookIn: 'formulas' as const });

it('replaces nothing when a result is not a valid formula', () => {
  const ctl = new EditorController();
  type(ctl, 1, 1, '=SUM(1,2)');
  type(ctl, 2, 1, '=SUM(3)');
  const result = replaceAll(ctl, opts('3)'), '3', [ctl.doc.ws]);
  expect(result.invalidFormula).toBe('=SUM(3');
  expect(getCell(ctl.doc.ws, 2, 1)?.value).toMatchObject({ formula: 'SUM(3)' });
  expect(ctl.doc.canUndo).toBe(true);
  ctl.doc.undo();
  expect(getCell(ctl.doc.ws, 2, 1)).toBeUndefined();
});

it('reads each result as typed input', () => {
  const ctl = new EditorController();
  type(ctl, 1, 1, 'x10');
  type(ctl, 2, 1, 'x50%');
  type(ctl, 3, 1, 'xTRUE');
  type(ctl, 4, 1, '0y10');
  expect(replaceAll(ctl, opts('x'), '', [ctl.doc.ws]).replaced).toBe(3);
  expect(replaceAll(ctl, opts('y'), 'x', [ctl.doc.ws]).replaced).toBe(1);
  const ws = ctl.doc.ws;
  expect(getCell(ws, 1, 1)?.value).toBe(10);
  expect(getCell(ws, 2, 1)?.value).toBe(0.5);
  expect(ctl.doc.styles.get(getCell(ws, 2, 1)?.styleId ?? 0).numFmt).toBe('0%');
  expect(getCell(ws, 3, 1)?.value).toBe(true);
  expect(getCell(ws, 4, 1)?.value).toBe('0x10');
});
