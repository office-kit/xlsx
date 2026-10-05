import { makeArrayFormula } from '@office-kit/xlsx/cell';
import { getCell, setCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

function withArray(): EditorController {
  const ctl = new EditorController();
  const ws = ctl.doc.ws;
  setCell(ws, 1, 1, 1);
  setCell(ws, 2, 1, 2);
  setCell(ws, 1, 2, makeArrayFormula('B1:B2', 'A1:A2*2', { cachedValue: 2 }));
  setCell(ws, 2, 2, 4);
  ctl.doc.calc.invalidateAll();
  ctl.doc.calc.recalculateAll();
  return ctl;
}

it('refuses typing into part of a legacy array', () => {
  const ctl = withArray();
  ctl.selectCell({ row: 2, col: 2 });
  ctl.startEdit('9');
  expect(ctl.commitEdit()).toBe(false);
  expect(ctl.dialog?.props?.['message']).toBe('partOfArray');
  expect(getCell(ctl.doc.ws, 2, 2)?.value).toBe(4);
});

it('refuses clearing part of it but clears the whole array', () => {
  const ctl = withArray();
  ctl.selectCell({ row: 1, col: 2 });
  A.clear(ctl, 'contents');
  expect(ctl.dialog?.props?.['message']).toBe('partOfArray');
  expect(getCell(ctl.doc.ws, 1, 2)?.value).toMatchObject({ t: 'array' });
  ctl.dialog = null;
  ctl.selectRange({ r1: 1, c1: 2, r2: 2, c2: 2 });
  A.clear(ctl, 'contents');
  expect(ctl.dialog).toBeNull();
  expect(getCell(ctl.doc.ws, 2, 2)?.value ?? null).toBeNull();
});

it('keeps the array in step when its inputs change', () => {
  const ctl = withArray();
  ctl.selectCell({ row: 2, col: 1 });
  ctl.startEdit('5');
  expect(ctl.commitEdit()).toBe(true);
  expect(getCell(ctl.doc.ws, 2, 2)?.value).toBe(10);
});
