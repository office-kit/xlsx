import { getCell, mergeCells, setCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';
import { sortRange } from './data.ts';

const values = (ctl: EditorController, col: number, rows: number[]) => rows.map((r) => getCell(ctl.doc.ws, r, col)?.value ?? null);

it('refuses to sort a range whose rows carry different merges', () => {
  const ctl = new EditorController();
  [3, 1, 2].forEach((v, i) => setCell(ctl.doc.ws, i + 1, 1, v));
  mergeCells(ctl.doc.ws, 'B2:B3');
  sortRange(ctl, { r1: 1, c1: 1, r2: 3, c2: 2 }, [{ col: 1, descending: false }], false);
  expect(values(ctl, 1, [1, 2, 3])).toEqual([3, 1, 2]);
  expect(ctl.dialog?.props?.['message']).toBe('mergedCellsSameSize');
});

it('sorts rows that all carry the same one-row merge', () => {
  const ctl = new EditorController();
  [3, 1, 2].forEach((v, i) => {
    setCell(ctl.doc.ws, i + 1, 1, v);
    mergeCells(ctl.doc.ws, `B${i + 1}:C${i + 1}`);
  });
  sortRange(ctl, { r1: 1, c1: 1, r2: 3, c2: 3 }, [{ col: 1, descending: false }], false);
  expect(values(ctl, 1, [1, 2, 3])).toEqual([1, 2, 3]);
});

it('refuses to fill over part of a merged cell', () => {
  const ctl = new EditorController();
  setCell(ctl.doc.ws, 1, 1, 1);
  mergeCells(ctl.doc.ws, 'A3:B3');
  A.autoFill(ctl, { r1: 1, c1: 1, r2: 1, c2: 1 }, { r1: 1, c1: 1, r2: 4, c2: 1 });
  expect(values(ctl, 1, [2, 4])).toEqual([null, null]);
  expect(ctl.dialog?.props?.['message']).toBe('mergedCellsSameSize');
});
