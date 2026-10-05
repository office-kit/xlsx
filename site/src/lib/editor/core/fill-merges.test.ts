import { getCell, mergeCells, setCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

const merges = (ctl: EditorController) => ctl.doc.ws.mergedCells.map((m) => `${m.minRow}:${m.minCol}-${m.maxRow}:${m.maxCol}`).sort();

it('filling a merged cell down repeats the merge on every filled row; undo removes them', () => {
  const ctl = new EditorController();
  setCell(ctl.doc.ws, 1, 1, 'x');
  mergeCells(ctl.doc.ws, 'A1:B1');
  A.autoFill(ctl, { r1: 1, c1: 1, r2: 1, c2: 2 }, { r1: 1, c1: 1, r2: 3, c2: 2 });
  expect(merges(ctl)).toEqual(['1:1-1:2', '2:1-2:2', '3:1-3:2']);
  expect(getCell(ctl.doc.ws, 3, 1)?.value).toBe('x');
  ctl.doc.undo();
  expect(merges(ctl)).toEqual(['1:1-1:2']);
});

it('a two-row block with a merge repeats it per block, and a partial last block gets none', () => {
  const ctl = new EditorController();
  setCell(ctl.doc.ws, 1, 1, 'h');
  mergeCells(ctl.doc.ws, 'A1:A2');
  A.autoFill(ctl, { r1: 1, c1: 1, r2: 2, c2: 1 }, { r1: 1, c1: 1, r2: 5, c2: 1 });
  expect(merges(ctl)).toEqual(['1:1-2:1', '3:1-4:1']);
});
