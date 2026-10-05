import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

it('Merge & Center pressed again on the merged cell unmerges it', () => {
  const ctl = new EditorController();
  ctl.selectRange({ r1: 1, c1: 1, r2: 1, c2: 3 });
  A.toggleMergeCenter(ctl);
  expect(ctl.doc.ws.mergedCells).toHaveLength(1);
  ctl.selectCell({ row: 1, col: 1 });
  A.toggleMergeCenter(ctl);
  expect(ctl.doc.ws.mergedCells).toHaveLength(0);
});
