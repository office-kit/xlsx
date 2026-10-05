import { makeCfRule, makeConditionalFormatting } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

function setup(): EditorController {
  const ctl = new EditorController();
  const ws = ctl.doc.ws;
  ctl.selectRange({ r1: 1, c1: 1, r2: 2, c2: 2 });
  A.merge(ctl, 'merge');
  ws.conditionalFormatting.push(makeConditionalFormatting({ sqref: 'A5:A9', rules: [makeCfRule({ type: 'cellIs', operator: 'greaterThan', priority: 1, formulas: ['3'] })] }));
  return ctl;
}

const cfRanges = (ctl: EditorController) => ctl.doc.ws.conditionalFormatting.flatMap((cf) => cf.sqref.ranges.map((r) => `${r.minRow}:${r.maxRow}`));

it.each(['formats', 'all'] as const)('Clear %s unmerges and takes conditional formatting off the range, and undoes', (kind) => {
  const ctl = setup();
  ctl.selectRange({ r1: 1, c1: 1, r2: 6, c2: 3 });
  A.clear(ctl, kind);
  expect(ctl.doc.ws.mergedCells).toEqual([]);
  expect(cfRanges(ctl)).toEqual(['7:9']);
  ctl.doc.undo();
  expect(ctl.doc.ws.mergedCells).toHaveLength(1);
  expect(cfRanges(ctl)).toEqual(['5:9']);
});

it('Clear Contents keeps merges and conditional formatting', () => {
  const ctl = setup();
  ctl.selectRange({ r1: 1, c1: 1, r2: 6, c2: 3 });
  A.clear(ctl, 'contents');
  expect(ctl.doc.ws.mergedCells).toHaveLength(1);
  expect(cfRanges(ctl)).toEqual(['5:9']);
});
