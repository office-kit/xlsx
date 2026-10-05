import { expect, it } from 'vitest';
import { EditorController } from './controller.svelte.ts';
import { sortRange } from './data.ts';
import { applySubtotals } from './subtotal.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

it('notes and links move with their rows when sorting, and undo puts them back', () => {
  const ctl = new EditorController();
  ['c', 'a', 'b'].forEach((t, i) => type(ctl, i + 1, 1, t));
  const ws = ctl.doc.ws;
  ws.legacyComments.push({ ref: 'A1', author: 'me', text: 'on c' });
  ws.hyperlinks.push({ ref: 'A2', target: 'https://example.com/a' });
  sortRange(ctl, { r1: 1, c1: 1, r2: 3, c2: 1 }, [{ col: 1, descending: false }], false);
  expect(ws.legacyComments.map((c) => c.ref)).toEqual(['A3']);
  expect(ws.hyperlinks.map((h) => h.ref)).toEqual(['A1']);
  ctl.doc.undo();
  expect(ctl.doc.ws.legacyComments.map((c) => c.ref)).toEqual(['A1']);
  expect(ctl.doc.ws.hyperlinks.map((h) => h.ref)).toEqual(['A2']);
});

it('undoing a whole-sheet step removes the sheet fields it added', () => {
  const ctl = new EditorController();
  [['Region', 'Sales'], ['East', '1'], ['West', '2']].forEach((r, i) => r.forEach((t, j) => type(ctl, i + 1, j + 1, t)));
  const ws = ctl.doc.ws;
  delete ws.sheetProperties;
  applySubtotals(ctl, { r1: 1, c1: 1, r2: 3, c2: 2 }, { groupBy: 1, fn: 'sum', columns: [2], replace: true, pageBreaks: false, summaryBelow: false }, { group: (s: string) => `${s} Total`, grand: 'Grand Total' });
  expect(ctl.doc.ws.sheetProperties?.outlinePr?.summaryBelow).toBe(false);
  ctl.doc.undo();
  expect(ctl.doc.ws.sheetProperties).toBeUndefined();
});
