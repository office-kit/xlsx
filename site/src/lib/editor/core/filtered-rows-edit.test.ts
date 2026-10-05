import { getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';
import { filterOwnerAt, setColumnFilter } from './filter.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

/** Item / Qty list filtered to "ap*": rows 3 and 5 are hidden. */
function filtered(): EditorController {
  const ctl = new EditorController();
  [['Item', 'Qty'], ['apple', '5'], ['banana', '20'], ['apricot', '12'], ['cherry', '30']].forEach((r, i) => r.forEach((t, j) => type(ctl, i + 1, j + 1, t)));
  ctl.selectCell({ row: 1, col: 1 });
  ctl.toggleFilter();
  const owner = filterOwnerAt(ctl.doc.ws, 1, 1);
  if (!owner) throw new Error('no filter');
  setColumnFilter(ctl, owner, 1, { kind: 'custom', colId: 0, conditions: [{ val: 'ap*' }] });
  return ctl;
}

const column = (ctl: EditorController, col: number) => [2, 3, 4, 5].map((r) => getCell(ctl.doc.ws, r, col)?.value ?? null);

// Both results were read back from Excel for Mac on the same filtered list.
it('Clear Contents on a filtered list leaves the hidden rows alone', () => {
  const ctl = filtered();
  ctl.selectRange({ r1: 2, c1: 2, r2: 5, c2: 2 });
  A.clear(ctl, 'contents');
  expect(column(ctl, 2)).toEqual([null, 20, null, 30]);
});

it('formatting a filtered list formats only the rows on show', () => {
  const ctl = filtered();
  ctl.selectRange({ r1: 2, c1: 2, r2: 5, c2: 2 });
  A.toggleBold(ctl);
  const bold = [2, 3, 4, 5].map((r) => ctl.doc.styles.get(getCell(ctl.doc.ws, r, 2)?.styleId ?? 0).bold === true);
  expect(bold).toEqual([true, false, true, false]);
});
