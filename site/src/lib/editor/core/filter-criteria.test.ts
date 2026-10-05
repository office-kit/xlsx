import { fromArrayBuffer, loadWorkbook } from '@office-kit/xlsx/io';
import { expect, it } from 'vitest';
import { EditorController } from './controller.svelte.ts';
import { filterOwnerAt, reapplyFilters, setColumnFilter } from './filter.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

const hidden = (ctl: EditorController) => [...ctl.doc.ws.rowDimensions].filter(([, d]) => d.hidden).map(([r]) => r).sort((a, b) => a - b);

function sales(): EditorController {
  const ctl = new EditorController();
  [['Item', 'Qty'], ['apple', '5'], ['banana', '20'], ['apricot', '12'], ['cherry', '30']].forEach((r, i) => r.forEach((t, j) => type(ctl, i + 1, j + 1, t)));
  ctl.selectCell({ row: 1, col: 1 });
  ctl.toggleFilter();
  return ctl;
}

it('a custom text filter is saved, survives undo/redo, and reloads', async () => {
  const ctl = sales();
  const owner = filterOwnerAt(ctl.doc.ws, 1, 1);
  if (!owner) throw new Error('no filter');
  setColumnFilter(ctl, owner, 1, { kind: 'custom', colId: 0, conditions: [{ val: 'ap*' }] });
  expect(hidden(ctl)).toEqual([3, 5]);
  ctl.doc.undo();
  expect(hidden(ctl)).toEqual([]);
  ctl.doc.redo();
  expect(hidden(ctl)).toEqual([3, 5]);
  const back = await loadWorkbook(fromArrayBuffer(await ctl.doc.toBytes()));
  const ws = back.sheets[0]?.sheet;
  expect(ws && 'autoFilter' in ws ? ws.autoFilter?.filterColumns : undefined).toEqual([{ kind: 'custom', colId: 0, conditions: [{ val: 'ap*' }] }]);
});

it('number conditions, Top 10 and above average filter like Excel and reapply on new data', () => {
  const ctl = sales();
  const owner = filterOwnerAt(ctl.doc.ws, 1, 2);
  if (!owner) throw new Error('no filter');
  setColumnFilter(ctl, owner, 2, { kind: 'custom', colId: 1, and: true, conditions: [{ operator: 'greaterThanOrEqual', val: '10' }, { operator: 'lessThanOrEqual', val: '20' }] });
  expect(hidden(ctl)).toEqual([2, 5]);
  setColumnFilter(ctl, owner, 2, { kind: 'top10', colId: 1, val: 2 });
  expect(hidden(ctl)).toEqual([2, 4]);
  setColumnFilter(ctl, owner, 2, { kind: 'dynamic', colId: 1, type: 'aboveAverage' });
  expect(hidden(ctl)).toEqual([2, 4]);
  type(ctl, 2, 2, '100');
  reapplyFilters(ctl);
  expect(hidden(ctl)).toEqual([3, 4, 5]);
});
