import { getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import { EditorController } from './controller.svelte.ts';
import { sortRange } from './data.ts';

function column(texts: string[], descending: boolean): string[] {
  const ctl = new EditorController();
  texts.forEach((t, i) => {
    ctl.selectCell({ row: i + 1, col: 1 });
    ctl.startEdit(t);
    expect(ctl.commitEdit()).toBe(true);
  });
  sortRange(ctl, { r1: 1, c1: 1, r2: texts.length, c2: 1 }, [{ col: 1, descending }], false);
  return texts.map((_, i) => {
    const v = getCell(ctl.doc.ws, i + 1, 1)?.value;
    return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? `=${v.formula}` : String(v);
  });
}

// Both orders were read back from Excel for Mac after Data ▸ Sort.
it('accents sort after the plain letter, case ties keep their order', () => {
  expect(column(['é', 'e', 'f', 'E'], false)).toEqual(['e', 'E', 'é', 'f']);
});

it('a formula returning "" sorts as text, not as a blank', () => {
  expect(column(['b', '=""', 'a', '5'], false)).toEqual(['5', '=""', 'a', 'b']);
  expect(column(['b', '=""', 'a', '5'], true)).toEqual(['b', 'a', '=""', '5']);
});
