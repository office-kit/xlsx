import { describe, expect, test } from 'vitest';
import { getCell } from '@office-kit/xlsx/worksheet';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';
import { guessHeader } from './data.ts';
import { filterBySelectedValue } from './filter.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

function cellValue(ctl: EditorController, row: number, col: number): unknown {
  const v = getCell(ctl.doc.ws, row, col)?.value;
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.cachedValue : v;
}

describe('data tools, compared with Excel', () => {
  test('SUBTOTAL 1-11 skips rows a filter hid but not rows hidden by hand', () => {
    const ctl = new EditorController();
    ['Key', 'a', 'b', 'a', 'b'].forEach((v, i) => type(ctl, i + 1, 1, v));
    ['Val', '1', '2', '4', '8'].forEach((v, i) => type(ctl, i + 1, 2, v));
    type(ctl, 7, 2, '=SUBTOTAL(9,B2:B5)');
    type(ctl, 8, 2, '=SUBTOTAL(109,B2:B5)');
    ctl.selectCell({ row: 2, col: 1 });
    filterBySelectedValue(ctl);
    expect(cellValue(ctl, 7, 2)).toBe(5);
    expect(cellValue(ctl, 8, 2)).toBe(5);
  });

  test('hiding rows by hand recalculates SUBTOTAL 101-111 at once', () => {
    const ctl = new EditorController();
    ['1', '2', '3', '4'].forEach((v, i) => type(ctl, i + 1, 1, v));
    type(ctl, 6, 1, '=SUBTOTAL(109,A1:A4)');
    type(ctl, 7, 1, '=SUBTOTAL(9,A1:A4)');
    ctl.selectRange({ r1: 2, c1: 1, r2: 3, c2: 1 });
    A.hideLines(ctl, 'row', true);
    expect(cellValue(ctl, 6, 1)).toBe(5);
    expect(cellValue(ctl, 7, 1)).toBe(10);
    A.hideLines(ctl, 'row', false);
    expect(cellValue(ctl, 6, 1)).toBe(10);
  });

  test('AutoFill continues dates a month or a year apart by months and years', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, '1/15/2024');
    type(ctl, 2, 1, '2/15/2024');
    A.autoFill(ctl, { r1: 1, c1: 1, r2: 2, c2: 1 }, { r1: 1, c1: 1, r2: 5, c2: 1 });
    // 3/15, 4/15, 5/15/2024
    expect([cellValue(ctl, 3, 1), cellValue(ctl, 4, 1), cellValue(ctl, 5, 1)]).toEqual([45366, 45397, 45427]);
    type(ctl, 1, 2, '1/31/2024');
    type(ctl, 2, 2, '2/29/2024');
    // Not the same day of the month: Excel steps by the 29-day gap.
    A.autoFill(ctl, { r1: 1, c1: 2, r2: 2, c2: 2 }, { r1: 1, c1: 2, r2: 3, c2: 2 });
    expect(cellValue(ctl, 3, 2)).toBe(45351 + 29);
    type(ctl, 1, 3, '1/1/2024');
    type(ctl, 2, 3, '1/1/2025');
    A.autoFill(ctl, { r1: 1, c1: 3, r2: 2, c2: 3 }, { r1: 1, c1: 3, r2: 3, c2: 3 });
    expect(cellValue(ctl, 3, 3)).toBe(46023); // 1/1/2026
  });

  test('a blank header cell over a formula column still reads as a header row', () => {
    const ctl = new EditorController();
    type(ctl, 1, 1, 'Name');
    type(ctl, 1, 2, 'Score');
    type(ctl, 2, 1, 'b');
    type(ctl, 2, 2, '2');
    type(ctl, 2, 3, '=B2*2');
    expect(guessHeader(ctl.doc.ws, { r1: 1, c1: 1, r2: 2, c2: 3 })).toBe(true);
  });
});
