import { getCell, setCell } from '@office-kit/xlsx/worksheet';
import { describe, expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';
import { findAll, replaceAll } from './find.ts';

function withValue(value: number, numFmt: string): EditorController {
  const ctl = new EditorController();
  setCell(ctl.doc.ws, 1, 1, value);
  ctl.selectCell({ row: 1, col: 1 });
  A.setNumberFormat(ctl, numFmt);
  return ctl;
}

describe('F2 then Enter keeps the value', () => {
  it.each([
    [0.123, '0%', '12.3%'],
    [45292.5, 'm/d/yy', '1/1/2024 12:00:00 PM'],
    [45292.75, 'mmm d', '1/1/2024 6:00:00 PM'],
    [45292 + (13 * 3600 + 45 * 60 + 30) / 86400, 'm/d/yy h:mm', '1/1/2024 1:45:30 PM'],
    [1, 'm/d/yy', '1/1/1900'],
    [0.5, 'h:mm', '12:00:00 PM'],
    [45292, 'yyyy-mm-dd', '1/1/2024'],
  ])('%s formatted %s edits as %s', (value, numFmt, text) => {
    const ctl = withValue(value, numFmt);
    ctl.startEdit();
    expect(ctl.edit?.text).toBe(text);
    expect(ctl.commitEdit()).toBe(true);
    expect(getCell(ctl.doc.ws, 1, 1)?.value).toBeCloseTo(value, 9);
  });

  it('shows a number outside the date range as a number', () => {
    const ctl = withValue(-1, 'm/d/yy');
    ctl.startEdit();
    expect(ctl.edit?.text).toBe('-1');
  });
});

describe('Find in formulas', () => {
  it('matches a date by the year the formula bar shows, and replaces it', () => {
    const ctl = withValue(45306, 'm/d/yy');
    const opts = { query: '2024', matchCase: false, wholeCell: false, byColumns: false, lookIn: 'formulas' as const };
    expect(findAll(ctl, opts, [ctl.doc.ws])).toHaveLength(1);
    expect(replaceAll(ctl, opts, '2025', [ctl.doc.ws]).replaced).toBe(1);
    expect(getCell(ctl.doc.ws, 1, 1)?.value).toBe(45306 + 366);
  });
});
