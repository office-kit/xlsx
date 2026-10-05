import { describe, expect, it } from 'vitest';
import { AxisIndex } from './axis.ts';
import { PAGE_GUTTER, PagedAxis } from './page-layout.ts';

describe('PagedAxis', () => {
  // 10 px cells, 35 px printable: 3 cells per page, 5 px margins.
  const base = new AxisIndex(10, 10, new Map());
  const axis = new PagedAxis(base, 35, 5, 5, new Set());

  it('starts a new sheet of paper when a cell no longer fits', () => {
    expect([...axis.starts]).toEqual([1, 4, 7, 10]);
    expect(axis.offsetOf(1)).toBe(PAGE_GUTTER + 5);
    expect(axis.offsetOf(3)).toBe(PAGE_GUTTER + 25);
    expect(axis.offsetOf(4)).toBe(PAGE_GUTTER + axis.pitch + 5);
  });

  it('maps pixels back to cells, margins to the nearest cell of the page', () => {
    expect(axis.indexAt(axis.offsetOf(5) + 1)).toBe(5);
    expect(axis.indexAt(PAGE_GUTTER + axis.pitch - 2)).toBe(3);
    expect(axis.indexAt(0)).toBe(1);
  });

  it('honours manual breaks and zero-size runs', () => {
    expect([...new PagedAxis(base, 35, 5, 5, new Set([1])).starts]).toEqual([1, 2, 5, 8]);
    const hidden = new AxisIndex(10, 10, new Map([[2, 0], [3, 0], [4, 0]]));
    expect([...new PagedAxis(hidden, 35, 5, 5, new Set()).starts]).toEqual([1, 7, 10]);
  });
});
