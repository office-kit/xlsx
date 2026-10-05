// Excel's selection model: one or more ranges (Ctrl/Cmd+click adds ranges),
// an active cell that always lies inside the active range, and an anchor the
// Shift-extension grows from.

import type { CellPos, Range } from './address.ts';
import { inRange, MAX_COL, MAX_ROW, rangeOf } from './address.ts';

export interface Selection {
  readonly ranges: readonly Range[];
  /** Index into `ranges` of the range that Shift-extension and Enter-cycling act on. */
  readonly activeRange: number;
  readonly active: CellPos;
  /** Corner the active range grows from; usually the active cell. */
  readonly anchor: CellPos;
}

export function singleCell(pos: CellPos): Selection {
  return { ranges: [rangeOf(pos)], activeRange: 0, active: pos, anchor: pos };
}

export function selectRange(range: Range, active: CellPos = { row: range.r1, col: range.c1 }): Selection {
  return { ranges: [range], activeRange: 0, active, anchor: active };
}

export function currentRange(sel: Selection): Range {
  return sel.ranges[sel.activeRange] ?? rangeOf(sel.active);
}

/** Replace the active range with anchor→`to`, keeping the active cell (Shift+click / Shift+arrow). */
export function extendTo(sel: Selection, to: CellPos): Selection {
  const cur = currentRange(sel);
  const next = rangeOf(sel.anchor, to);
  // Whole rows (Shift+Space, a row header) stay whole rows when extended, and
  // whole columns stay whole columns, even though the anchor sits inside them.
  const fullWidth = cur.c1 === 1 && cur.c2 === MAX_COL;
  const fullHeight = cur.r1 === 1 && cur.r2 === MAX_ROW;
  const ranges = sel.ranges.slice();
  ranges[sel.activeRange] = {
    r1: fullHeight ? 1 : next.r1,
    r2: fullHeight ? MAX_ROW : next.r2,
    c1: fullWidth ? 1 : next.c1,
    c2: fullWidth ? MAX_COL : next.c2,
  };
  return { ...sel, ranges };
}

/** Add a new range (Cmd/Ctrl+click), making it active. */
export function addRange(sel: Selection, pos: CellPos): Selection {
  const ranges = [...sel.ranges, rangeOf(pos)];
  return { ranges, activeRange: ranges.length - 1, active: pos, anchor: pos };
}

export function isMultiCell(sel: Selection): boolean {
  const r = currentRange(sel);
  return sel.ranges.length > 1 || r.r1 !== r.r2 || r.c1 !== r.c2;
}

export function selectionContains(sel: Selection, row: number, col: number): boolean {
  return sel.ranges.some((r) => inRange(r, row, col));
}

export function wholeColumns(c1: number, c2: number): Range {
  return { r1: 1, r2: MAX_ROW, c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
}

export function wholeRows(r1: number, r2: number): Range {
  return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: 1, c2: MAX_COL };
}

/**
 * Enter/Tab inside a multi-cell selection move the active cell through the
 * selection without collapsing it, wrapping row-wise (Tab) or column-wise
 * (Enter) and then on to the next range — Excel's data-entry cycle.
 */
export function cycleActive(sel: Selection, direction: 'down' | 'up' | 'right' | 'left'): Selection {
  const range = currentRange(sel);
  let { row, col } = sel.active;
  let rangeIndex = sel.activeRange;
  const forward = direction === 'down' || direction === 'right';
  const columnWise = direction === 'down' || direction === 'up';
  const step = forward ? 1 : -1;
  const advance = (): boolean => {
    if (columnWise) {
      row += step;
      if (row > range.r2 || row < range.r1) {
        row = forward ? range.r1 : range.r2;
        col += step;
        if (col > range.c2 || col < range.c1) return false;
      }
    } else {
      col += step;
      if (col > range.c2 || col < range.c1) {
        col = forward ? range.c1 : range.c2;
        row += step;
        if (row > range.r2 || row < range.r1) return false;
      }
    }
    return true;
  };
  if (!advance()) {
    rangeIndex = (rangeIndex + (forward ? 1 : sel.ranges.length - 1)) % sel.ranges.length;
    const next = sel.ranges[rangeIndex] ?? range;
    row = forward ? next.r1 : next.r2;
    col = forward ? next.c1 : next.c2;
  }
  const pos = { row, col };
  return { ranges: sel.ranges, activeRange: rangeIndex, active: pos, anchor: sel.anchor };
}
