// Screen ↔ sheet mapping for the canvas grid, shared by the painter and the
// pointer handlers so hit-testing always agrees with what was drawn.
//
// Layout: a header strip (column letters) across the top and a gutter (row
// numbers) down the left, then up to four panes when panes are frozen. Frozen
// rows/columns never scroll; the main pane scrolls by `scrollX`/`scrollY`
// content pixels.

import type { Range } from '../core/address.ts';
import { MAX_COL, MAX_ROW } from '../core/address.ts';
import type { Axis } from '../core/axis.ts';

export interface GeometryInput {
  readonly cols: Axis;
  readonly rows: Axis;
  readonly zoom: number;
  readonly scrollX: number;
  readonly scrollY: number;
  readonly frozenRows: number;
  readonly frozenCols: number;
  readonly width: number;
  readonly height: number;
  readonly showHeaders: boolean;
  /** Deepest row / column outline level (0 = no outline gutter). */
  readonly rowOutlineLevels: number;
  readonly colOutlineLevels: number;
}

/** Width of one outline level's strip, at 100% zoom. */
export const OUTLINE_STEP = 14;

export type HitTarget =
  | { readonly kind: 'corner' }
  | { readonly kind: 'outline'; readonly x: number; readonly y: number }
  | { readonly kind: 'colHeader'; readonly col: number; readonly x: number }
  | { readonly kind: 'rowHeader'; readonly row: number; readonly y: number }
  | { readonly kind: 'cell'; readonly row: number; readonly col: number };

export class GridGeometry {
  readonly cols: Axis;
  readonly rows: Axis;
  readonly zoom: number;
  readonly scrollX: number;
  readonly scrollY: number;
  readonly frozenRows: number;
  readonly frozenCols: number;
  readonly width: number;
  readonly height: number;
  /** Total left / top gutters: outline strips plus row numbers / column letters. */
  readonly headerW: number;
  readonly headerH: number;
  /** Outline strips left of the row numbers and above the column letters. */
  readonly outlineW: number;
  readonly outlineH: number;
  /** Screen extent of the frozen columns / rows (0 when none). */
  readonly frozenW: number;
  readonly frozenH: number;

  constructor(input: GeometryInput) {
    this.cols = input.cols;
    this.rows = input.rows;
    this.zoom = input.zoom;
    this.scrollX = input.scrollX;
    this.scrollY = input.scrollY;
    this.frozenRows = input.frozenRows;
    this.frozenCols = input.frozenCols;
    this.width = input.width;
    this.height = input.height;
    // Page Layout view's axes start with a margin, so "nothing frozen" must be 0, not offsetOf(1).
    this.frozenW = this.frozenCols > 0 ? this.cols.offsetOf(this.frozenCols + 1) * this.zoom : 0;
    this.frozenH = this.frozenRows > 0 ? this.rows.offsetOf(this.frozenRows + 1) * this.zoom : 0;
    const scale = Math.max(0.6, Math.min(this.zoom, 2));
    // Level n needs n+1 strips: one per level plus the innermost detail strip.
    this.outlineW = input.rowOutlineLevels > 0 ? Math.round((input.rowOutlineLevels + 1) * OUTLINE_STEP * scale) : 0;
    this.outlineH = input.colOutlineLevels > 0 ? Math.round((input.colOutlineLevels + 1) * OUTLINE_STEP * scale) : 0;
    if (input.showHeaders) {
      // The gutter widens with the digit count of the last visible row number.
      const lastRow = this.rows.indexAt((this.scrollY + this.height) / this.zoom + this.frozenH / this.zoom);
      const digits = String(lastRow).length;
      this.headerH = this.outlineH + Math.round(20 * scale);
      this.headerW = this.outlineW + Math.round((Math.max(3, digits) * 8 + 10) * scale);
    } else {
      this.headerW = this.outlineW;
      this.headerH = this.outlineH;
    }
  }

  /** Left screen edge of column `c`. */
  colX(c: number): number {
    const content = this.cols.offsetOf(c) * this.zoom;
    return this.headerW + (c <= this.frozenCols ? content : content - this.scrollX);
  }

  rowY(r: number): number {
    const content = this.rows.offsetOf(r) * this.zoom;
    return this.headerH + (r <= this.frozenRows ? content : content - this.scrollY);
  }

  colW(c: number): number {
    return this.cols.sizeOf(c) * this.zoom;
  }

  rowH(r: number): number {
    return this.rows.sizeOf(r) * this.zoom;
  }

  /** Screen rectangle of a range (clipped to nothing; callers clip to panes). */
  rectOf(range: Range): { x: number; y: number; w: number; h: number } {
    const x = this.colX(range.c1);
    const y = this.rowY(range.r1);
    const x2 = this.colX(Math.min(range.c2, MAX_COL) + 1);
    const y2 = this.rowY(Math.min(range.r2, MAX_ROW) + 1);
    return { x, y, w: x2 - x, h: y2 - y };
  }

  /** First and last column visible in the scrolling pane. */
  mainCols(): [number, number] {
    const first = this.cols.indexAt(this.frozenW / this.zoom + this.scrollX / this.zoom);
    const last = this.cols.indexAt((this.scrollX + this.width - this.headerW) / this.zoom);
    return [Math.max(first, this.frozenCols + 1), last];
  }

  mainRows(): [number, number] {
    const first = this.rows.indexAt(this.frozenH / this.zoom + this.scrollY / this.zoom);
    const last = this.rows.indexAt((this.scrollY + this.height - this.headerH) / this.zoom);
    return [Math.max(first, this.frozenRows + 1), last];
  }

  colAt(x: number): number {
    const local = x - this.headerW;
    if (local < this.frozenW) return this.cols.indexAt(local / this.zoom);
    return this.cols.indexAt((local + this.scrollX) / this.zoom);
  }

  rowAt(y: number): number {
    const local = y - this.headerH;
    if (local < this.frozenH) return this.rows.indexAt(local / this.zoom);
    return this.rows.indexAt((local + this.scrollY) / this.zoom);
  }

  hit(x: number, y: number): HitTarget {
    if (x < this.outlineW || y < this.outlineH) return { kind: 'outline', x, y };
    if (x < this.headerW && y < this.headerH) return { kind: 'corner' };
    if (y < this.headerH) return { kind: 'colHeader', col: this.colAt(x), x };
    if (x < this.headerW) return { kind: 'rowHeader', row: this.rowAt(y), y };
    return { kind: 'cell', row: this.rowAt(y), col: this.colAt(x) };
  }

  /**
   * Column whose right edge is within a few pixels of `x` in the header — the
   * resize grip. Prefers the column to the left so a zero-width (hidden)
   * column is not grabbed by accident.
   */
  colResizeEdge(x: number): number | undefined {
    const c = this.colAt(x);
    const right = this.colX(c) + this.colW(c);
    if (Math.abs(x - right) <= 3) return c;
    const left = this.colX(c);
    if (Math.abs(x - left) <= 3 && c > 1) return this.cols.nextVisible(c, -1);
    return undefined;
  }

  rowResizeEdge(y: number): number | undefined {
    const r = this.rowAt(y);
    const bottom = this.rowY(r) + this.rowH(r);
    if (Math.abs(y - bottom) <= 3) return r;
    const top = this.rowY(r);
    if (Math.abs(y - top) <= 2 && r > 1) return this.rows.nextVisible(r, -1);
    return undefined;
  }

  /** Scroll offsets that bring `row`/`col` fully into the scrolling pane (unchanged when already visible). */
  scrollToReveal(row: number, col: number): { x: number; y: number } {
    let x = this.scrollX;
    let y = this.scrollY;
    if (col > this.frozenCols) {
      const left = this.cols.offsetOf(col) * this.zoom - this.frozenW;
      const right = left + this.colW(col);
      const viewW = this.width - this.headerW - this.frozenW;
      if (left < x) x = left;
      else if (right > x + viewW) x = Math.min(left, right - viewW);
    }
    if (row > this.frozenRows) {
      const top = this.rows.offsetOf(row) * this.zoom - this.frozenH;
      const bottom = top + this.rowH(row);
      const viewH = this.height - this.headerH - this.frozenH;
      if (top < y) y = top;
      else if (bottom > y + viewH) y = Math.min(top, bottom - viewH);
    }
    return { x: Math.max(0, x), y: Math.max(0, y) };
  }
}
