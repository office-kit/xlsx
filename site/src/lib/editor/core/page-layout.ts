// Page Layout view: the sheet laid out on sheets of paper. Each axis is
// wrapped so cells keep their own sizes but every page starts on a fresh
// piece of paper (margin, printable area, margin, then a grey gutter), which
// lets the ordinary geometry, painter and hit-testing work unchanged.
//
// Pages are found with one binary search per page rather than a walk over
// every row, so a 1,048,576-row axis costs a few thousand searches.

import type { Axis, AxisIndex } from './axis.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';
import { manualBreaks, paperMetrics, printArea, printScale, type PaperMetrics } from './pages.ts';

/** Grey space between and around the sheets of paper, in sheet pixels at 100%. */
export const PAGE_GUTTER = 18;

export class PagedAxis implements Axis {
  readonly count: number;
  /** First index of each page, ascending; starts[0] = 1. */
  readonly starts: Int32Array;
  /** Paper extent along this axis: lead margin + printable size + trail margin. */
  readonly paper: number;
  readonly lead: number;
  readonly printable: number;
  readonly #base: AxisIndex;

  constructor(base: AxisIndex, printable: number, lead: number, trail: number, manual: ReadonlySet<number>) {
    this.#base = base;
    this.count = base.count;
    this.lead = lead;
    this.printable = printable;
    this.paper = lead + printable + trail;
    this.starts = pageStarts(base, printable, manual);
  }

  get pitch(): number {
    return this.paper + PAGE_GUTTER;
  }

  /** Page (0-based) containing `index`. */
  pageOf(index: number): number {
    let lo = 0;
    let hi = this.starts.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >>> 1;
      if ((this.starts[mid] ?? 1) <= index) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  /** Leading edge of page `p`'s paper. */
  paperOffset(p: number): number {
    return PAGE_GUTTER + p * this.pitch;
  }

  /** Index range of page `p` (end exclusive). */
  pageSpan(p: number): { start: number; end: number } {
    return { start: this.starts[p] ?? this.count + 1, end: this.starts[p + 1] ?? this.count + 1 };
  }

  sizeOf(index: number): number {
    return this.#base.sizeOf(index);
  }

  isHidden(index: number): boolean {
    return this.#base.isHidden(index);
  }

  offsetOf(index: number): number {
    if (index > this.count) return this.total;
    const p = this.pageOf(index);
    const start = this.starts[p] ?? 1;
    return this.paperOffset(p) + this.lead + this.#base.offsetOf(index) - this.#base.offsetOf(start);
  }

  get total(): number {
    return this.paperOffset(this.starts.length);
  }

  indexAt(px: number): number {
    const p = Math.max(0, Math.min(this.starts.length - 1, Math.floor((px - PAGE_GUTTER) / this.pitch)));
    const { start, end } = this.pageSpan(p);
    const local = px - this.paperOffset(p) - this.lead;
    if (local <= 0) return this.#base.indexAt(this.#base.offsetOf(start));
    // Past the last cell of the page (right margin or gutter) maps to that last cell.
    return Math.min(end - 1, this.#base.indexAt(this.#base.offsetOf(start) + local));
  }

  nextVisible(index: number, step: 1 | -1): number {
    return this.#base.nextVisible(index, step);
  }
}

function pageStarts(base: AxisIndex, printable: number, manual: ReadonlySet<number>): Int32Array {
  const sortedManual = [...manual].sort((a, b) => a - b);
  let m = 0;
  const starts: number[] = [1];
  let start = 1;
  while (start <= base.count) {
    // The first index that no longer fits on this page; at least one cell per page.
    let next = base.indexAt(base.offsetOf(start) + printable);
    if (base.offsetOf(next + 1) - base.offsetOf(start) <= printable) next++;
    next = Math.max(start + 1, next);
    while (m < sortedManual.length && (sortedManual[m] ?? 0) < start) m++;
    const manualNext = sortedManual[m];
    if (manualNext !== undefined && manualNext + 1 < next) next = manualNext + 1;
    if (next > base.count) break;
    starts.push(next);
    start = next;
  }
  return Int32Array.from(starts);
}

export interface PageLayout {
  readonly cols: PagedAxis;
  readonly rows: PagedAxis;
  readonly metrics: PaperMetrics;
}

export function buildPageLayout(doc: SpreadsheetEditor): PageLayout {
  const ws = doc.ws;
  const area = printArea(doc);
  const scale = area ? printScale(doc, area) : ((ws.pageSetup?.scale ?? 100) / 100);
  const m = paperMetrics(ws, scale);
  const breaks = manualBreaks(ws);
  return {
    cols: new PagedAxis(doc.cols, m.paperW - m.left - m.right, m.left, m.right, breaks.cols),
    rows: new PagedAxis(doc.rows, m.paperH - m.top - m.bottom, m.top, m.bottom, breaks.rows),
    metrics: m,
  };
}
