// Point and window queries over a sheet's merged ranges.
//
// Navigation asks "is this cell inside a merge?" on every keystroke and the
// painter asks "which merges touch the viewport?" every frame. A sheet can
// carry thousands of merges, so they are bucketed by blocks of rows; a merge
// spanning an enormous number of rows goes to a short list checked always,
// which keeps the index small without a per-cell map.

import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { fromBoundaries, inRange, rangesIntersect, type Range } from './address.ts';

const BUCKET_ROWS = 64;
const MAX_BUCKETS_PER_MERGE = 256;

export class MergeIndex {
  readonly all: readonly Range[];
  readonly #buckets = new Map<number, Range[]>();
  readonly #wide: Range[] = [];

  constructor(ws: Worksheet) {
    this.all = ws.mergedCells.map(fromBoundaries);
    for (const m of this.all) {
      const b1 = Math.floor(m.r1 / BUCKET_ROWS);
      const b2 = Math.floor(m.r2 / BUCKET_ROWS);
      if (b2 - b1 >= MAX_BUCKETS_PER_MERGE) {
        this.#wide.push(m);
        continue;
      }
      for (let b = b1; b <= b2; b++) {
        let list = this.#buckets.get(b);
        if (!list) {
          list = [];
          this.#buckets.set(b, list);
        }
        list.push(m);
      }
    }
  }

  at(row: number, col: number): Range | undefined {
    for (const m of this.#buckets.get(Math.floor(row / BUCKET_ROWS)) ?? []) if (inRange(m, row, col)) return m;
    for (const m of this.#wide) if (inRange(m, row, col)) return m;
    return undefined;
  }

  /** Merges overlapping `window`, each reported once. */
  intersecting(window: Range): Range[] {
    const out = new Set<Range>();
    const b1 = Math.floor(window.r1 / BUCKET_ROWS);
    const b2 = Math.floor(window.r2 / BUCKET_ROWS);
    if (b2 - b1 > this.#buckets.size) {
      for (const m of this.all) if (rangesIntersect(m, window)) out.add(m);
      return [...out];
    }
    for (let b = b1; b <= b2; b++) for (const m of this.#buckets.get(b) ?? []) if (rangesIntersect(m, window)) out.add(m);
    for (const m of this.#wide) if (rangesIntersect(m, window)) out.add(m);
    return [...out];
  }

  /** Grow `range` until no merge straddles its edge (Excel never selects part of a merge). */
  expand(range: Range): Range {
    let r = range;
    for (;;) {
      let grown = r;
      for (const m of this.intersecting(r)) {
        grown = {
          r1: Math.min(grown.r1, m.r1),
          c1: Math.min(grown.c1, m.c1),
          r2: Math.max(grown.r2, m.r2),
          c2: Math.max(grown.c2, m.c2),
        };
      }
      if (grown.r1 === r.r1 && grown.c1 === r.c1 && grown.r2 === r.r2 && grown.c2 === r.c2) return r;
      r = grown;
    }
  }
}

/**
 * Whether `range` can be sorted (or filled) line by line: Excel requires every
 * merged cell in it to be the same size, an unmerged cell counting as 1×1, so
 * each line must carry the same merges, each within that one line.
 */
export function mergesAllowLineMoves(ws: Worksheet, range: Range, byRows: boolean): boolean {
  const lines = new Map<number, string[]>();
  for (const m of ws.mergedCells) {
    const r = fromBoundaries(m);
    if (!rangesIntersect(r, range)) continue;
    const inside = r.r1 >= range.r1 && r.r2 <= range.r2 && r.c1 >= range.c1 && r.c2 <= range.c2;
    const oneLine = byRows ? r.r1 === r.r2 : r.c1 === r.c2;
    if (!inside || !oneLine) return false;
    const line = byRows ? r.r1 : r.c1;
    const spans = lines.get(line) ?? [];
    spans.push(byRows ? `${r.c1}:${r.c2}` : `${r.r1}:${r.r2}`);
    lines.set(line, spans);
  }
  if (lines.size === 0) return true;
  const count = byRows ? range.r2 - range.r1 + 1 : range.c2 - range.c1 + 1;
  if (lines.size !== count) return false;
  const signatures = new Set([...lines.values()].map((spans) => spans.sort().join(',')));
  return signatures.size === 1;
}
