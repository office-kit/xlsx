// Reverse-dependency index: "which formulas read this cell or rectangle?"
//
// Range references are never expanded into per-cell edges — `SUM(A:A)` is one
// entry, however many cells it covers. Each sheet keeps
//   - a Map from cell key to the formulas reading exactly that cell, and
//   - a `RangeIndex` of distinct rectangles, each with the formulas reading it.
// The RangeIndex buckets rectangles by 64-column blocks (rectangles wider than
// a few blocks share one "wide" bucket). Inside a bucket, rectangles sit in
// chunks of 32 sorted by first row, each chunk remembering its largest last
// row, so a query skips every chunk that ends above it or starts below it.
// Insertions land in a small unsorted tail and deletions are tombstoned; the
// bucket re-sorts itself once either grows large, which keeps a single-cell
// edit from re-sorting a 100k-entry column.

export interface Rect {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

interface Entry<T> extends Rect {
  readonly key: string;
  readonly readers: Set<T>;
  dead: boolean;
}

const BLOCK_COLS = 64;
const MAX_BLOCKS_PER_ENTRY = 4;
const CHUNK = 32;
const WIDE = -1;

class Bucket<T> {
  private sorted: Entry<T>[] = [];
  private chunkMax: number[] = [];
  private tail: Entry<T>[] = [];
  private deadCount = 0;

  add(e: Entry<T>): void {
    this.tail.push(e);
    if (this.tail.length > Math.max(64, this.sorted.length >> 3)) this.rebuild();
  }

  noteDead(): void {
    this.deadCount++;
    if (this.deadCount > 64 && this.deadCount > this.sorted.length >> 1) this.rebuild();
  }

  query(q: Rect, visit: (e: Entry<T>) => void): void {
    for (const e of this.tail) if (!e.dead && overlaps(e, q)) visit(e);
    const s = this.sorted;
    // Entries with r1 <= q.r2 form a prefix of the sorted array.
    let lo = 0;
    let hi = s.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if ((s[mid]?.r1 ?? 0) <= q.r2) lo = mid + 1;
      else hi = mid;
    }
    for (let chunk = 0; chunk * CHUNK < lo; chunk++) {
      if ((this.chunkMax[chunk] ?? 0) < q.r1) continue;
      const end = Math.min(lo, (chunk + 1) * CHUNK);
      for (let i = chunk * CHUNK; i < end; i++) {
        const e = s[i];
        if (e !== undefined && !e.dead && overlaps(e, q)) visit(e);
      }
    }
  }

  private rebuild(): void {
    this.sorted = [...this.sorted, ...this.tail].filter((e) => !e.dead).sort((a, b) => a.r1 - b.r1);
    this.tail = [];
    this.deadCount = 0;
    this.chunkMax = [];
    for (let i = 0; i < this.sorted.length; i += CHUNK) {
      let max = 0;
      for (let j = i; j < Math.min(i + CHUNK, this.sorted.length); j++) max = Math.max(max, this.sorted[j]?.r2 ?? 0);
      this.chunkMax.push(max);
    }
  }
}

const overlaps = (a: Rect, b: Rect): boolean => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;

export class RangeIndex<T> {
  private readonly entries = new Map<string, Entry<T>>();
  private readonly buckets = new Map<number, Bucket<T>>();

  get size(): number {
    return this.entries.size;
  }

  add(rect: Rect, reader: T): void {
    const key = `${rect.r1},${rect.c1},${rect.r2},${rect.c2}`;
    let e = this.entries.get(key);
    if (e === undefined) {
      e = { ...rect, key, readers: new Set(), dead: false };
      this.entries.set(key, e);
      for (const b of this.bucketIds(rect)) this.bucket(b).add(e);
    }
    e.readers.add(reader);
  }

  remove(rect: Rect, reader: T): void {
    const key = `${rect.r1},${rect.c1},${rect.r2},${rect.c2}`;
    const e = this.entries.get(key);
    if (e === undefined) return;
    e.readers.delete(reader);
    if (e.readers.size > 0) return;
    e.dead = true;
    this.entries.delete(key);
    for (const b of this.bucketIds(rect)) this.buckets.get(b)?.noteDead();
  }

  /** Every reader of a rectangle overlapping `q`. */
  query(q: Rect, out: Set<T>): void {
    const visit = (e: Entry<T>): void => {
      for (const r of e.readers) out.add(r);
    };
    this.buckets.get(WIDE)?.query(q, visit);
    const first = Math.floor((q.c1 - 1) / BLOCK_COLS);
    const last = Math.floor((q.c2 - 1) / BLOCK_COLS);
    for (let b = first; b <= last; b++) this.buckets.get(b)?.query(q, visit);
  }

  private bucketIds(rect: Rect): number[] {
    const first = Math.floor((rect.c1 - 1) / BLOCK_COLS);
    const last = Math.floor((rect.c2 - 1) / BLOCK_COLS);
    if (last - first + 1 > MAX_BLOCKS_PER_ENTRY) return [WIDE];
    const ids: number[] = [];
    for (let b = first; b <= last; b++) ids.push(b);
    return ids;
  }

  private bucket(id: number): Bucket<T> {
    let b = this.buckets.get(id);
    if (b === undefined) {
      b = new Bucket();
      this.buckets.set(id, b);
    }
    return b;
  }
}
