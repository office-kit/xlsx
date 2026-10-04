// Pixel layout of one grid axis (rows or columns).
//
// A sheet has 1,048,576 rows and 16,384 columns, but only a handful of them
// carry a custom size. Storing an offset per index would cost 8 MB per sheet
// and an O(n) rebuild per resize, so the axis keeps the default size plus a
// sorted list of overrides with prefix sums of their deltas. `offsetOf` and
// `indexAt` are then binary searches whose cost depends on the number of
// overrides, not on the sheet size.

/** What the grid needs from an axis; Page Layout view wraps an AxisIndex with page gaps. */
export interface Axis {
  readonly count: number;
  sizeOf(index: number): number;
  isHidden(index: number): boolean;
  offsetOf(index: number): number;
  readonly total: number;
  indexAt(px: number): number;
  nextVisible(index: number, step: 1 | -1): number;
}

export class AxisIndex implements Axis {
  readonly count: number;
  readonly defaultSize: number;
  /** Ascending indices whose size differs from `defaultSize` (hidden = 0). */
  readonly #keys: Int32Array;
  readonly #sizes: Float64Array;
  /** `#deltaPrefix[k]` = Σ (size - default) over `#keys[0..k)`. */
  readonly #deltaPrefix: Float64Array;

  constructor(count: number, defaultSize: number, overrides: ReadonlyMap<number, number>) {
    this.count = count;
    this.defaultSize = defaultSize;
    const keys = [...overrides.keys()].filter((k) => k >= 1 && k <= count && overrides.get(k) !== defaultSize);
    keys.sort((a, b) => a - b);
    this.#keys = Int32Array.from(keys);
    this.#sizes = Float64Array.from(keys, (k) => overrides.get(k) ?? defaultSize);
    this.#deltaPrefix = new Float64Array(keys.length + 1);
    for (let k = 0; k < keys.length; k++) {
      this.#deltaPrefix[k + 1] = (this.#deltaPrefix[k] ?? 0) + (this.#sizes[k] ?? 0) - defaultSize;
    }
  }

  /** Number of override keys strictly below `index`. */
  #rank(index: number): number {
    let lo = 0;
    let hi = this.#keys.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if ((this.#keys[mid] ?? 0) < index) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  sizeOf(index: number): number {
    const k = this.#rank(index);
    return this.#keys[k] === index ? (this.#sizes[k] ?? this.defaultSize) : this.defaultSize;
  }

  isHidden(index: number): boolean {
    return this.sizeOf(index) === 0;
  }

  /** Pixel position of the leading edge of `index` (1-based; offsetOf(1) = 0). */
  offsetOf(index: number): number {
    return (index - 1) * this.defaultSize + (this.#deltaPrefix[this.#rank(index)] ?? 0);
  }

  /** Total extent of indices 1..count. */
  get total(): number {
    return this.offsetOf(this.count + 1);
  }

  /** Index whose span contains `px` (clamped to [1, count]); hidden indices are never returned. */
  indexAt(px: number): number {
    if (px <= 0) return this.#firstVisibleFrom(1);
    let lo = 1;
    let hi = this.count;
    while (lo < hi) {
      const mid = (lo + hi + 1) >>> 1;
      if (this.offsetOf(mid) <= px) lo = mid;
      else hi = mid - 1;
    }
    return this.#firstVisibleFrom(lo);
  }

  // Hidden indices are consecutive override keys, so skipping a hidden run is
  // a walk over the key array rather than one binary search per index.
  #firstVisibleFrom(index: number): number {
    let i = index;
    let k = this.#rank(i);
    while (i < this.count && this.#keys[k] === i && this.#sizes[k] === 0) {
      i++;
      k++;
    }
    return i;
  }

  /** Next visible index after `index` in direction `step` (±1), or `index` itself at the edge. */
  nextVisible(index: number, step: 1 | -1): number {
    let i = index + step;
    if (step === 1) {
      if (i > this.count) return index;
      i = this.#firstVisibleFrom(i);
      return this.isHidden(i) ? index : i;
    }
    let k = this.#rank(i);
    while (i >= 1 && this.#keys[k] === i && this.#sizes[k] === 0) {
      i--;
      k--;
    }
    return i < 1 ? index : i;
  }
}
