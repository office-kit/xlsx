// Cell-range value object + set operations.
//
// The struct is the same `CellRangeBoundaries` we already use across the
// coordinate parser; this module adds the worksheet-level operations
// (containment, shift, union, intersection, iteration) and a `MultiCellRange`
// lite wrapper for sqref-style attributes.

import type { Cell } from '../cell/cell.js';
import {
  boundariesToRangeString,
  type CellRangeBoundaries,
  coordinateToTuple,
  MAX_COL,
  MAX_ROW,
  rangeBoundaries,
  type RangeRef,
  tupleToCoordinate,
} from '../utils/coordinate.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';

/** Re-export under the plan's canonical name. */
export type CellRange = CellRangeBoundaries;

/** Build a CellRange from explicit 1-based bounds. */
export function makeCellRange(minRow: number, minCol: number, maxRow: number, maxCol: number): CellRange {
  if (
    !Number.isInteger(minRow) ||
    !Number.isInteger(minCol) ||
    !Number.isInteger(maxRow) ||
    !Number.isInteger(maxCol)
  ) {
    throw new OpenXmlSchemaError('CellRange bounds must be integers');
  }
  if (minRow < 1 || maxRow < 1 || minRow > MAX_ROW || maxRow > MAX_ROW) {
    throw new OpenXmlSchemaError(`CellRange row bounds must be in [1, ${MAX_ROW}]`);
  }
  if (minCol < 1 || maxCol < 1 || minCol > MAX_COL || maxCol > MAX_COL) {
    throw new OpenXmlSchemaError(`CellRange col bounds must be in [1, ${MAX_COL}]`);
  }
  return {
    minRow: Math.min(minRow, maxRow),
    minCol: Math.min(minCol, maxCol),
    maxRow: Math.max(minRow, maxRow),
    maxCol: Math.max(minCol, maxCol),
  };
}

/**
 * Resolve a {@link RangeRef} to numeric bounds: A1 expressions go through
 * {@link rangeBoundaries}, pre-computed bounds through {@link makeCellRange}.
 * Both paths validate against the grid and normalise inverted bounds, so
 * `{ minRow: 5, maxRow: 1 }` behaves like `"A5:A1"` rather than iterating zero
 * rows, and a fractional or off-grid bound throws here instead of half-way
 * through the caller's loop.
 */
export function parseRange(input: RangeRef): CellRange {
  if (typeof input === 'string') return rangeBoundaries(input);
  return makeCellRange(input.minRow, input.minCol, input.maxRow, input.maxCol);
}

/** Format a CellRange back into the canonical OOXML string. */
export function rangeToString(r: CellRange): string {
  return boundariesToRangeString(r);
}

/**
 * Compute the bounding A1-style range string for a list of cells. Walks the
 * input once to find min/max row+col. A single-cell input returns a single-cell
 * ref (`"A1"`); two or more cells (even collinear) return the `"A1:B5"` form.
 *
 * Throws when the array is empty — there's no meaningful zero-cell range, and
 * silently returning `""` would defeat downstream `parseRange` consumers.
 */
export function cellRangeFromCells(cells: ReadonlyArray<Pick<Cell, 'row' | 'col'>>): string {
  if (cells.length === 0) {
    throw new OpenXmlSchemaError('cellRangeFromCells: cells array must be non-empty');
  }
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = Number.NEGATIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = Number.NEGATIVE_INFINITY;
  for (const c of cells) {
    if (c.row < minRow) minRow = c.row;
    if (c.row > maxRow) maxRow = c.row;
    if (c.col < minCol) minCol = c.col;
    if (c.col > maxCol) maxCol = c.col;
  }
  if (minRow === maxRow && minCol === maxCol) {
    return tupleToCoordinate(minCol, minRow);
  }
  return boundariesToRangeString({ minRow, minCol, maxRow, maxCol });
}

/** Inclusive containment of a single (row, col) within a range. */
export function rangeContainsCell(r: CellRange, row: number, col: number): boolean {
  return row >= r.minRow && row <= r.maxRow && col >= r.minCol && col <= r.maxCol;
}

/**
 * A1-string convenience for {@link rangeContainsCell}. Parses `cellRef` (e.g.
 * `"B3"`) and `rangeRef` (e.g. `"A1:C5"`) and returns `true` iff the cell sits
 * inside the range (boundary-inclusive). Throws when either input is malformed.
 */
export function isCellInRange(cellRef: string, rangeRef: string): boolean {
  const { col, row } = coordinateToTuple(cellRef);
  return rangeContainsCell(parseRange(rangeRef), row, col);
}

/**
 * A1-string convenience for {@link rangeContainsRange}. Returns `true` iff the
 * `inner` range is wholly contained by `outer` (boundary-inclusive).
 * Single-cell refs are accepted on either side via parseRange. Throws on
 * malformed input.
 */
export function isRangeInRange(inner: string, outer: string): boolean {
  return rangeContainsRange(parseRange(outer), parseRange(inner));
}

/**
 * A1-string convenience for {@link rangesOverlap}. Returns `true` iff the two
 * ranges share at least one cell. Boundary-inclusive (a 1-row gap = no
 * overlap). Single-cell refs are accepted via parseRange. Throws on malformed
 * input.
 */
export function rangesOverlapStr(a: string, b: string): boolean {
  return rangesOverlap(parseRange(a), parseRange(b));
}

/**
 * A1-string convenience for {@link unionRange}. Returns the smallest A1 range
 * that contains both inputs (always non-null; ranges that don't overlap still
 * get a valid bounding box).
 */
export function unionRangeStr(a: string, b: string): string {
  return rangeToString(unionRange(parseRange(a), parseRange(b)));
}

/**
 * A1-string convenience for {@link intersectionRange}. Returns the shared
 * rectangular sub-range as A1 string, or `undefined` when the inputs are
 * disjoint.
 */
export function intersectionRangeStr(a: string, b: string): string | undefined {
  const r = intersectionRange(parseRange(a), parseRange(b));
  return r === null ? undefined : rangeToString(r);
}

/**
 * A1-string convenience for {@link shiftRange}. Translates the range by `(dr,
 * dc)` integer offsets and re-serialises. Negative offsets shift up/left.
 * Throws when the resulting bounds fall outside the OOXML grid (rows
 * 1..1048576, cols 1..16384).
 */
export function shiftRangeStr(range: string, dr: number, dc: number): string {
  return rangeToString(shiftRange(parseRange(range), dr, dc));
}

/**
 * A1-string convenience for {@link rangeArea}. Returns the inclusive cell count
 * covered by the range (rows × cols). Single-cell refs return 1.
 */
export function rangeAreaStr(range: string): number {
  return rangeArea(parseRange(range));
}

/**
 * A1-string range dimensions: how many rows and columns the range spans
 * (inclusive). Distinct from {@link rangeAreaStr} which returns the product.
 * Single-cell refs return `{ rows: 1, cols: 1 }`.
 */
export function rangeDimensionsStr(range: string): { rows: number; cols: number } {
  const r = parseRange(range);
  return { rows: r.maxRow - r.minRow + 1, cols: r.maxCol - r.minCol + 1 };
}

/**
 * Expand (or shrink) an A1 range by adding `deltaRows` to its bottom edge and
 * `deltaCols` to its right edge. The top-left corner is preserved. Negative
 * deltas shrink the range; the result must still have at least 1 row and 1
 * column (otherwise throws).
 *
 * Useful for "this is the data range — also reserve room for a totals row" or
 * "include one more column to the right" patterns.
 */
export function expandRangeStr(range: string, deltaRows: number, deltaCols: number): string {
  if (!Number.isInteger(deltaRows) || !Number.isInteger(deltaCols)) {
    throw new OpenXmlSchemaError('expandRangeStr: deltas must be integers');
  }
  const r = parseRange(range);
  return rangeToString(makeCellRange(r.minRow, r.minCol, r.maxRow + deltaRows, r.maxCol + deltaCols));
}

/** Inclusive containment of `inner` within `outer`. */
export function rangeContainsRange(outer: CellRange, inner: CellRange): boolean {
  return (
    inner.minRow >= outer.minRow &&
    inner.maxRow <= outer.maxRow &&
    inner.minCol >= outer.minCol &&
    inner.maxCol <= outer.maxCol
  );
}

/**
 * Shift a range by (dr, dc) integer offsets. Throws when the result falls
 * outside the OOXML grid (rows 1..1048576, cols 1..16384) rather than clamping
 * to it, so a shift that would silently change which cells the range covers
 * fails instead. {@link shiftRangeStr} is the same rule on A1 strings.
 */
export function shiftRange(r: CellRange, dr: number, dc: number): CellRange {
  if (!Number.isInteger(dr) || !Number.isInteger(dc)) {
    throw new OpenXmlSchemaError('shiftRange: dr / dc must be integers');
  }
  return makeCellRange(r.minRow + dr, r.minCol + dc, r.maxRow + dr, r.maxCol + dc);
}

/** Bounding-box union of two ranges. Always non-null. */
export function unionRange(a: CellRange, b: CellRange): CellRange {
  return {
    minRow: Math.min(a.minRow, b.minRow),
    minCol: Math.min(a.minCol, b.minCol),
    maxRow: Math.max(a.maxRow, b.maxRow),
    maxCol: Math.max(a.maxCol, b.maxCol),
  };
}

/** Returns the rectangular intersection of two ranges, or `null` when disjoint. */
export function intersectionRange(a: CellRange, b: CellRange): CellRange | null {
  const minRow = Math.max(a.minRow, b.minRow);
  const minCol = Math.max(a.minCol, b.minCol);
  const maxRow = Math.min(a.maxRow, b.maxRow);
  const maxCol = Math.min(a.maxCol, b.maxCol);
  if (minRow > maxRow || minCol > maxCol) return null;
  return { minRow, minCol, maxRow, maxCol };
}

/** True iff two ranges share at least one cell. */
export function rangesOverlap(a: CellRange, b: CellRange): boolean {
  return intersectionRange(a, b) !== null;
}

/** Inclusive cell count covered by a range. */
export function rangeArea(r: CellRange): number {
  return (r.maxRow - r.minRow + 1) * (r.maxCol - r.minCol + 1);
}

/** Yield every (row, col) coordinate in the range, row-major. */
export function* iterRangeCoordinates(r: CellRange): IterableIterator<{ row: number; col: number }> {
  for (let row = r.minRow; row <= r.maxRow; row++) {
    for (let col = r.minCol; col <= r.maxCol; col++) {
      yield { row, col };
    }
  }
}

// ---- MultiCellRange --------------------------------------------------------

/**
 * Excel's `sqref` attribute: a space-separated list of CellRanges. Used by data
 * validations, conditional formatting, hyperlinks etc.
 */
export interface MultiCellRange {
  ranges: CellRange[];
}

export function makeMultiCellRange(ranges: ReadonlyArray<CellRange> = []): MultiCellRange {
  return { ranges: ranges.slice() };
}

/** Parse an sqref string: `"A1:B2 D5 E10:F20"`. Whitespace-delimited. */
export function parseMultiCellRange(input: string): MultiCellRange {
  const ranges: CellRange[] = [];
  for (const piece of input.split(/\s+/)) {
    if (piece.length === 0) continue;
    ranges.push(parseRange(piece));
  }
  return { ranges };
}

/** Format a MultiCellRange back into an sqref string. */
export function multiCellRangeToString(m: MultiCellRange): string {
  return m.ranges.map(rangeToString).join(' ');
}

/** Total cell count across all ranges (no de-duplication of overlaps). */
export function multiCellRangeArea(m: MultiCellRange): number {
  let n = 0;
  for (const r of m.ranges) n += rangeArea(r);
  return n;
}

/** True iff any contained range covers (row, col). */
export function multiCellRangeContainsCell(m: MultiCellRange, row: number, col: number): boolean {
  for (const r of m.ranges) if (rangeContainsCell(r, row, col)) return true;
  return false;
}
