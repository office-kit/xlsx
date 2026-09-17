// Caps on the content a read will model, and the running totals that enforce
// them. The companion to the zip layer's decompression guard: that one bounds
// the bytes an archive inflates to, this one bounds the cells those bytes turn
// into, which is what decides how long a read takes and how much heap it holds.

import { formatSheetQualifiedRef, tupleToCoordinate } from '../utils/coordinate.js';
import { OpenXmlContentLimitError, OpenXmlError } from '../utils/exceptions.js';

/**
 * Caps on the content one read will model. Both are unlimited when absent, so
 * a caller that asks for nothing gets what it always got.
 */
export interface ContentLimits {
  /** Maximum cells the read will place before it refuses to continue. */
  maxCells?: number;
  /** Maximum rows the read will place before it refuses to continue. */
  maxRows?: number;
}

/** {@link ContentLimits} with every field settled; `Infinity` is "unlimited". */
interface ResolvedContentLimits {
  readonly maxCells: number;
  readonly maxRows: number;
}

// A cap has to be a positive integer to mean anything: a NaN, a negative or a
// zero turns `count > cap` into a gate that either never opens or never
// closes. Reject those where the caller hands them over rather than at the
// first cell, where the message would name a cell instead of the mistake.
const requirePositiveInteger = (field: string, value: number): void => {
  if (typeof value !== 'number' || !Number.isInteger(value) || value <= 0) {
    throw new OpenXmlError(`contentLimits.${field} must be a positive integer; got ${String(value)}`);
  }
};

/** Running totals for one read, shared by every worksheet it covers. */
export interface ContentBudget {
  readonly limits: ResolvedContentLimits;
  cells: number;
  rows: number;
}

/** Build the budget for one read. `undefined` limits give an unlimited budget. */
export function makeContentBudget(input: ContentLimits | undefined): ContentBudget {
  if (input?.maxCells !== undefined) requirePositiveInteger('maxCells', input.maxCells);
  if (input?.maxRows !== undefined) requirePositiveInteger('maxRows', input.maxRows);
  return {
    limits: {
      maxCells: input?.maxCells ?? Number.POSITIVE_INFINITY,
      maxRows: input?.maxRows ?? Number.POSITIVE_INFINITY,
    },
    cells: 0,
    rows: 0,
  };
}

/**
 * Charge one cell to the budget. Called before the cell is built, so a read
 * that is going to be refused does not first allocate what it is refusing.
 *
 * Takes the coordinate as a column and a row rather than a formatted ref: this
 * runs once per cell of the sheet, and the ref is only needed on the one call
 * that throws.
 */
export function chargeCell(budget: ContentBudget, sheet: string, col: number, row: number): void {
  budget.cells++;
  if (budget.cells > budget.limits.maxCells) {
    const at = formatSheetQualifiedRef(sheet, tupleToCoordinate(col, row));
    throw new OpenXmlContentLimitError(
      `worksheet: reading ${at} passes contentLimits.maxCells of ${budget.limits.maxCells}`,
    );
  }
}

/** Charge one row to the budget. */
export function chargeRow(budget: ContentBudget, sheet: string, row: number): void {
  budget.rows++;
  if (budget.rows > budget.limits.maxRows) {
    throw new OpenXmlContentLimitError(
      `worksheet: reading row ${row} of ${sheet} passes contentLimits.maxRows of ${budget.limits.maxRows}`,
    );
  }
}
