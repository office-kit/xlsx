// The evaluator's view of the workbook. `CalcEngine` implements it over the
// live model; keeping it an interface lets the evaluator and the function
// library stay ignorant of graphs, caches and spill bookkeeping.

import type { AstNode, StructuredSpec } from './ast.ts';
import type { Area, CalcArray, CalcError, CalcScalar } from './types.ts';

export interface EvalHost {
  readonly date1904: boolean;
  /** Canonical sheet name for a case-insensitive lookup, or undefined when there is no such worksheet. */
  sheetName(name: string): string | undefined;
  /** Worksheet names from `first` to `last` in tab order, for 3-D references. */
  sheetsBetween(first: string, last: string): string[] | undefined;
  /** Value of one cell. Formula cells yield their current result. */
  cellValue(sheet: string, row: number, col: number): CalcScalar;
  /**
   * Values of `area` as a row-major array. Areas far larger than the sheet's
   * populated extent are trimmed to it, so `A:A` costs the used rows, not a
   * million.
   */
  areaValues(area: Area): CalcArray;
  /** Parsed body of the defined name visible from `sheet` (sheet scope first), if any. */
  definedName(name: string, sheet: string): AstNode | undefined;
  /** Resolve a structured reference; `table` undefined means the table containing (row, col). */
  structuredArea(table: string | undefined, spec: StructuredSpec, sheet: string, row: number, col: number): Area | CalcError;
  /** The area a dynamic-array formula at (row, col) currently spills into. */
  spillArea(sheet: string, row: number, col: number): Area | undefined;
  /** Formula text (without `=`) of a formula cell, undefined otherwise. */
  formulaText(sheet: string, row: number, col: number): string | undefined;
  /** Whether the cell holds a SUBTOTAL / AGGREGATE formula, which those functions skip to avoid double counting. */
  isSubtotalCell(sheet: string, row: number, col: number): boolean;
  isRowHidden(sheet: string, row: number): boolean;
  /** Number format code of a cell, for CELL("format"). */
  numberFormat(sheet: string, row: number, col: number): string;
  /** Current moment as a serial (for NOW / TODAY). */
  now(): number;
  random(): number;
}
