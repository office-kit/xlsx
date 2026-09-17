// Cell-value type inference. Maps a JS runtime value to the OOXML cell
// `t` attribute value. Mirrors openpyxl's `Cell._bind_value` /
// `_TYPES` / error-code path in openpyxl/openpyxl/cell/cell.py.

import { startsWithEquals } from './formula-text.js';

/**
 * OOXML `t` attribute values. Note that 'inlineStr' is treated
 * separately — the writer chooses between 's' (shared string) and
 * 'inlineStr' based on workbook settings, not the value itself.
 */
export type CellDataType = 'n' | 's' | 'b' | 'd' | 'f' | 'e';

/**
 * Excel error tokens this library knows by name. On a write, a string outside
 * this set is text, so the set is what {@link inferCellType} and `makeErrorValue`
 * accept as an error.
 *
 * It is not a closed description of the format: Excel has added nine tokens
 * since 2018 and can add more, so a `t="e"` cell read out of a file keeps
 * whatever token it carries whether or not that token is listed here.
 */
export const ERROR_CODES: ReadonlySet<string> = new Set([
  // Pre-2018, and all a transitional-conformance file can hold.
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  // Excel 365: dynamic arrays, Power Query, linked data types, Python.
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#CONNECT!',
  '#BUSY!',
  '#UNKNOWN!',
  '#PYTHON!',
  '#EXTERNAL!',
]);

/**
 * Infer the cell `t` attribute for a runtime value.
 *
 * - `boolean` → 'b'
 * - `number` → 'n' (incl. integer numerics; date inference is left to
 *   the caller because Excel decides on type via the cell's number
 *   format, not the raw value)
 * - `Date` → 'd'
 * - string starting with `=` → 'f' (formula)
 * - string in {@link ERROR_CODES} → 'e'
 * - any other string → 's'
 * - `null` / `undefined` → 'n' (empty)
 *
 * Throws nothing — returns 'n' as the no-information fallback.
 */
export function inferCellType(value: unknown): CellDataType {
  if (typeof value === 'boolean') return 'b';
  if (typeof value === 'number') return 'n';
  if (value instanceof Date) return 'd';
  if (typeof value === 'string') {
    if (startsWithEquals(value)) return 'f';
    if (ERROR_CODES.has(value)) return 'e';
    return 's';
  }
  return 'n';
}
