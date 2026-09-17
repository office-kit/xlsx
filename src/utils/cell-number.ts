// Numeric `<v>` text, shared by the two worksheet readers so that one corrupt
// value is answered the same way whether the caller loads a workbook or
// streams it.

import { cellLabel, quoteCellText } from './cell-text.js';
import { OpenXmlSchemaError } from './exceptions.js';

/**
 * Value of a numeric cell's `<v>`: a finite number, or `null` when the element
 * is absent or blank, which is an empty cell.
 *
 * `Number.parseFloat` answers NaN for text and ±Infinity for an exponent past
 * the double range. The writer refuses to serialise either, so a file carrying
 * one is rejected here, while the sheet and cell that hold it are still known,
 * rather than at the far end of the pipeline on save.
 */
export function parseCellNumber(raw: string | undefined, sheet: string, col: number, row: number): number | null {
  if (raw === undefined) return null;
  const n = Number.parseFloat(raw);
  if (Number.isFinite(n)) return n;
  // parseFloat answers NaN for blank text too, and a blank `<v>` is an empty
  // cell rather than a corrupt one. Both readers keep element text verbatim, so
  // `<v>` in a pretty-printed part arrives with its indentation attached.
  if (raw.trim() === '') return null;
  throw new OpenXmlSchemaError(
    `worksheet: <v>${quoteCellText(raw)}</v> at ${cellLabel(sheet, col, row)} is not a finite number`,
  );
}
