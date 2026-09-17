// Shared error-message construction for the cell readers, so that one file is
// rejected in the same words whether the caller loads a workbook or streams it.

import { formatSheetQualifiedRef, tupleToCoordinate } from './coordinate.js';
import { OpenXmlSchemaError } from './exceptions.js';

/** How much of an offending value an error echoes. Cell text is untrusted and unbounded. */
const MAX_QUOTED_LENGTH = 40;

/** Quote cell text into an error message, bounded, with the full length named. */
export function quoteCellText(raw: string): string {
  if (raw.length <= MAX_QUOTED_LENGTH) return raw;
  return `${raw.slice(0, MAX_QUOTED_LENGTH)}... (${raw.length} chars)`;
}

/** `Sheet1!B2` label for an error message. */
export function cellLabel(sheet: string, col: number, row: number): string {
  return formatSheetQualifiedRef(sheet, tupleToCoordinate(col, row));
}

/** A `t` outside ST_CellType (ECMA-376 §18.18.11). Neither reader can guess what the `<v>` holds. */
export function unknownCellType(t: string, sheet: string, col: number, row: number): OpenXmlSchemaError {
  return new OpenXmlSchemaError(
    `worksheet: unknown cell type t="${quoteCellText(t)}" at ${cellLabel(sheet, col, row)}`,
  );
}
