// Error token of a `t="e"` cell, shared by the two worksheet readers so that
// one file is answered the same way whether the caller loads a workbook or
// streams it.

import { cellLabel, quoteCellText } from './cell-text.js';
import { OpenXmlSchemaError } from './exceptions.js';

/** Every error token Excel has ever written carries this sigil. */
const isErrorToken = (raw: string): raw is `#${string}` => raw.startsWith('#');

/**
 * Token of an error cell. `t="e"` is the file declaring that the cell holds an
 * error, so the token is kept verbatim even when `ERROR_CODES` does not list
 * it: Excel keeps adding tokens, and dropping one loses the cell's value to
 * buy nothing. A `t="e"` carrying something that is not a token at all is a
 * producer bug, and is reported rather than stored.
 *
 * The return type spells the sigil out instead of naming `ExcelErrorCode`
 * because nothing under `src/utils/` depends on the cell model.
 */
export function parseCellErrorCode(
  raw: string | undefined,
  sheet: string,
  col: number,
  row: number,
): `#${string}` {
  if (raw !== undefined && isErrorToken(raw)) return raw;
  throw new OpenXmlSchemaError(
    `worksheet: <v>${quoteCellText(raw ?? '')}</v> at ${cellLabel(sheet, col, row)} is not an Excel error token`,
  );
}
