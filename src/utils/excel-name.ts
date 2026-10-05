// Names Excel accepts for defined names and tables. Excel's Name Manager and
// Table Name box refuse anything else, and a saved file carrying one is
// "repaired" on open.

import { OpenXmlSchemaError } from './exceptions.js';
import { isValidCellRef } from './coordinate.js';

const MAX_NAME_LENGTH = 255;
// A letter, `_` or `\` first; then letters, digits, `_`, `.`, `\` or `?`.
const NAME_RE = /^[\p{L}_\\][\p{L}\p{N}_.\\?]*$/u;
// R1C1 references (`R`, `C`, `R2C3`, `r5`) read as cells in either mode.
const R1C1_RE = /^(?:[Rr]\d*[Cc]?\d*|[Cc]\d*)$/;

/** Throw unless `name` is a name Excel accepts; `what` prefixes the message. */
export function assertExcelName(name: string, what: string): void {
  let problem: string | undefined;
  if (name.length === 0 || name.length > MAX_NAME_LENGTH) problem = `must be 1-${MAX_NAME_LENGTH} characters`;
  else if (!NAME_RE.test(name)) problem = 'must start with a letter, "_" or "\\" and contain only letters, digits, "_", ".", "\\" or "?"';
  else if (isValidCellRef(name) || R1C1_RE.test(name)) problem = 'reads as a cell reference';
  if (problem) throw new OpenXmlSchemaError(`${what} "${name}" ${problem}`);
}
