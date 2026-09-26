// Formula-text normalisation, shared by every path that carries a formula.
//
// A spreadsheet UI shows `=SUM(A1:A3)`, but OOXML stores the text without the
// leading `=` (ECMA-376 §18.3.1.40) in `<f>`, `<formula1>`, `<cfRule>`'s
// `<formula>` and `<definedName>` alike, and Excel reports a file that carries
// one as damaged. Exactly one `=` comes off, because exactly one is what the UI
// spelling adds, which leaves the question of what to do with `'==A1'`. The two
// boundaries want opposite answers, so they get one function each:
//
// - a call (`make*` constructor, or a hand-built value reaching a serialiser)
//   is rejected. `'==A1'` is invalid in Excel's formula bar too, and stripping
//   the second `=` would silently store `A1`, a different formula.
// - a file is repaired. A workbook must not fail to open over one malformed
//   `<f>`, and the text was never valid stored text to begin with.

import { quoteCellText } from './cell-text.js';
import { OpenXmlSchemaError } from './exceptions.js';

const EQUALS_SIGN_CODE = 61;

/** True iff `text` is spelled the way a spreadsheet UI shows a formula. */
function startsWithEquals(text: string): boolean {
  return text.charCodeAt(0) === EQUALS_SIGN_CODE;
}

/**
 * True iff Excel reads `text`, typed into a cell, as a formula. A lone `=` is
 * the exception to the prefix rule: Excel stores it as text, where `'= '` is
 * refused like any other malformed formula.
 */
export function spellsFormula(text: string): boolean {
  return startsWithEquals(text) && text.length > 1;
}

/** Drop one leading `=` and the whitespace around it: `' =SUM(A1)'` → `'SUM(A1)'`. */
const stripPrefix = (text: string): string => {
  const trimmed = text.trim();
  return startsWithEquals(trimmed) ? trimmed.slice(1).trim() : trimmed;
};

/**
 * Normalise formula text a caller supplied. `context` names the call or the
 * element being written, and `at`, when given, where it sits, so a rejection
 * says where the text came from. They stay apart so the writer's hot path
 * builds the message only when it throws.
 *
 * @throws OpenXmlSchemaError if the text still begins with `=` once the leading
 * one is gone. Writing it as-is stores `<f>=A1</f>`, the shape this function
 * exists to keep out of the file, and stripping again invents a reading for
 * input that has none.
 */
export function normalizeFormulaText(text: string, context: string, at?: string): string {
  const body = stripPrefix(text);
  if (startsWithEquals(body)) {
    throw new OpenXmlSchemaError(
      `${context}${at === undefined ? '' : ` at ${at}`}: formula text still begins with "=" once the leading one is removed` +
        ` (got "${quoteCellText(text)}"). OOXML stores formula text without it and Excel` +
        ' reports a file that carries one as damaged',
    );
  }
  return body;
}

/**
 * Normalise formula text read out of a file, which carries no `=` when the
 * producer was correct. Every leading `=` comes off rather than throwing:
 * refusing the load would cost the whole workbook, and no reading of the rest
 * of the expression changes with the prefix gone.
 */
export function repairFormulaTextFromFile(text: string): string {
  let body = text.trim();
  while (startsWithEquals(body)) body = body.slice(1).trim();
  return body;
}
