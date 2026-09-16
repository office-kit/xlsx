// Formula-text normalisation, shared by every path that carries a formula.
//
// A spreadsheet UI shows `=SUM(A1:A3)`, but OOXML stores the text without the
// leading `=` (ECMA-376 §18.3.1.40) in `<f>`, `<formula1>`, `<cfRule>`'s
// `<formula>` and `<definedName>` alike, and Excel reports a file that carries
// one as damaged. The `make*` constructors normalise what a caller hands them,
// and the serialisers normalise again: the model interfaces are public, so a
// caller can build one as a literal and reach the writer without passing a
// constructor.

const EQUALS_SIGN_CODE = 61;

/** True iff `text` is spelled the way a spreadsheet UI shows a formula. */
export function startsWithEquals(text: string): boolean {
  return text.charCodeAt(0) === EQUALS_SIGN_CODE;
}

/**
 * Drop the leading `=` OOXML formula text must not carry, and the whitespace
 * around it: `' =SUM(A1)'` would otherwise keep the `=` in the stored text.
 */
export function normalizeFormulaText(text: string): string {
  const trimmed = text.trim();
  return startsWithEquals(trimmed) ? trimmed.slice(1).trim() : trimmed;
}
