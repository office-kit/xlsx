// XML 1.0 cell-string escaping. Mirrors openpyxl/openpyxl/utils/escape.py.
//
// Excel emits cell strings with control characters and other illegal
// XML 1.0 codepoints encoded as `_xHHHH_` (uppercase hex). The
// underscore itself is a literal in normal text but a sequence opener
// in escape position; an existing `_xHHHH_` in the input is therefore
// re-escaped to `_x005F_xHHHH_` so it round-trips losslessly.

// Escape every C0 control character (U+0000 through U+001F) plus the
// surrogate range U+D800–U+DFFF. NUL is invalid in XML 1.0 entirely.
// This range deliberately covers `\t` (U+0009), `\n` (U+000A) and `\r`
// (U+000D) even though XML 1.0 considers them legal whitespace — XML
// parsers normalise CRLF / lone CR to LF on read, so a cell string
// containing `\r` would silently lose its CR without the `_x000D_`
// encoding. openpyxl escapes the same `\x01-\x19` range; we add NUL
// to keep the writer well-formed when callers feed in binary data.
const ILLEGAL_RE =
  // biome-ignore lint/suspicious/noControlCharactersInRegex: by design — these are the codepoints we replace
  /[\x00-\x1F\ud800-\udfff]/g;
const ESCAPED_PATTERN_RE = /(_)(x[0-9A-Fa-f]{4}_)/g;

const toHex4 = (n: number): string => n.toString(16).toUpperCase().padStart(4, '0');

/**
 * Escape a string for safe storage in an OOXML cell. Already-escaped
 * sequences (`_xHHHH_`) are protected by escaping their leading
 * underscore; illegal codepoints are replaced with their `_xHHHH_`
 * representation.
 */
export function escapeCellString(s: string): string {
  // Re-escape any existing `_xHHHH_` so it round-trips; the underscore
  // becomes `_x005F_` and the rest of the sequence is left as-is.
  const protectedString = s.replace(ESCAPED_PATTERN_RE, '_x005F_$2');
  return protectedString.replace(ILLEGAL_RE, (ch) => `_x${toHex4(ch.charCodeAt(0))}_`);
}

const UNESCAPE_RE = /_x([0-9A-Fa-f]{4})_/g;

/**
 * Inverse of {@link escapeCellString}. Looking from left to right
 * we replace any `_xHHHH_` sequence with the corresponding code unit;
 * the protected `_x005F_` becomes a literal underscore which the
 * subsequent replacements skip safely (replace's regex is non-overlapping).
 */
export function unescapeCellString(s: string): string {
  return s.replace(UNESCAPE_RE, (_full, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)));
}

// ---- XML escape helpers ----------------------------------------------------
//
// One canonical implementation for text-node and attribute escaping. The
// previous codebase carried three near-identical copies (one each in save.ts,
// xml/serializer.ts, xml/stream-writer.ts); they disagreed about `>`-in-attribute
// handling and would have drifted further apart over time. Keeping them in a
// single place also makes future fixes (e.g. surrogate-pair scrubbing) land
// once rather than three times.

/**
 * Escape a string for safe placement in an XML text node. Replaces the three
 * codepoints that would otherwise terminate the text region or open a markup
 * sequence (`&`, `<`, `>`); `"` and `'` are not legal markup terminators
 * inside text and stay verbatim.
 */
export function escapeXmlText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Escape a string for safe placement inside a `"`-quoted XML attribute.
 * Handles `&`, `<`, `>` and the `"` that would otherwise close the value.
 *
 * Note: this deliberately does NOT escape `\r` / `\n` / `\t` to numeric
 * character references. XML 1.0 attribute-value normalisation would
 * collapses them to spaces, but the DOM read path (fast-xml-parser) does not
 * apply that normalisation, so the literal bytes round-trip through
 * `loadWorkbook`. Leaving them literal also matches what Excel itself emits.
 */
export function escapeXmlAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
