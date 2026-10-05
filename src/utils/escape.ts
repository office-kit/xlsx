// XML 1.0 cell-string escaping. Mirrors openpyxl/openpyxl/utils/escape.py.
//
// Excel emits cell strings with control characters and other illegal
// XML 1.0 codepoints encoded as `_xHHHH_` (uppercase hex). The
// underscore itself is a literal in normal text but a sequence opener
// in escape position; an underscore that opens one in the input is
// therefore written as `_x005F_` so the text round-trips losslessly.

import { OpenXmlSchemaError } from './exceptions.js';

// Every codepoint XML 1.0 cannot carry: the C0 controls, an unpaired
// surrogate (which has no UTF-8 encoding), and U+FFFE / U+FFFF. A character
// reference is no way out either, since `&#0;` is as illegal as the raw byte.
// This range deliberately covers `\t` (U+0009), `\n` (U+000A) and `\r`
// (U+000D) even though XML 1.0 considers them legal whitespace: XML
// parsers normalise CRLF / lone CR to LF on read, so a cell string
// containing `\r` would silently lose its CR without the `_x000D_`
// encoding. openpyxl escapes the same `\x01-\x19` range; we add NUL
// to keep the writer well-formed when callers feed in binary data.
//
// The `u` flag makes the class match code points rather than UTF-16 code
// units, so a well-formed surrogate pair is a single character outside
// D800-DFFF and passes through intact.
//
// The class lives in a string so the lookahead below can share it. A codepoint
// this pass escapes but the lookahead does not know about is how text grows a
// sequence the reader then decodes.
const ILLEGAL_CLASS = '[\\x00-\\x1F\\u{D800}-\\u{DFFF}\\u{FFFE}\\u{FFFF}]';
const ILLEGAL_RE = new RegExp(ILLEGAL_CLASS, 'gu');

// An underscore opens a sequence when four hex digits and a closer follow it.
// The closer is either a literal underscore or an illegal codepoint, since
// escaping that codepoint puts an underscore in its place: `SKU_x0041` + `\n`
// goes out as `SKU_x0041_x000A_`, whose `_x0041_` a reader decodes to `A`.
// Matching the opener alone leaves the closer free to open the next sequence,
// which is how `_x0041_x0042_` gets both of its halves protected.
const SEQUENCE_OPENER_RE = new RegExp(`_(?=x[0-9A-Fa-f]{4}(?:_|${ILLEGAL_CLASS}))`, 'gu');

const toHex4 = (n: number): string => n.toString(16).toUpperCase().padStart(4, '0');

/**
 * Escape a string for safe storage in an OOXML cell. An underscore that opens
 * an `_xHHHH_` sequence becomes `_x005F_`; illegal codepoints are replaced
 * with their `_xHHHH_` representation. {@link unescapeCellString} recovers the
 * input for any string.
 */
export function escapeCellString(s: string): string {
  // Protection runs first: the sequences the illegal pass emits are the
  // writer's own and have to stay decodable.
  const protectedString = s.replace(SEQUENCE_OPENER_RE, '_x005F_');
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
 * collapses them to spaces, but the DOM read path (`parseXml`) does not
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

// The codepoints XML 1.0 cannot carry in a text node under any encoding: the
// C0 controls apart from tab / LF / CR, unpaired surrogates, and U+FFFE / U+FFFF. A character
// reference is no help either, since `&#0;` is as illegal as the raw byte.
// Matching on code points (`u` flag) is what keeps a well-formed surrogate
// pair, and so every astral character, out of `\p{Cs}`.
const UNREPRESENTABLE_RE =
  // biome-ignore lint/suspicious/noControlCharactersInRegex: by design, these are the codepoints being rejected
  /[\0-\x08\x0B\x0C\x0E-\x1F\p{Cs}\uFFFE\uFFFF]/u;

/**
 * Escape text for an OOXML text node the reader hands back verbatim: `<f>`,
 * `<formula>`, `<formula1>`, `<formula2>` and a `t="str"` cached formula
 * result. The `_xHHHH_` convention of {@link escapeCellString} must not be
 * applied to these: nothing decodes them on read, so an encoded sequence would
 * reach Excel as the literal text `_x000D_`.
 *
 * Tab and LF stay literal; CR becomes a numeric character reference, which
 * survives the CR-to-LF normalisation every conforming parser applies to raw
 * text. A codepoint XML 1.0 cannot represent therefore has nowhere to go and is
 * rejected rather than silently rewritten.
 *
 * `node` and `at` name the offending value in that error (`worksheet: formula`
 * / `A1`). They stay separate so that a formula-heavy sheet does not build one
 * message string per cell for an error it never raises.
 */
export function escapeXmlTextVerbatim(s: string, node: string, at: string): string {
  const unrepresentable = UNREPRESENTABLE_RE.exec(s);
  if (unrepresentable) {
    const codePoint = toHex4(s.charCodeAt(unrepresentable.index));
    throw new OpenXmlSchemaError(`${node} at ${at} contains U+${codePoint}, which XML 1.0 cannot represent`);
  }
  return escapeXmlText(s).replace(/\r/g, '&#13;');
}
