// Inline rich-text runs.
//
// A rich-text cell value is `{ kind: 'rich-text', runs }`. Each run is
// a string segment with an optional InlineFont describing the in-line
// formatting (font name, size, bold / italic / underline, colour, …).
// Run-level fields use OOXML's short attribute names (`sz`, `b`, `i`,
// `u`) so the writer can splice them into `<rPr>` directly without
// renaming.

import type { Color } from '../styles/colors.js';
import type { FontScheme } from '../styles/fonts.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';

/** Underline styles per openpyxl's cell-level NestedNoneSet. */
export type InlineUnderline = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

export type InlineVertAlign = 'baseline' | 'superscript' | 'subscript';

export interface InlineFont {
  readonly name?: string;
  readonly sz?: number;
  readonly b?: boolean;
  readonly i?: boolean;
  readonly u?: InlineUnderline;
  readonly strike?: boolean;
  readonly outline?: boolean;
  readonly shadow?: boolean;
  readonly condense?: boolean;
  readonly extend?: boolean;
  readonly vertAlign?: InlineVertAlign;
  readonly color?: Color;
  readonly family?: number;
  readonly charset?: number;
  readonly scheme?: FontScheme;
}

export interface TextRun {
  readonly text: string;
  readonly font?: InlineFont;
}

/** A frozen array of TextRuns. The shared cell value shape under `kind: 'rich-text'`. */
export type RichText = ReadonlyArray<TextRun>;

export function makeTextRun(text: string, font?: InlineFont): TextRun {
  if (typeof text !== 'string') {
    throw new OpenXmlSchemaError('makeTextRun: text must be a string');
  }
  return Object.freeze(font !== undefined ? { text, font } : { text });
}

export function makeRichText(runs: ReadonlyArray<TextRun | { text: string; font?: InlineFont }>): RichText {
  const out = runs.map((r) => (Object.isFrozen(r) ? (r as TextRun) : makeTextRun(r.text, r.font)));
  return Object.freeze(out);
}

/**
 * Concatenate the plain-text content of a rich-text value (rich-text
 * read paths often want the raw text without formatting).
 */
export function richTextToString(rt: RichText): string {
  let out = '';
  for (const r of rt) out += r.text;
  return out;
}
