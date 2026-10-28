// Legacy comments.
//
// Legacy comments live in `xl/commentsN.xml` (typed comment list + authors
// index) plus a VML drawing part for the popup balloon shape. The VML is
// regenerated on save from the fields below; the rest of the original shape
// (fill, shadow, CSS margins) is not carried, and Excel places the box by
// `<x:Anchor>` rather than by those margins.

import { type RichText, richTextToString } from '../cell/rich-text.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';

/**
 * Where a note's box sits, as VML `<x:Anchor>` writes it: each edge is a
 * zero-based column or row plus an offset into it in pixels.
 */
export interface LegacyCommentAnchor {
  fromCol: number;
  fromColOffset: number;
  fromRow: number;
  fromRowOffset: number;
  toCol: number;
  toColOffset: number;
  toRow: number;
  toRowOffset: number;
}

export interface LegacyComment {
  /** Cell reference — typically a single cell ("A1") but Excel allows ranges. */
  ref: string;
  /** Display name of the comment author. */
  author: string;
  /** Plain-text body. This is the note's content; `richText` only formats it. */
  text: string;
  /**
   * Run formatting for `text`, such as the bold author line Excel writes. It
   * is saved only while its runs join to exactly `text`, so changing `text`
   * alone saves the note as plain text instead of reviving stale runs.
   */
  richText?: RichText;
  /** `<x:Visible/>`: the box is always shown, not only on hover. */
  visible?: boolean;
  /**
   * The box's position and size. Without it Excel places a default-sized box
   * next to the cell. It does not follow `ref`, so drop it when moving a note.
   */
  anchor?: LegacyCommentAnchor;
}

/**
 * Build a note. Throws `OpenXmlSchemaError` when `richText` is given but its
 * runs do not join to `text`: the writer would drop them.
 */
export function makeLegacyComment(opts: LegacyComment): LegacyComment {
  if (opts.richText !== undefined && richTextToString(opts.richText) !== opts.text) {
    throw new OpenXmlSchemaError(`comment ${opts.ref}: richText runs must join to text`);
  }
  return {
    ref: opts.ref,
    author: opts.author,
    text: opts.text,
    ...(opts.richText !== undefined ? { richText: opts.richText } : {}),
    ...(opts.visible !== undefined ? { visible: opts.visible } : {}),
    ...(opts.anchor !== undefined ? { anchor: opts.anchor } : {}),
  };
}
