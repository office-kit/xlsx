// xl/commentsN.xml read/write, plus the VML drawing that carries each note's box.

import { richTextToString } from '../cell/rich-text.js';
import { coordinateToTuple } from '../utils/coordinate.js';
import { escapeXmlAttr, escapeXmlText } from '../utils/escape.js';
import { OpenXmlError, OpenXmlSchemaError } from '../utils/exceptions.js';
import { parseRichString, serializeRichString, type SharedStringEntry } from '../workbook/shared-strings.js';
import { MARKUP_COMPAT_NS, SHEET_MAIN_NS, X16_NS } from '../xml/namespaces.js';
import { parseXml } from '../xml/parser.js';
import { findChild, findChildren, type XmlNode } from '../xml/tree.js';
import type { LegacyComment, LegacyCommentAnchor } from './comments.js';
import { makeLegacyComment } from './comments.js';

const COMMENTS_TAG = `{${SHEET_MAIN_NS}}comments`;
const AUTHORS_TAG = `{${SHEET_MAIN_NS}}authors`;
const AUTHOR_TAG = `{${SHEET_MAIN_NS}}author`;
const COMMENT_LIST_TAG = `{${SHEET_MAIN_NS}}commentList`;
const COMMENT_TAG = `{${SHEET_MAIN_NS}}comment`;
const TEXT_TAG = `{${SHEET_MAIN_NS}}text`;

const VML_NS = 'urn:schemas-microsoft-com:vml';
const VML_EXCEL_NS = 'urn:schemas-microsoft-com:office:excel';
const VML_SHAPE_TAG = `{${VML_NS}}shape`;
const CLIENT_DATA_TAG = `{${VML_EXCEL_NS}}ClientData`;
const ANCHOR_TAG = `{${VML_EXCEL_NS}}Anchor`;
const VISIBLE_TAG = `{${VML_EXCEL_NS}}Visible`;
const ROW_TAG = `{${VML_EXCEL_NS}}Row`;
const COLUMN_TAG = `{${VML_EXCEL_NS}}Column`;

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeText = escapeXmlText;
const escapeAttr = escapeXmlAttr;

/** Parse a `xl/commentsN.xml` payload into a flat LegacyComment list. */
export function parseCommentsXml(bytes: Uint8Array | string): LegacyComment[] {
  const root = parseXml(bytes);
  if (root.name !== COMMENTS_TAG) {
    throw new OpenXmlSchemaError(`parseCommentsXml: root is "${root.name}", expected comments`);
  }
  const authors: string[] = [];
  const authorsEl = findChild(root, AUTHORS_TAG);
  if (authorsEl) {
    for (const a of findChildren(authorsEl, AUTHOR_TAG)) authors.push(a.text ?? '');
  }
  const out: LegacyComment[] = [];
  const listEl = findChild(root, COMMENT_LIST_TAG);
  if (!listEl) return out;
  for (const c of findChildren(listEl, COMMENT_TAG)) {
    const ref = c.attrs['ref'];
    const authorIdRaw = c.attrs['authorId'];
    if (!ref) throw new OpenXmlSchemaError('parseCommentsXml: <comment> missing @ref');
    const authorId = authorIdRaw !== undefined ? Number.parseInt(authorIdRaw, 10) : 0;
    const author = authors[authorId] ?? '';
    const textEl = findChild(c, TEXT_TAG);
    // <text> is a CT_Rst, the same body as a shared string.
    const body = textEl ? parseRichString(textEl) : '';
    out.push(
      typeof body === 'string'
        ? makeLegacyComment({ ref, author, text: body })
        : makeLegacyComment({ ref, author, text: richTextToString(body.runs), richText: body.runs }),
    );
  }
  return out;
}

/**
 * A comment as written: `uid` (`xr:uid`) is set on threaded-comment
 * placeholders, which Excel ties back to their thread by it.
 */
type CommentEntry = LegacyComment & { uid?: string };

/**
 * Serialise a LegacyComment array to a `xl/commentsN.xml` payload. Authors are
 * deduped: each unique `author` becomes one `<author>` entry, and comments
 * reference it by index.
 */
export function commentsToBytes(comments: ReadonlyArray<CommentEntry>): Uint8Array {
  return new TextEncoder().encode(serializeComments(comments));
}

export function serializeComments(comments: ReadonlyArray<CommentEntry>): string {
  const authorIndex = new Map<string, number>();
  const authors: string[] = [];
  for (const c of comments) {
    if (!authorIndex.has(c.author)) {
      authorIndex.set(c.author, authors.length);
      authors.push(c.author);
    }
  }
  const withUid = comments.some((c) => c.uid !== undefined);
  const rootNs = withUid ? ` xmlns:mc="${MARKUP_COMPAT_NS}" mc:Ignorable="xr" xmlns:xr="${X16_NS}"` : '';
  const parts: string[] = [XML_HEADER, `<comments xmlns="${SHEET_MAIN_NS}"${rootNs}><authors>`];
  for (const a of authors) parts.push(`<author>${escapeText(a)}</author>`);
  parts.push('</authors><commentList>');
  for (const c of comments) {
    const id = authorIndex.get(c.author) ?? 0;
    const uid = c.uid !== undefined ? ` shapeId="0" xr:uid="${escapeAttr(c.uid)}"` : '';
    const body: SharedStringEntry =
      c.richText !== undefined && richTextToString(c.richText) === c.text ? { kind: 'rich-text', runs: c.richText } : c.text;
    parts.push(`<comment ref="${escapeAttr(c.ref)}" authorId="${id}"${uid}><text>${serializeRichString(body)}</text></comment>`);
  }
  parts.push('</commentList></comments>');
  return parts.join('');
}

/** The parts of a note's VML shape that {@link commentsVmlDrawing} writes back. */
type CommentBox = Pick<LegacyComment, 'visible' | 'anchor'>;

const boxKey = (row: number, col: number): string => `${row}:${col}`;

/**
 * The zero-based row / column VML `<x:Row>` / `<x:Column>` give a note: the
 * top-left cell of its ref, which may be a range.
 */
const noteCell = (ref: string): { row: number; col: number } | undefined => {
  try {
    const t = coordinateToTuple(ref.split(':')[0] ?? ref);
    return { row: t.row - 1, col: t.col - 1 };
  } catch (err) {
    if (err instanceof OpenXmlError) return undefined;
    throw err;
  }
};

const ANCHOR_FIELDS = [
  'fromCol', 'fromColOffset', 'fromRow', 'fromRowOffset', 'toCol', 'toColOffset', 'toRow', 'toRowOffset',
] as const satisfies ReadonlyArray<keyof LegacyCommentAnchor>;

const UINT = /^\d+$/;

const parseAnchor = (raw: string): LegacyCommentAnchor | undefined => {
  const parts = raw.split(',').map((part) => part.trim());
  if (parts.length !== ANCHOR_FIELDS.length || !parts.every((part) => UINT.test(part))) return undefined;
  const nums = parts.map(Number);
  const at = (i: number): number => nums[i] ?? 0;
  return {
    fromCol: at(0), fromColOffset: at(1), fromRow: at(2), fromRowOffset: at(3),
    toCol: at(4), toColOffset: at(5), toRow: at(6), toRowOffset: at(7),
  };
};

/** Anchor values go into the VML verbatim, so anything but a non-negative integer is refused. */
const anchorText = (anchor: LegacyCommentAnchor, ref: string): string =>
  ANCHOR_FIELDS.map((field) => {
    const v = anchor[field];
    if (!Number.isSafeInteger(v) || v < 0) {
      throw new OpenXmlSchemaError(`comment ${ref}: anchor ${field} must be a non-negative integer; got ${String(v)}`);
    }
    return v;
  }).join(', ');

/**
 * Read each note's box out of a comment VML part and copy it onto the comment
 * on the same cell.
 *
 * The box is presentation only: Excel lays out a default one when it is
 * missing. So a VML part the XML parser rejects (Excel 2003-era files put bare
 * `<br>` in the text box; a DOCTYPE lands here too, unexpanded) or an
 * unreadable `<x:Anchor>` yields no box rather than failing a load that never
 * depended on it. The part itself is regenerated on save, not carried.
 */
export function applyCommentVml(comments: LegacyComment[], bytes: Uint8Array): void {
  let root: XmlNode;
  try {
    root = parseXml(bytes);
  } catch (err) {
    if (err instanceof OpenXmlSchemaError) return;
    throw err;
  }
  const boxes = new Map<string, CommentBox>();
  for (const shape of findChildren(root, VML_SHAPE_TAG)) {
    const data = findChild(shape, CLIENT_DATA_TAG);
    if (data?.attrs['ObjectType'] !== 'Note') continue;
    const row = findChild(data, ROW_TAG)?.text?.trim();
    const col = findChild(data, COLUMN_TAG)?.text?.trim();
    if (row === undefined || col === undefined || !UINT.test(row) || !UINT.test(col)) continue;
    const box: CommentBox = {};
    // x:Visible is a VML boolean whose empty content means true.
    const visible = findChild(data, VISIBLE_TAG);
    if (visible) box.visible = !/^(f|false|0)$/i.test(visible.text?.trim() ?? '');
    const anchorText = findChild(data, ANCHOR_TAG)?.text;
    const anchor = anchorText !== undefined ? parseAnchor(anchorText) : undefined;
    if (anchor) box.anchor = anchor;
    boxes.set(boxKey(Number(row), Number(col)), box);
  }
  for (const c of comments) {
    const cell = noteCell(c.ref);
    const box = cell && boxes.get(boxKey(cell.row, cell.col));
    if (box?.visible !== undefined) c.visible = box.visible;
    if (box?.anchor !== undefined) c.anchor = box.anchor;
  }
}

/**
 * The VML drawing paired with a comments part: one note shape per comment.
 *
 * Excel won't open the file unless every legacy comment is paired with a
 * `<v:shape>` carrying `<x:ClientData ObjectType="Note">` plus the cell anchor
 * row/column — without those, Excel reports "removed Records: Comment from
 * /xl/comments1.xml" and silently drops them.
 *
 * The `<v:shapetype id="_x0000_t202">` / `<v:shape type="#_x0000_t202">` pair
 * mirrors what openpyxl's ShapeWriter emits.
 */
export function commentsVmlDrawing(comments: ReadonlyArray<LegacyComment> = []): Uint8Array {
  const parts: string[] = [
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
    '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
    '<v:stroke joinstyle="miter"/>',
    '<v:path gradientshapeok="t" o:connecttype="rect"/>',
    '</v:shapetype>',
  ];
  // Comment shape ids start at 1026 (matches openpyxl/Excel convention).
  let shapeId = 1026;
  if (comments.length === 0) {
    parts.push('<v:shape type="#_x0000_t202" style="visibility:hidden"><x:ClientData ObjectType="Note"/></v:shape>');
  } else {
    for (const c of comments) {
      // Leave at 0,0 if the ref can't be parsed.
      const { row, col } = noteCell(c.ref) ?? { row: 0, col: 0 };
      const idAttr = `_x0000_s${String(shapeId).padStart(4, '0')}`;
      shapeId++;
      const visibility = c.visible ? 'visible' : 'hidden';
      parts.push(
        `<v:shape id="${idAttr}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:${visibility}" fillcolor="#ffffe1" o:insetmode="auto">`,
        '<v:fill color2="#ffffe1"/>',
        '<v:shadow color="black" obscured="t"/>',
        '<v:path o:connecttype="none"/>',
        '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>',
        '<x:ClientData ObjectType="Note">',
        '<x:MoveWithCells/>',
        '<x:SizeWithCells/>',
      );
      // Child order follows what Excel writes.
      if (c.anchor) parts.push(`<x:Anchor>${anchorText(c.anchor, c.ref)}</x:Anchor>`);
      parts.push('<x:AutoFill>False</x:AutoFill>', `<x:Row>${row}</x:Row>`, `<x:Column>${col}</x:Column>`);
      if (c.visible) parts.push('<x:Visible/>');
      parts.push('</x:ClientData>', '</v:shape>');
    }
  }
  parts.push('</xml>');
  return new TextEncoder().encode(parts.join(''));
}
