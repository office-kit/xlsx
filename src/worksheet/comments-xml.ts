// xl/commentsN.xml read/write.

import { escapeXmlAttr, escapeXmlText } from '../utils/escape.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { MARKUP_COMPAT_NS, SHEET_MAIN_NS, X16_NS } from '../xml/namespaces.js';
import { parseXml } from '../xml/parser.js';
import { findChild, findChildren, type XmlNode } from '../xml/tree.js';
import { coordinateToTuple } from '../utils/coordinate.js';
import type { LegacyComment } from './comments.js';
import { makeLegacyComment } from './comments.js';

const COMMENTS_TAG = `{${SHEET_MAIN_NS}}comments`;
const AUTHORS_TAG = `{${SHEET_MAIN_NS}}authors`;
const AUTHOR_TAG = `{${SHEET_MAIN_NS}}author`;
const COMMENT_LIST_TAG = `{${SHEET_MAIN_NS}}commentList`;
const COMMENT_TAG = `{${SHEET_MAIN_NS}}comment`;
const TEXT_TAG = `{${SHEET_MAIN_NS}}text`;
const T_TAG = `{${SHEET_MAIN_NS}}t`;
const R_TAG = `{${SHEET_MAIN_NS}}r`;

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeText = escapeXmlText;
const escapeAttr = escapeXmlAttr;

/** Concatenate every `<t>` text node found inside a `<text>` body. */
const collectText = (textEl: XmlNode): string => {
  const direct = findChild(textEl, T_TAG);
  if (direct) return direct.text ?? '';
  let out = '';
  for (const child of textEl.children) {
    if (child.name === R_TAG) {
      const t = findChild(child, T_TAG);
      if (t?.text) out += t.text;
    } else if (child.name === T_TAG && child.text) {
      out += child.text;
    }
  }
  return out;
};

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
    const text = textEl ? collectText(textEl) : '';
    out.push(makeLegacyComment({ ref, author, text }));
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
    parts.push(
      `<comment ref="${escapeAttr(c.ref)}" authorId="${id}"${uid}><text><t>${escapeText(c.text)}</t></text></comment>`,
    );
  }
  parts.push('</commentList></comments>');
  return parts.join('');
}

/**
 * Bare-bones VML drawing payload Excel tolerates as a comment-shape
 * placeholder. We don't preserve the original VML shapes (stage-1 trade-off);
 * this stub guarantees the worksheet rels stay consistent.
 *
 * Excel won't open the file unless every legacy comment is paired with a
 * `<v:shape>` carrying `<x:ClientData ObjectType="Note">` plus the cell anchor
 * row/column — without those, Excel reports "removed Records: Comment from
 * /xl/comments1.xml" and silently drops them.
 *
 * The `<v:shapetype id="_x0000_t202">` / `<v:shape type="#_x0000_t202">` pair
 * mirrors what openpyxl's ShapeWriter emits.
 */
export function placeholderVmlDrawing(comments: ReadonlyArray<LegacyComment> = []): Uint8Array {
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
      // The ref may be a single cell or a range; anchor at the top-left.
      const firstRef = c.ref.split(':')[0] ?? c.ref;
      let row = 0;
      let col = 0;
      try {
        const t = coordinateToTuple(firstRef);
        row = t.row - 1; // VML uses 0-based row/column.
        col = t.col - 1;
      } catch {
        // Leave at 0,0 if the ref can't be parsed.
      }
      const idAttr = `_x0000_s${String(shapeId).padStart(4, '0')}`;
      shapeId++;
      parts.push(
        `<v:shape id="${idAttr}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">`,
        '<v:fill color2="#ffffe1"/>',
        '<v:shadow color="black" obscured="t"/>',
        '<v:path o:connecttype="none"/>',
        '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>',
        '<x:ClientData ObjectType="Note">',
        '<x:MoveWithCells/>',
        '<x:SizeWithCells/>',
        '<x:AutoFill>False</x:AutoFill>',
        `<x:Row>${row}</x:Row>`,
        `<x:Column>${col}</x:Column>`,
        '</x:ClientData>',
        '</v:shape>',
      );
    }
  }
  parts.push('</xml>');
  return new TextEncoder().encode(parts.join(''));
}
