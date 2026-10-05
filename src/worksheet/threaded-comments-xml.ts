// xl/threadedComments/threadedCommentN.xml and xl/persons/person.xml
// read/write, plus the legacy-comment placeholders Excel pairs with each thread.

import type { Person } from '../workbook/persons.js';
import { escapeXmlAttr, escapeXmlText } from '../utils/escape.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { SHEET_MAIN_NS, THREADED_COMMENTS_NS } from '../xml/namespaces.js';
import { parseXml } from '../xml/parser.js';
import { findChild, findChildren } from '../xml/tree.js';
import type { LegacyComment } from './comments.js';
import type { ThreadedComment, ThreadedCommentMention } from './threaded-comments.js';

const THREADED_COMMENTS_TAG = `{${THREADED_COMMENTS_NS}}ThreadedComments`;
const THREADED_COMMENT_TAG = `{${THREADED_COMMENTS_NS}}threadedComment`;
const TEXT_TAG = `{${THREADED_COMMENTS_NS}}text`;
const MENTIONS_TAG = `{${THREADED_COMMENTS_NS}}mentions`;
const MENTION_TAG = `{${THREADED_COMMENTS_NS}}mention`;
const PERSON_LIST_TAG = `{${THREADED_COMMENTS_NS}}personList`;
const PERSON_TAG = `{${THREADED_COMMENTS_NS}}person`;

export const THREADED_COMMENT_REL = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
export const THREADED_COMMENTS_TYPE = 'application/vnd.ms-excel.threadedcomments+xml';
export const PERSON_REL = 'http://schemas.microsoft.com/office/2017/10/relationships/person';
export const PERSON_TYPE = 'application/vnd.ms-excel.person+xml';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ROOT_NS = `xmlns="${THREADED_COMMENTS_NS}" xmlns:x="${SHEET_MAIN_NS}"`;

const requireAttr = (attrs: Readonly<Record<string, string>>, name: string, where: string): string => {
  const v = attrs[name];
  if (v === undefined) throw new OpenXmlSchemaError(`${where}: missing @${name}`);
  return v;
};

const parseIndex = (raw: string, where: string): number => {
  const n = Number.parseInt(raw, 10);
  if (!Number.isInteger(n) || n < 0) throw new OpenXmlSchemaError(`${where}: "${raw}" is not a non-negative integer`);
  return n;
};

export function parseThreadedCommentsXml(bytes: Uint8Array | string): ThreadedComment[] {
  const root = parseXml(bytes);
  if (root.name !== THREADED_COMMENTS_TAG) {
    throw new OpenXmlSchemaError(`parseThreadedCommentsXml: root is "${root.name}", expected ThreadedComments`);
  }
  const out: ThreadedComment[] = [];
  for (const el of findChildren(root, THREADED_COMMENT_TAG)) {
    const where = 'parseThreadedCommentsXml: <threadedComment>';
    const tc: ThreadedComment = {
      id: requireAttr(el.attrs, 'id', where),
      ref: requireAttr(el.attrs, 'ref', where),
      personId: requireAttr(el.attrs, 'personId', where),
      text: findChild(el, TEXT_TAG)?.text ?? '',
    };
    const dT = el.attrs['dT'];
    if (dT !== undefined) tc.created = dT;
    const parentId = el.attrs['parentId'];
    if (parentId !== undefined) tc.parentId = parentId;
    const done = el.attrs['done'];
    if (done === '1' || done === 'true') tc.done = true;
    const mentionsEl = findChild(el, MENTIONS_TAG);
    if (mentionsEl) {
      const mentions: ThreadedCommentMention[] = [];
      for (const m of findChildren(mentionsEl, MENTION_TAG)) {
        const mw = 'parseThreadedCommentsXml: <mention>';
        mentions.push({
          personId: requireAttr(m.attrs, 'mentionpersonId', mw),
          mentionId: requireAttr(m.attrs, 'mentionId', mw),
          startIndex: parseIndex(requireAttr(m.attrs, 'startIndex', mw), mw),
          length: parseIndex(requireAttr(m.attrs, 'length', mw), mw),
        });
      }
      if (mentions.length > 0) tc.mentions = mentions;
    }
    out.push(tc);
  }
  return out;
}

export function threadedCommentsToBytes(comments: ReadonlyArray<ThreadedComment>): Uint8Array {
  const parts: string[] = [XML_HEADER, `<ThreadedComments ${ROOT_NS}>`];
  for (const c of comments) {
    let attrs = ` ref="${escapeXmlAttr(c.ref)}"`;
    if (c.created !== undefined) attrs += ` dT="${escapeXmlAttr(c.created)}"`;
    attrs += ` personId="${escapeXmlAttr(c.personId)}" id="${escapeXmlAttr(c.id)}"`;
    if (c.parentId !== undefined) attrs += ` parentId="${escapeXmlAttr(c.parentId)}"`;
    if (c.done) attrs += ' done="1"';
    parts.push(`<threadedComment${attrs}><text>${escapeXmlText(c.text)}</text>`);
    if (c.mentions && c.mentions.length > 0) {
      parts.push('<mentions>');
      for (const m of c.mentions) {
        parts.push(
          `<mention mentionpersonId="${escapeXmlAttr(m.personId)}" mentionId="${escapeXmlAttr(m.mentionId)}" startIndex="${m.startIndex}" length="${m.length}"/>`,
        );
      }
      parts.push('</mentions>');
    }
    parts.push('</threadedComment>');
  }
  parts.push('</ThreadedComments>');
  return new TextEncoder().encode(parts.join(''));
}

export function parsePersonsXml(bytes: Uint8Array | string): Person[] {
  const root = parseXml(bytes);
  if (root.name !== PERSON_LIST_TAG) {
    throw new OpenXmlSchemaError(`parsePersonsXml: root is "${root.name}", expected personList`);
  }
  const out: Person[] = [];
  for (const el of findChildren(root, PERSON_TAG)) {
    const where = 'parsePersonsXml: <person>';
    const p: Person = {
      displayName: requireAttr(el.attrs, 'displayName', where),
      id: requireAttr(el.attrs, 'id', where),
    };
    const userId = el.attrs['userId'];
    if (userId !== undefined) p.userId = userId;
    const providerId = el.attrs['providerId'];
    if (providerId !== undefined) p.providerId = providerId;
    out.push(p);
  }
  return out;
}

export function personsToBytes(persons: ReadonlyArray<Person>): Uint8Array {
  const parts: string[] = [XML_HEADER, `<personList ${ROOT_NS}>`];
  for (const p of persons) {
    let attrs = ` displayName="${escapeXmlAttr(p.displayName)}" id="${escapeXmlAttr(p.id)}"`;
    if (p.userId !== undefined) attrs += ` userId="${escapeXmlAttr(p.userId)}"`;
    if (p.providerId !== undefined) attrs += ` providerId="${escapeXmlAttr(p.providerId)}"`;
    parts.push(`<person${attrs}/>`);
  }
  parts.push('</personList>');
  return new TextEncoder().encode(parts.join(''));
}

/** Legacy-comment author Excel gives a thread's placeholder: `tc=` + root id. */
export const threadPlaceholderAuthor = (rootId: string): string => `tc=${rootId}`;

// Verbatim from Excel, which shows this to readers without threaded comments.
const PLACEHOLDER_PREAMBLE =
  '[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ';

/**
 * One legacy comment per thread, the way Excel writes them: authored by
 * `tc=<root id>`, tagged with the root id as `uid`, and carrying the root text
 * followed by every reply. Excel matches thread and placeholder by that id;
 * without the placeholder it drops the thread as unreadable.
 */
export function threadPlaceholders(
  comments: ReadonlyArray<ThreadedComment>,
): Array<LegacyComment & { uid: string }> {
  const repliesByRoot = new Map<string, ThreadedComment[]>();
  for (const c of comments) {
    if (c.parentId === undefined) continue;
    const list = repliesByRoot.get(c.parentId);
    if (list) list.push(c);
    else repliesByRoot.set(c.parentId, [c]);
  }
  const out: Array<LegacyComment & { uid: string }> = [];
  for (const root of comments) {
    if (root.parentId !== undefined) continue;
    let text = PLACEHOLDER_PREAMBLE + root.text;
    for (const reply of repliesByRoot.get(root.id) ?? []) text += `\nReply:\n    ${reply.text}`;
    out.push({ ref: root.ref, author: threadPlaceholderAuthor(root.id), text, uid: root.id });
  }
  return out;
}
