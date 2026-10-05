// Threaded comments (Excel 365's "Comments", as opposed to legacy "Notes").
//
// A thread is a root comment plus the replies whose `parentId` names it. They
// live in `xl/threadedComments/threadedCommentN.xml` and reference their
// authors by id through the workbook-wide person list (`Workbook.persons`).
//
// Excel also writes every thread as a legacy comment whose author is
// `tc=<root id>`, so readers that predate threaded comments still show the
// conversation. Those placeholders are derived at save time and dropped at load
// time; they are never part of `Worksheet.legacyComments`.

/** An `@mention` inside a comment's text. */
export interface ThreadedCommentMention {
  /** Id of the mentioned person in `Workbook.persons`. */
  personId: string;
  /** Id of the mention itself, unique within the workbook. */
  mentionId: string;
  /** Offset of the mention within the comment text, in UTF-16 code units. */
  startIndex: number;
  length: number;
}

export interface ThreadedComment {
  /** GUID in braces, e.g. `{9C3F…}`; replies point at their root by this id. */
  id: string;
  /** The single cell the thread is attached to. Replies repeat their root's ref. */
  ref: string;
  /** Id of the author in `Workbook.persons`. */
  personId: string;
  /** Creation time as written by Excel (`xsd:dateTime`, e.g. `2024-05-01T09:30:00.00`). */
  created?: string;
  text: string;
  /** Root comment id; absent on the root itself. */
  parentId?: string;
  /** Resolved state. Only meaningful on the root comment. */
  done?: boolean;
  mentions?: ThreadedCommentMention[];
}

/** A fresh `{GUID}` in the upper-case, braced form Excel writes. */
export const newOfficeGuid = (): string => `{${globalThis.crypto.randomUUID().toUpperCase()}}`;

/**
 * Build a threaded comment with a new id. `created` defaults to now, written
 * in UTC with centiseconds and no zone suffix, which is what Excel writes.
 */
export function makeThreadedComment(opts: {
  ref: string;
  personId: string;
  text: string;
  parentId?: string;
  created?: Date;
}): ThreadedComment {
  const out: ThreadedComment = {
    id: newOfficeGuid(),
    ref: opts.ref,
    personId: opts.personId,
    created: (opts.created ?? new Date()).toISOString().slice(0, 22),
    text: opts.text,
  };
  if (opts.parentId !== undefined) out.parentId = opts.parentId;
  return out;
}
