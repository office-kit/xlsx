// Threaded comments (Review ▸ New Comment), the conversation-style successor
// of notes. A thread is a root comment plus replies pointing at it; authors
// live in the workbook's person list. Every edit is one undo step.

import type { Person, Workbook } from '@office-kit/xlsx/workbook';
import { makePerson } from '@office-kit/xlsx/workbook';
import type { ThreadedComment, Worksheet } from '@office-kit/xlsx/worksheet';
import { makeThreadedComment } from '@office-kit/xlsx/worksheet';
import { columnIndexFromLetter, coordinateFromString } from '@office-kit/xlsx/utils';
import { cellAddress, inRange } from './address.ts';
import type { EditorController } from './controller.svelte.ts';
import type { Transaction } from './history.ts';

export interface CommentThread {
  readonly root: ThreadedComment;
  readonly replies: readonly ThreadedComment[];
  readonly row: number;
  readonly col: number;
}

/** Top-left cell of a comment or note ref. */
export function refCell(ref: string): { row: number; col: number } {
  const p = coordinateFromString(ref.replaceAll('$', '').split(':')[0] ?? 'A1');
  return { row: p.row, col: columnIndexFromLetter(p.column) };
}

/** Every thread on the sheet in reading order (row by row), which is also Previous / Next order. */
export function commentThreads(ws: Worksheet): CommentThread[] {
  const replies = new Map<string, ThreadedComment[]>();
  for (const c of ws.threadedComments ?? []) {
    if (c.parentId === undefined) continue;
    const list = replies.get(c.parentId);
    if (list) list.push(c);
    else replies.set(c.parentId, [c]);
  }
  return (ws.threadedComments ?? [])
    .filter((c) => c.parentId === undefined)
    .map((root) => {
      const { row, col } = refCell(root.ref);
      return { root, replies: replies.get(root.id) ?? [], row, col };
    })
    .sort((a, b) => a.row - b.row || a.col - b.col);
}

export function threadAt(ws: Worksheet, row: number, col: number): CommentThread | undefined {
  return commentThreads(ws).find((t) => t.row === row && t.col === col);
}

export function hasNoteAt(ws: Worksheet, row: number, col: number): boolean {
  return ws.legacyComments.some((c) => {
    const p = refCell(c.ref);
    return p.row === row && p.col === col;
  });
}

export function personName(wb: Workbook, personId: string): string {
  return (wb.persons ?? []).find((p) => p.id === personId)?.displayName ?? '';
}

// ---- current user ---------------------------------------------------------------

const USER_NAME_KEY = 'xl-editor.commentUserName';
const DEFAULT_USER_NAME = 'User';

/** The name new comments are posted under; a per-browser setting. */
export function commentUserName(): string {
  try {
    return localStorage.getItem(USER_NAME_KEY)?.trim() || DEFAULT_USER_NAME;
  } catch {
    // Storage is unavailable in private windows and sandboxed frames.
    return DEFAULT_USER_NAME;
  }
}

export function setCommentUserName(name: string): void {
  try {
    localStorage.setItem(USER_NAME_KEY, name.trim());
  } catch {
    // Without storage the name simply isn't remembered.
  }
}

/** The person entry for the current user, added to the workbook on first use. */
function currentPerson(wb: Workbook, tx: Transaction): Person {
  const name = commentUserName();
  const existing = (wb.persons ?? []).find((p) => p.displayName === name);
  if (existing) return existing;
  tx.workbook('persons');
  const person = makePerson({ displayName: name });
  wb.persons = [...(wb.persons ?? []), person];
  return person;
}

// ---- edits ----------------------------------------------------------------------

export function addComment(ctl: EditorController, row: number, col: number, text: string): void {
  const { wb, ws } = ctl.doc;
  ctl.doc.transact('New Comment', (tx) => {
    tx.sheet(ws, 'threadedComments');
    const person = currentPerson(wb, tx);
    ws.threadedComments = [...(ws.threadedComments ?? []), makeThreadedComment({ ref: cellAddress(row, col), personId: person.id, text })];
  });
}

export function replyToComment(ctl: EditorController, root: ThreadedComment, text: string): void {
  const { wb, ws } = ctl.doc;
  ctl.doc.transact('Reply', (tx) => {
    tx.sheet(ws, 'threadedComments');
    const person = currentPerson(wb, tx);
    const reply = makeThreadedComment({ ref: root.ref, personId: person.id, text, parentId: root.id });
    // Keep a thread's comments together, as Excel writes them.
    const last = (ws.threadedComments ?? []).findLastIndex((c) => c.id === root.id || c.parentId === root.id);
    ws.threadedComments = (ws.threadedComments ?? []).toSpliced(last + 1, 0, reply);
  });
}

function updateComment(ctl: EditorController, label: string, id: string, patch: (c: ThreadedComment) => ThreadedComment): void {
  const ws = ctl.doc.ws;
  ctl.doc.transact(label, (tx) => {
    tx.sheet(ws, 'threadedComments');
    ws.threadedComments = (ws.threadedComments ?? []).map((c) => (c.id === id ? patch(c) : c));
  });
}

export function editComment(ctl: EditorController, id: string, text: string): void {
  updateComment(ctl, 'Edit Comment', id, (c) => ({ ...c, text }));
}

export function setThreadResolved(ctl: EditorController, rootId: string, done: boolean): void {
  updateComment(ctl, done ? 'Resolve Thread' : 'Reopen Thread', rootId, (c) => {
    const next = { ...c };
    if (done) next.done = true;
    else delete next.done;
    return next;
  });
}

/** Delete one reply; deleting a root takes its replies with it. */
export function deleteComment(ctl: EditorController, comment: ThreadedComment): void {
  const ws = ctl.doc.ws;
  const rootId = comment.parentId === undefined ? comment.id : undefined;
  ctl.doc.transact(rootId ? 'Delete Thread' : 'Delete Reply', (tx) => {
    tx.sheet(ws, 'threadedComments');
    ws.threadedComments = (ws.threadedComments ?? []).filter((c) => c.id !== comment.id && (rootId === undefined || c.parentId !== rootId));
  });
}

/** Review ▸ Delete: every thread whose cell is in the selection. */
export function deleteCommentsInSelection(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const ranges = ctl.doc.selection.ranges;
  const doomed = new Set(commentThreads(ws).filter((t) => ranges.some((r) => inRange(r, t.row, t.col))).map((t) => t.root.id));
  if (doomed.size === 0) return;
  ctl.doc.transact('Delete Comment', (tx) => {
    tx.sheet(ws, 'threadedComments');
    ws.threadedComments = (ws.threadedComments ?? []).filter((c) => !doomed.has(c.parentId ?? c.id));
  });
}

// ---- navigation -----------------------------------------------------------------

/** Insert ▸ New Comment: open the card on the active cell, composing unless a thread is already there. */
export function newComment(ctl: EditorController): void {
  const { row, col } = ctl.doc.selection.active;
  // A cell holds a note or a thread, never both; Excel greys New Comment out.
  if (hasNoteAt(ctl.doc.ws, row, col)) {
    ctl.toast = 'cmtNoteHere';
    return;
  }
  ctl.commentCard = { row, col };
}

export function stepComment(ctl: EditorController, step: 1 | -1): void {
  const threads = commentThreads(ctl.doc.ws);
  if (threads.length === 0) return;
  const a = ctl.doc.selection.active;
  const next =
    step === 1
      ? (threads.find((p) => p.row > a.row || (p.row === a.row && p.col > a.col)) ?? threads[0])
      : (threads.findLast((p) => p.row < a.row || (p.row === a.row && p.col < a.col)) ?? threads.at(-1));
  if (!next) return;
  ctl.selectCell({ row: next.row, col: next.col });
  ctl.reveal(next.row, next.col);
  ctl.commentCard = { row: next.row, col: next.col };
}

/** Excel's card timestamp; `created` is UTC written without a zone suffix. */
export function formatCommentTime(created: string | undefined, locale: string): string {
  if (!created) return '';
  const d = new Date(/[zZ]|[+-]\d\d:\d\d$/.test(created) ? created : `${created}Z`);
  if (Number.isNaN(d.getTime())) return '';
  return d.toLocaleString(locale, { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit' });
}
