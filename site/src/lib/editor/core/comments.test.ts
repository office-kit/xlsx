import { describe, expect, test } from 'vitest';
import { fromArrayBuffer, loadWorkbook } from '@office-kit/xlsx/io';
import { makeLegacyComment } from '@office-kit/xlsx/worksheet';
import * as A from './actions.ts';
import {
  addComment,
  commentThreads,
  deleteComment,
  deleteCommentsInSelection,
  editComment,
  formatCommentTime,
  newComment,
  replyToComment,
  setThreadResolved,
  stepComment,
  threadAt,
} from './comments.ts';
import { EditorController } from './controller.svelte.ts';

function rootAt(ctl: EditorController, row: number, col: number) {
  const th = threadAt(ctl.doc.ws, row, col);
  if (!th) throw new Error(`no thread at ${row},${col}`);
  return th.root;
}

describe('threaded comments', () => {
  test('a new comment adds the current user once and is one undo step', () => {
    const ctl = new EditorController();
    addComment(ctl, 2, 2, 'First');
    replyToComment(ctl, rootAt(ctl, 2, 2), 'Second');
    expect(ctl.doc.wb.persons?.map((p) => p.displayName)).toEqual(['User']);
    const th = threadAt(ctl.doc.ws, 2, 2);
    expect(th?.root.text).toBe('First');
    expect(th?.replies.map((r) => r.text)).toEqual(['Second']);

    A.undo(ctl);
    expect(threadAt(ctl.doc.ws, 2, 2)?.replies).toEqual([]);
    A.undo(ctl);
    expect(ctl.doc.ws.threadedComments).toEqual([]);
    expect(ctl.doc.wb.persons).toEqual([]);
    A.redo(ctl);
    expect(threadAt(ctl.doc.ws, 2, 2)?.root.text).toBe('First');
  });

  test('edit, resolve and delete', () => {
    const ctl = new EditorController();
    addComment(ctl, 1, 1, 'Root');
    const root = rootAt(ctl, 1, 1);
    replyToComment(ctl, root, 'Reply');
    editComment(ctl, root.id, 'Root, edited');
    setThreadResolved(ctl, root.id, true);
    expect(rootAt(ctl, 1, 1)).toMatchObject({ text: 'Root, edited', done: true });
    setThreadResolved(ctl, root.id, false);
    expect(rootAt(ctl, 1, 1).done).toBeUndefined();

    const reply = threadAt(ctl.doc.ws, 1, 1)?.replies[0];
    if (!reply) throw new Error('missing reply');
    deleteComment(ctl, reply);
    expect(threadAt(ctl.doc.ws, 1, 1)?.replies).toEqual([]);
    deleteComment(ctl, rootAt(ctl, 1, 1));
    expect(ctl.doc.ws.threadedComments).toEqual([]);
  });

  test('Review ▸ Delete removes the threads in the selection only', () => {
    const ctl = new EditorController();
    addComment(ctl, 1, 1, 'keep');
    addComment(ctl, 3, 3, 'drop');
    replyToComment(ctl, rootAt(ctl, 3, 3), 'drop too');
    ctl.selectCell({ row: 3, col: 3 });
    deleteCommentsInSelection(ctl);
    expect(ctl.doc.ws.threadedComments?.map((c) => c.text)).toEqual(['keep']);
  });

  test('Previous / Next walk threads in reading order and open the card', () => {
    const ctl = new EditorController();
    addComment(ctl, 5, 1, 'c');
    addComment(ctl, 1, 3, 'a');
    addComment(ctl, 2, 1, 'b');
    expect(commentThreads(ctl.doc.ws).map((t) => t.root.text)).toEqual(['a', 'b', 'c']);
    ctl.selectCell({ row: 1, col: 1 });
    stepComment(ctl, 1);
    expect(ctl.doc.selection.active).toEqual({ row: 1, col: 3 });
    expect(ctl.commentCard).toEqual({ row: 1, col: 3 });
    stepComment(ctl, -1);
    expect(ctl.doc.selection.active).toEqual({ row: 5, col: 1 });
  });

  test('deleting a row takes its threads, inserting one shifts them', () => {
    const ctl = new EditorController();
    addComment(ctl, 3, 2, 'on B3');
    ctl.selectCell({ row: 1, col: 1 });
    A.insertLines(ctl, 'row');
    expect(ctl.doc.ws.threadedComments?.[0]?.ref).toBe('B4');
  });

  test('a cell with a note gets no comment, and Delete Note keeps threads', () => {
    const ctl = new EditorController();
    ctl.doc.ws.legacyComments.push(makeLegacyComment({ ref: 'A1', author: 'x', text: 'note' }));
    ctl.selectCell({ row: 1, col: 1 });
    newComment(ctl);
    expect(ctl.commentCard).toBeNull();
    expect(ctl.toast).toBe('cmtNoteHere');

    addComment(ctl, 2, 2, 'thread');
    ctl.selectRange({ r1: 1, c1: 1, r2: 2, c2: 2 });
    A.deleteNotes(ctl);
    expect(ctl.doc.ws.legacyComments).toEqual([]);
    expect(ctl.doc.ws.threadedComments).toHaveLength(1);
  });

  test('times are read as UTC', () => {
    expect(formatCommentTime('2024-05-01T09:30:00.00', 'en-US')).toBe(new Date(Date.UTC(2024, 4, 1, 9, 30)).toLocaleString('en-US', { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit' }));
    expect(formatCommentTime(undefined, 'en-US')).toBe('');
  });
});

describe('copying a sheet with threads', () => {
  // Excel drops a thread whose id (or mention id) also appears on another sheet.
  const ids = (ctl: EditorController) =>
    ctl.doc.wb.sheets.flatMap((s) => (s.kind === 'worksheet' ? (s.sheet.threadedComments ?? []).map((c) => c.id) : []));
  const shape = (ctl: EditorController, index: number) => {
    const ref = ctl.doc.wb.sheets[index];
    if (ref?.kind !== 'worksheet') throw new Error('expected a worksheet');
    return commentThreads(ref.sheet).map((t) => ({
      ref: t.root.ref,
      text: t.root.text,
      person: t.root.personId,
      replies: t.replies.map((r) => [r.text, r.parentId === t.root.id, r.personId]),
    }));
  };

  test('keeps ids unique and conversations intact through undo, redo, save, reopen and a second copy', async () => {
    const ctl = new EditorController();
    addComment(ctl, 2, 2, 'First');
    replyToComment(ctl, rootAt(ctl, 2, 2), 'Second');
    const root = rootAt(ctl, 2, 2);
    root.mentions = [{ personId: root.personId, mentionId: '{11111111-2222-3333-4444-555555555555}', startIndex: 0, length: 5 }];
    const original = shape(ctl, 0);
    const originalIds = ids(ctl);

    A.duplicateSheet(ctl, 0, 1);
    expect(ctl.doc.wb.sheets).toHaveLength(2);
    expect(new Set(ids(ctl)).size).toBe(4);
    expect(shape(ctl, 1)).toEqual(original);
    expect(shape(ctl, 0)).toEqual(original);
    expect(ids(ctl).slice(0, 2)).toEqual(originalIds);
    const mentionIds = ctl.doc.wb.sheets.flatMap((sh) =>
      sh.kind === 'worksheet' ? (sh.sheet.threadedComments ?? []).flatMap((c) => (c.mentions ?? []).map((m) => m.mentionId)) : [],
    );
    expect(new Set(mentionIds).size).toBe(2);

    A.undo(ctl);
    expect(ctl.doc.wb.sheets).toHaveLength(1);
    expect(ids(ctl)).toEqual(originalIds);
    A.redo(ctl);
    expect(new Set(ids(ctl)).size).toBe(4);

    const reopened = new EditorController();
    reopened.doc.replaceWorkbook(await loadWorkbook(fromArrayBuffer(await ctl.doc.toBytes())), 'copy.xlsx');
    expect(new Set(ids(reopened)).size).toBe(4);
    expect(shape(reopened, 1)).toEqual(original);

    A.duplicateSheet(reopened, 1, 2);
    expect(new Set(ids(reopened)).size).toBe(6);
    expect(shape(reopened, 2)).toEqual(original);
  });
});
