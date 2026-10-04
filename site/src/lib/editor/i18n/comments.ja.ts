// Japanese strings for threaded comments, using the terms of Excel for Mac's
// Japanese UI.

import type { commentsEn } from './comments.en.ts';

export const commentsJa: Partial<Record<keyof typeof commentsEn, string>> = {
  cmtNewComment: '新しいコメント',
  cmtDelete: '削除',
  cmtPrevious: '前のコメント',
  cmtNext: '次のコメント',
  cmtShowComments: 'コメントの表示',
  cmtReply: '返信',
  cmtReplyToComment: 'コメントに返信',
  cmtDeleteComment: 'コメントの削除',
  cmtReplyPlaceholder: '返信...',
  cmtStartPlaceholder: '会話を開始',
  cmtPost: '投稿',
  cmtCancel: 'キャンセル',
  cmtSave: '保存',
  cmtEdit: 'コメントの編集',
  cmtDeleteThread: 'スレッドの削除',
  cmtResolve: '解決',
  cmtReopen: 'スレッドを再開',
  cmtResolved: '解決済み',
  cmtMoreActions: 'その他のスレッド操作',
  cmtPaneTitle: 'コメント',
  cmtNew: '新規',
  cmtNone: 'このシートにはコメントがありません。',
  cmtClosePane: 'コメントを閉じる',
  cmtUserName: 'ユーザー名',
  cmtNoteHere: 'このセルにはメモがあります。コメントを追加するにはメモを削除してください。',
};
