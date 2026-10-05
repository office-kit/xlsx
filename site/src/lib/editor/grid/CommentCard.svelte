<script lang="ts">
  // One threaded-comment card, as Excel for Mac draws it: the conversation on a
  // cell, a reply box, and a "…" menu per comment. With no thread on the cell
  // it is the compose card of New Comment. Used floating beside the cell and
  // docked in the Comments pane.
  import type { ThreadedComment } from '@office-kit/xlsx/worksheet';
  import { cellAddress } from '../core/address.ts';
  import {
    addComment,
    commentUserName,
    deleteComment,
    editComment,
    formatCommentTime,
    personName,
    replyToComment,
    setThreadResolved,
    threadAt,
  } from '../core/comments.ts';
  import { getEditor } from '../core/context.ts';
  import { i18n, t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';

  interface Props {
    row: number;
    col: number;
    /** Docked cards sit in the pane's flow and show their cell address. */
    docked?: boolean;
    /** Hover previews are read-only until clicked. */
    preview?: boolean;
    onclose?: () => void;
  }
  let { row, col, docked = false, preview = false, onclose }: Props = $props();

  const ctl = getEditor();
  const doc = ctl.doc;

  const thread = $derived.by(() => {
    void doc.version;
    return threadAt(doc.ws, row, col);
  });
  const comments = $derived(thread ? [thread.root, ...thread.replies] : []);
  const resolved = $derived(thread?.root.done === true);

  let draft = $state('');
  let menuFor = $state<string | null>(null);
  let editingId = $state<string | null>(null);
  let editText = $state('');

  function author(c: ThreadedComment): string {
    void doc.version;
    return personName(doc.wb, c.personId);
  }

  function initials(name: string): string {
    const parts = name.trim().split(/\s+/);
    return parts.map((p) => [...p][0] ?? '').join('').slice(0, 2).toUpperCase();
  }

  function post(): void {
    const text = draft.trim();
    if (!text) return;
    if (thread) replyToComment(ctl, thread.root, text);
    else addComment(ctl, row, col, text);
    draft = '';
  }

  function startEdit(c: ThreadedComment): void {
    menuFor = null;
    editingId = c.id;
    editText = c.text;
  }

  function saveEdit(): void {
    const text = editText.trim();
    if (editingId && text) editComment(ctl, editingId, text);
    editingId = null;
  }

  function remove(c: ThreadedComment): void {
    menuFor = null;
    deleteComment(ctl, c);
  }

  function resolve(done: boolean): void {
    menuFor = null;
    if (thread) setThreadResolved(ctl, thread.root.id, done);
  }

  // ⌘↩ posts, as in Excel; Escape backs out one level.
  function textKey(ev: KeyboardEvent, submit: () => void, cancel: () => void): void {
    if (ev.key === 'Enter' && (ev.metaKey || ev.ctrlKey)) {
      ev.preventDefault();
      submit();
    } else if (ev.key === 'Escape') {
      ev.preventDefault();
      cancel();
    }
  }

  function autofocus(el: HTMLTextAreaElement): void {
    if (!preview && !docked) el.focus();
  }

  function editFocus(el: HTMLTextAreaElement): void {
    el.focus();
    el.select();
  }
</script>

<!-- The listeners only fence the card off from the grid underneath (keys typed
     here must not move the selection) and pin a hover preview; the controls
     inside carry their own semantics. -->
<!-- svelte-ignore a11y_no_noninteractive_element_interactions -->
<div
  class="card"
  class:docked
  class:resolved
  class:preview
  role="group"
  aria-label={t('cmtPaneTitle')}
  onpointerdown={(e) => e.stopPropagation()}
  onkeydown={(e) => e.stopPropagation()}
  onclick={() => {
    if (preview) ctl.commentCard = { row, col };
  }}
>
  {#if docked}
    <div class="where">
      <button class="cell-link" onclick={() => { ctl.selectCell({ row, col }); ctl.reveal(row, col); }}>{cellAddress(row, col)}</button>
      {#if resolved}<span class="chip">{t('cmtResolved')}</span>{/if}
    </div>
  {:else if resolved}
    <div class="where"><span class="chip">{t('cmtResolved')}</span></div>
  {/if}

  {#each comments as c, i (c.id)}
    {@const name = author(c)}
    <div class="comment" class:reply={i > 0}>
      <div class="head">
        <span class="avatar" aria-hidden="true">{initials(name)}</span>
        <div class="who">
          <strong>{name}</strong>
          <span class="time">{formatCommentTime(c.created, i18n.locale)}</span>
        </div>
        {#if !preview}
          <div class="more">
            <button class="xl-btn icon" title={t('cmtMoreActions')} aria-label={t('cmtMoreActions')} aria-expanded={menuFor === c.id} onclick={() => (menuFor = menuFor === c.id ? null : c.id)}>
              <Icon name="more" size={14} />
            </button>
            {#if menuFor === c.id}
              <div class="menu xl-menu" role="menu">
                <button class="xl-menu-item" role="menuitem" onclick={() => startEdit(c)}>{t('cmtEdit')}</button>
                {#if i === 0}
                  <button class="xl-menu-item" role="menuitem" onclick={() => remove(c)}>{t('cmtDeleteThread')}</button>
                  <button class="xl-menu-item" role="menuitem" onclick={() => resolve(!resolved)}>{t(resolved ? 'cmtReopen' : 'cmtResolve')}</button>
                {:else}
                  <button class="xl-menu-item" role="menuitem" onclick={() => remove(c)}>{t('cmtDeleteComment')}</button>
                {/if}
              </div>
            {/if}
          </div>
        {/if}
      </div>
      {#if editingId === c.id}
        <textarea class="xl-input text-in" rows="3" bind:value={editText} use:editFocus onkeydown={(e) => textKey(e, saveEdit, () => (editingId = null))}></textarea>
        <div class="actions">
          <button class="xl-btn outlined" onclick={() => (editingId = null)}>{t('cmtCancel')}</button>
          <button class="xl-btn primary" disabled={!editText.trim()} onclick={saveEdit}>{t('cmtSave')}</button>
        </div>
      {:else}
        <div class="body">{c.text}</div>
      {/if}
    </div>
  {/each}

  {#if !preview}
    {#if !thread}
      <div class="head">
        <span class="avatar" aria-hidden="true">{initials(commentUserName())}</span>
        <div class="who"><strong>{commentUserName()}</strong></div>
      </div>
    {/if}
    {#if resolved}
      <div class="actions">
        <button class="xl-btn outlined" onclick={() => resolve(false)}>{t('cmtReopen')}</button>
      </div>
    {:else}
      <textarea
        class="xl-input text-in"
        rows={thread ? 1 : 3}
        placeholder={t(thread ? 'cmtReplyPlaceholder' : 'cmtStartPlaceholder')}
        aria-label={t(thread ? 'cmtReply' : 'cmtNewComment')}
        bind:value={draft}
        use:autofocus
        onkeydown={(e) => textKey(e, post, () => (draft ? (draft = '') : onclose?.()))}
      ></textarea>
      {#if draft || !thread}
        <div class="actions">
          {#if !thread && onclose}<button class="xl-btn outlined" onclick={() => onclose?.()}>{t('cmtCancel')}</button>{/if}
          <button class="xl-btn primary" disabled={!draft.trim()} title="⌘↩" onclick={post}>{t('cmtPost')}</button>
        </div>
      {/if}
    {/if}
  {/if}
</div>

<style>
  .card {
    width: 280px;
    padding: 8px 10px;
    background: #fff;
    border: 1px solid var(--xl-border-strong);
    border-radius: 6px;
    box-shadow: var(--xl-shadow-lg);
    font-size: 12.5px;
    color: var(--xl-text);
    cursor: default;
  }
  .card.docked {
    width: auto;
    box-shadow: var(--xl-shadow);
  }
  .card.preview {
    cursor: pointer;
  }
  .card.resolved .comment {
    opacity: 0.7;
  }
  .where {
    display: flex;
    align-items: center;
    gap: 6px;
    margin-bottom: 4px;
  }
  .cell-link {
    border: 0;
    background: none;
    padding: 0;
    font: inherit;
    font-weight: 600;
    color: var(--xl-accent);
    cursor: pointer;
  }
  .chip {
    padding: 0 6px;
    border-radius: 8px;
    background: var(--xl-hover);
    color: var(--xl-text-2);
    font-size: 11px;
  }
  .comment + .comment {
    margin-top: 6px;
    padding-top: 6px;
    border-top: 1px solid var(--xl-border);
  }
  .head {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .avatar {
    flex: none;
    display: grid;
    place-items: center;
    width: 24px;
    height: 24px;
    border-radius: 50%;
    background: #7b61c4;
    color: #fff;
    font-size: 10px;
    font-weight: 600;
  }
  .who {
    flex: 1;
    min-width: 0;
    display: flex;
    flex-direction: column;
    line-height: 1.25;
  }
  .who strong {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .time {
    color: var(--xl-text-3);
    font-size: 11px;
  }
  .more {
    position: relative;
  }
  .menu {
    position: absolute;
    top: 100%;
    right: 0;
    min-width: 160px;
  }
  .body {
    margin: 4px 0 0 32px;
    white-space: pre-wrap;
    overflow-wrap: anywhere;
  }
  .text-in {
    display: block;
    box-sizing: border-box;
    width: 100%;
    margin-top: 8px;
    resize: vertical;
    font: inherit;
  }
  .actions {
    display: flex;
    justify-content: flex-end;
    gap: 6px;
    margin-top: 6px;
  }
</style>
