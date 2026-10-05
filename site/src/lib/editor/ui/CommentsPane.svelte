<script lang="ts">
  // Review ▸ Show Comments: every thread on the sheet, docked on the right,
  // plus the name new comments are posted under.
  import { commentThreads, commentUserName, newComment, setCommentUserName } from '../core/comments.ts';
  import { getEditor } from '../core/context.ts';
  import CommentCard from '../grid/CommentCard.svelte';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from './Icon.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;

  const threads = $derived.by(() => {
    void doc.version;
    return commentThreads(doc.ws);
  });

  let userName = $state(commentUserName());
</script>

<aside class="pane" aria-label={t('cmtPaneTitle')}>
  <header>
    <span class="title">{t('cmtPaneTitle')}</span>
    <button class="xl-btn icon" title={t('cmtClosePane')} aria-label={t('cmtClosePane')} onclick={() => (ctl.commentsPane = false)}><Icon name="close" size={14} /></button>
  </header>
  <div class="tools">
    <button class="xl-btn" onclick={() => newComment(ctl)}><Icon name="plus" size={14} /><span>{t('cmtNew')}</span></button>
    <label class="user">
      <span>{t('cmtUserName')}</span>
      <input class="xl-input" bind:value={userName} onchange={() => setCommentUserName(userName)} />
    </label>
  </div>
  <div class="list">
    {#each threads as th (th.root.id)}
      <CommentCard row={th.row} col={th.col} docked />
    {:else}
      <p class="empty">{t('cmtNone')}</p>
    {/each}
  </div>
</aside>

<style>
  .pane {
    display: flex;
    flex-direction: column;
    width: 300px;
    flex: none;
    min-height: 0;
    background: var(--xl-panel);
    border-left: 1px solid var(--xl-border);
  }
  header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 6px 10px 2px;
  }
  .title {
    font-size: 15px;
    font-weight: 600;
  }
  .tools {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    padding: 4px 10px 8px;
    border-bottom: 1px solid var(--xl-border);
  }
  .user {
    display: flex;
    align-items: center;
    gap: 4px;
    font-size: 11.5px;
    color: var(--xl-text-2);
  }
  .user input {
    width: 100px;
  }
  .list {
    display: flex;
    flex-direction: column;
    gap: 8px;
    flex: 1;
    min-height: 0;
    overflow: auto;
    padding: 8px 10px;
    background: var(--xl-bg);
  }
  .empty {
    color: var(--xl-text-2);
    font-size: 12px;
  }
</style>
