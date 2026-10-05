<script lang="ts">
  import * as A from '../core/actions.ts';
  import { deleteCommentsInSelection, newComment, stepComment } from '../core/comments.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';

  const ctl = getEditor();
  const protectedSheet = $derived.by(() => {
    void ctl.doc.version;
    return ctl.doc.ws.sheetProtection?.sheet === true;
  });
  const protectedBook = $derived.by(() => {
    void ctl.doc.version;
    return ctl.doc.wb.workbookProtection?.lockStructure === true;
  });
</script>

<Group label={t('groupProofing')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('workbookStatistics')}><Icon name="calc" size={30} /><span>{t('workbookStatistics')}</span></button>
</Group>

<Group label={t('groupComments')}>
  <button class="xl-btn big" onclick={() => newComment(ctl)}><Icon name="comment-new" size={30} /><span>{t('cmtNewComment')}</span></button>
  <button class="xl-btn big" onclick={() => deleteCommentsInSelection(ctl)}><Icon name="comment-delete" size={30} /><span>{t('cmtDelete')}</span></button>
  <button class="xl-btn big" onclick={() => stepComment(ctl, -1)}><Icon name="comment-prev" size={30} /><span>{t('cmtPrevious')}</span></button>
  <button class="xl-btn big" onclick={() => stepComment(ctl, 1)}><Icon name="comment-next" size={30} /><span>{t('cmtNext')}</span></button>
  <button class="xl-btn big" aria-pressed={ctl.commentsPane} onclick={() => (ctl.commentsPane = !ctl.commentsPane)}><Icon name="comment" size={30} /><span>{t('cmtShowComments')}</span></button>
</Group>

<Group label={t('groupNotes')}>
  <MenuButton large icon="note" label={t('groupNotes')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { ctl.openDialog('note'); close(); }}>{t('newNote')}</button>
      <button class="xl-menu-item" onclick={() => { A.stepNote(ctl, -1); close(); }}>{t('previousNote')}</button>
      <button class="xl-menu-item" onclick={() => { A.stepNote(ctl, 1); close(); }}>{t('nextNote')}</button>
      <button class="xl-menu-item" onclick={() => { A.deleteNotes(ctl); close(); }}>{t('deleteNote')}</button>
    {/snippet}
  </MenuButton>
</Group>

<Group label={t('groupProtect')}>
  <button class="xl-btn big" aria-pressed={protectedSheet} onclick={() => (protectedSheet ? A.unprotectSheet(ctl) : ctl.openDialog('protectSheet'))}><Icon name="protect-sheet" size={30} /><span>{protectedSheet ? t('unprotectSheet') : t('protectSheet')}</span></button>
  <button class="xl-btn big" aria-pressed={protectedBook} onclick={() => A.toggleWorkbookProtection(ctl)}><Icon name="protect-workbook" size={30} /><span>{t('protectWorkbook')}</span></button>
</Group>
