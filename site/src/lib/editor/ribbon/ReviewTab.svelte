<script lang="ts">
  import * as A from '../core/actions.ts';
  import { deleteCommentsInSelection, newComment, stepComment } from '../core/comments.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
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
  <button class="xl-btn big" onclick={() => ctl.openDialog('workbookStatistics')}><Icon name="calc" size={24} /><span>{t('workbookStatistics')}</span></button>
</Group>

<Group label={t('groupComments')}>
  <button class="xl-btn big" onclick={() => newComment(ctl)}><Icon name="note" size={24} /><span>{t('cmtNewComment')}</span></button>
  <div class="col">
    <button class="xl-btn" onclick={() => deleteCommentsInSelection(ctl)}>{t('cmtDelete')}</button>
    <button class="xl-btn" onclick={() => stepComment(ctl, -1)}>{t('cmtPrevious')}</button>
    <button class="xl-btn" onclick={() => stepComment(ctl, 1)}>{t('cmtNext')}</button>
  </div>
  <button class="xl-btn big" aria-pressed={ctl.commentsPane} onclick={() => (ctl.commentsPane = !ctl.commentsPane)}><Icon name="note" size={24} /><span>{t('cmtShowComments')}</span></button>
</Group>

<Group label={t('groupNotes')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('note')}><Icon name="note" size={24} /><span>{t('newNote')}</span></button>
  <div class="col">
    <button class="xl-btn" onclick={() => A.stepNote(ctl, -1)}>{t('previousNote')}</button>
    <button class="xl-btn" onclick={() => A.stepNote(ctl, 1)}>{t('nextNote')}</button>
    <button class="xl-btn" onclick={() => A.deleteNotes(ctl)}>{t('deleteNote')}</button>
  </div>
</Group>

<Group label={t('groupProtect')}>
  <button class="xl-btn big" aria-pressed={protectedBook} onclick={() => A.toggleWorkbookProtection(ctl)}><Icon name="lock" size={24} /><span>{t('protectWorkbook')}</span></button>
  <button class="xl-btn big" aria-pressed={protectedSheet} onclick={() => (protectedSheet ? A.unprotectSheet(ctl) : ctl.openDialog('protectSheet'))}><Icon name="lock" size={24} /><span>{protectedSheet ? t('unprotectSheet') : t('protectSheet')}</span></button>
</Group>

<style>
  .col {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
</style>
