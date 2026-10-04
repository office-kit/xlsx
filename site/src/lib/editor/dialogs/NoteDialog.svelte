<script lang="ts">
  import { makeLegacyComment } from '@office-kit/xlsx/worksheet';
  import { cellAddress } from '../core/address.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  const ws = doc.ws;
  const { row, col } = doc.selection.active;
  const ref = cellAddress(row, col);
  const isHere = (r: string) => r.replaceAll('$', '').split(':')[0] === ref;
  const existing = ws.legacyComments.find((c) => isHere(c.ref));
  let author = $state(existing?.author ?? doc.wb.properties?.lastModifiedBy ?? doc.wb.properties?.creator ?? t('dlgDefaultAuthor'));
  let text = $state(existing?.text ?? '');

  function save(remove: boolean): void {
    doc.transact(remove ? 'Delete Note' : existing ? 'Edit Note' : 'New Note', (tx) => {
      tx.sheet(ws, 'legacyComments');
      const others = ws.legacyComments.filter((c) => !isHere(c.ref));
      ws.legacyComments = remove ? others : [...others, makeLegacyComment({ ref, author: author.trim(), text })];
    });
  }

  function onok(): void {
    // An emptied note is removed, as Excel drops a note whose text was cleared.
    if (text.trim() === '') {
      if (existing) save(true);
      return;
    }
    save(false);
  }
</script>

<Dialog title={t(existing ? 'editNote' : 'newNote')} width={360} {onok}>
  {#snippet footerStart()}
    {#if existing}<button class="xl-btn outlined" onclick={() => { save(true); ctl.closeDialog(); }}>{t('deleteNote')}</button>{/if}
  {/snippet}
  <div class="row">
    <label for="note-author">{t('dlgAuthor')}</label>
    <input id="note-author" class="xl-input grow" bind:value={author} />
  </div>
  <label class="lbl" for="note-text">{t('dlgNoteFor', { cell: ref })}</label>
  <textarea id="note-text" class="xl-input area" bind:value={text} data-autofocus></textarea>
  <p class="hint">{t('dlgNoteHint')}</p>
</Dialog>

<style>
  .area {
    width: 100%;
    height: 120px;
    resize: vertical;
    background: #fffbd6;
  }
  .lbl {
    display: block;
    margin: 6px 0 3px;
  }
</style>
