<script lang="ts">
  import { untrack } from 'svelte';
  import { duplicateSheet, moveSheetTo } from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();
  // DialogHost re-creates a dialog for every opening, so its arguments are read once.
  const args = untrack(() => props) ?? {};

  const ctl = getEditor();
  const index = typeof args['index'] === 'number' ? args['index'] : ctl.doc.activeSheetIndex;
  const sheets = ctl.doc.wb.sheets.map((s) => s.sheet.title);
  /** Position the sheet goes before; `sheets.length` is "(move to end)". */
  let before = $state(index);
  let copy = $state(false);

  function onok(): void {
    if (copy) duplicateSheet(ctl, index, before);
    // Removing the sheet first shifts every later position down by one.
    else moveSheetTo(ctl, index, before > index ? before - 1 : before);
  }
</script>

<Dialog title={t('dlgMoveCopy')} width={320} {onok}>
  <div class="lbl">{t('dlgMoveSelectedSheets')}</div>
  <div class="row">
    <label for="mc-book">{t('dlgToBook')}</label>
    <select id="mc-book" class="xl-select grow" disabled><option>{ctl.doc.fileName}</option></select>
  </div>
  <div class="lbl">{t('dlgBeforeSheet')}</div>
  <div class="list sheets" role="listbox" aria-label={t('dlgBeforeSheet')}>
    {#each sheets as title, i (i)}
      <button role="option" aria-selected={before === i} onclick={() => (before = i)}>{title}</button>
    {/each}
    <button role="option" aria-selected={before === sheets.length} onclick={() => (before = sheets.length)}>{t('dlgMoveToEnd')}</button>
  </div>
  <label class="check"><input type="checkbox" bind:checked={copy} />{t('dlgCreateCopy')}</label>
</Dialog>

<style>
  .sheets {
    height: 150px;
    margin-bottom: 6px;
  }
</style>
