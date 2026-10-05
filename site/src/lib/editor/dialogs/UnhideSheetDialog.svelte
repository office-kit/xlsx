<script lang="ts">
  import { setSheetState } from '@office-kit/xlsx/workbook';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  // Very hidden sheets are only reachable from VBA in Excel, so they are not listed.
  const hidden = ctl.doc.wb.sheets.map((s, index) => ({ index, title: s.sheet.title, state: s.state })).filter((s) => s.state === 'hidden');
  let selected = $state<number[]>(hidden[0] ? [hidden[0].index] : []);

  function toggle(index: number, additive: boolean): void {
    if (!additive) selected = [index];
    else selected = selected.includes(index) ? selected.filter((i) => i !== index) : [...selected, index];
  }

  function onok(): void {
    const doc = ctl.doc;
    const first = Math.min(...selected);
    doc.transact('Unhide Sheet', (tx) => {
      tx.workbook('sheets');
      for (const s of hidden) if (selected.includes(s.index)) setSheetState(doc.wb, s.title, 'visible');
    });
    doc.activateSheet(first);
  }
</script>

<Dialog title={t('dlgUnhide')} width={300} {onok} okDisabled={selected.length === 0}>
  <div class="lbl">{t('dlgUnhideSheets')}</div>
  <div class="list sheets" role="listbox" aria-multiselectable="true" aria-label={t('dlgUnhideSheets')}>
    {#each hidden as s (s.index)}
      <button role="option" aria-selected={selected.includes(s.index)} onclick={(e) => toggle(s.index, e.metaKey || e.ctrlKey || e.shiftKey)}>{s.title}</button>
    {:else}
      <div class="hint empty">{t('dlgNoHiddenSheets')}</div>
    {/each}
  </div>
</Dialog>

<style>
  .sheets {
    height: 150px;
  }
  .empty {
    padding: 8px;
  }
</style>
