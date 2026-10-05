<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { goToReference } from '../core/names.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  const wb = ctl.doc.wb;
  // Built-in names (Print_Area, _FilterDatabase) and hidden names are not offered, as in Excel.
  const names = wb.definedNames
    .filter((d) => !d.hidden && !d.name.startsWith('_xlnm.') && (d.scope === undefined || d.scope === ctl.doc.activeSheetIndex))
    .map((d) => d.name)
    .sort((a, b) => a.localeCompare(b));
  let reference = $state('');
  let notice = $state<string | null>(null);

  function onok(): boolean {
    const text = reference.trim();
    if (!text) return false;
    if (!goToReference(ctl, text)) {
      notice = t('invalidReference');
      return false;
    }
    return true;
  }
</script>

<Dialog title={t('dlgGoTo')} width={340} {onok} bind:notice>
  {#snippet footerStart()}
    <button class="xl-btn outlined" onclick={() => ctl.openDialog('gotoSpecial')}>{t('dlgSpecialEllipsis')}</button>
  {/snippet}
  <div class="lbl">{t('dlgGoToList')}</div>
  <div class="list names" role="listbox" aria-label={t('dlgGoToList')}>
    {#each names as name (name)}
      <button role="option" aria-selected={reference === name} onclick={() => (reference = name)} ondblclick={() => { reference = name; if (onok()) ctl.closeDialog(); }}>{name}</button>
    {:else}
      <div class="hint empty">{t('noNames')}</div>
    {/each}
  </div>
  <div class="row">
    <label for="goto-ref">{t('dlgReference')}</label>
    <input id="goto-ref" class="xl-input grow" bind:value={reference} spellcheck="false" data-autofocus />
  </div>
</Dialog>

<style>
  .names {
    height: 150px;
    margin-bottom: 6px;
  }
  .empty {
    padding: 6px;
  }
</style>
