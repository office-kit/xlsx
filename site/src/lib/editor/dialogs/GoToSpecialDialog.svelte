<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import type { GoToSpecial } from '../core/find.ts';
  import { goToSpecialRanges, selectRanges } from '../core/goto-special.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const OPTIONS: ReadonlyArray<[GoToSpecial, MessageKey]> = [
    ['comments', 'dlgGsNotes'],
    ['constants', 'dlgGsConstants'],
    ['formulas', 'dlgGsFormulas'],
    ['blanks', 'dlgGsBlanks'],
    ['currentRegion', 'dlgGsCurrentRegion'],
    ['lastCell', 'dlgGsLastCell'],
    ['visible', 'dlgGsVisible'],
    ['conditionalFormats', 'dlgGsConditionalFormats'],
    ['dataValidation', 'dlgGsDataValidation'],
  ];

  const ctl = getEditor();
  let kind = $state<GoToSpecial>('constants');
  const kinds = $state({ numbers: true, text: true, logicals: true, errors: true });
  let notice = $state<string | null>(null);
  const typed = $derived(kind === 'constants' || kind === 'formulas');

  function onok(): boolean {
    if (!selectRanges(ctl, goToSpecialRanges(ctl, kind, kinds))) {
      notice = t('noCellsFound');
      return false;
    }
    return true;
  }
</script>

<Dialog title={t('dlgGoToSpecial')} width={360} {onok} bind:notice>
  <div class="opts" role="radiogroup" aria-label={t('dlgGoToSpecial')}>
    {#each OPTIONS as [value, label] (value)}
      <label class="check"><input type="radio" name="gs" {value} bind:group={kind} />{t(label)}</label>
      {#if value === 'formulas'}
        <div class="sub">
          <label class="check"><input type="checkbox" bind:checked={kinds.numbers} disabled={!typed} />{t('nfNumber')}</label>
          <label class="check"><input type="checkbox" bind:checked={kinds.text} disabled={!typed} />{t('nfText')}</label>
          <label class="check"><input type="checkbox" bind:checked={kinds.logicals} disabled={!typed} />{t('dlgGsLogicals')}</label>
          <label class="check"><input type="checkbox" bind:checked={kinds.errors} disabled={!typed} />{t('dlgGsErrors')}</label>
        </div>
      {/if}
    {/each}
  </div>
</Dialog>

<style>
  .sub {
    margin-left: 22px;
    display: grid;
    grid-template-columns: 1fr 1fr;
  }
</style>
