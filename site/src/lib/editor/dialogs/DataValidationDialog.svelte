<script lang="ts">
  import type { DataValidationErrorStyle, DataValidationOperator } from '@office-kit/xlsx/worksheet';
  import { getEditor } from '../core/context.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';
  import { activeRule, applyValidation, blankForm, formFor, usesOperator, usesSecondValue, type Allow } from './data-validation.ts';

  const ALLOW: ReadonlyArray<[Allow, MessageKey]> = [
    ['any', 'dlgDvAny'],
    ['whole', 'dlgDvWhole'],
    ['decimal', 'dlgDvDecimal'],
    ['list', 'dlgDvList'],
    ['date', 'dlgDvDate'],
    ['time', 'dlgDvTime'],
    ['textLength', 'dlgDvTextLength'],
    ['custom', 'dlgDvCustom'],
  ];
  const OPERATORS: ReadonlyArray<[DataValidationOperator, MessageKey]> = [
    ['between', 'dlgOpBetween'],
    ['notBetween', 'dlgOpNotBetween'],
    ['equal', 'dlgOpEqual'],
    ['notEqual', 'dlgOpNotEqual'],
    ['greaterThan', 'dlgOpGreaterThan'],
    ['lessThan', 'dlgOpLessThan'],
    ['greaterThanOrEqual', 'dlgOpGreaterOrEqual'],
    ['lessThanOrEqual', 'dlgOpLessOrEqual'],
  ];
  const STYLES: ReadonlyArray<[DataValidationErrorStyle, MessageKey]> = [
    ['stop', 'dlgDvStop'],
    ['warning', 'dlgDvWarning'],
    ['information', 'dlgDvInformation'],
  ];

  const ctl = getEditor();
  const existing = activeRule(ctl);
  let form = $state(formFor(ctl, existing));
  let tab = $state<'settings' | 'input' | 'error'>('settings');
  let notice = $state<string | null>(null);

  const ranged = $derived(usesSecondValue(form));
  const label1 = $derived<MessageKey>(
    form.allow === 'list' ? 'dlgDvSource' : form.allow === 'custom' ? 'dlgDvFormula' : ranged ? 'dlgDvMinimum' : form.operator === 'greaterThan' || form.operator === 'greaterThanOrEqual' ? 'dlgDvMinimum' : form.operator === 'lessThan' || form.operator === 'lessThanOrEqual' ? 'dlgDvMaximum' : 'dlgDvValue',
  );

  function onok(): boolean {
    const err = applyValidation(ctl, form, existing);
    if (err) {
      tab = 'settings';
      notice = t(err);
      return false;
    }
    return true;
  }

  function clearAll(): void {
    form = blankForm();
  }
</script>

<Dialog title={t('dataValidation')} width={460} {onok} bind:notice>
  {#snippet footerStart()}
    <button class="xl-btn outlined" onclick={clearAll}>{t('clearAll')}</button>
  {/snippet}
  <div class="seg" role="tablist">
    <button role="tab" aria-selected={tab === 'settings'} onclick={() => (tab = 'settings')}>{t('dlgDvSettings')}</button>
    <button role="tab" aria-selected={tab === 'input'} onclick={() => (tab = 'input')}>{t('dlgDvInputMessage')}</button>
    <button role="tab" aria-selected={tab === 'error'} onclick={() => (tab = 'error')}>{t('dlgDvErrorAlert')}</button>
  </div>
  <div class="panel" role="tabpanel">
    {#if tab === 'settings'}
      <div class="lbl">{t('dlgDvCriteria')}</div>
      <div class="row">
        <label for="dv-allow">{t('dlgDvAllow')}</label>
        <select id="dv-allow" class="xl-select grow" bind:value={form.allow}>
          {#each ALLOW as [value, label] (value)}<option {value}>{t(label)}</option>{/each}
        </select>
      </div>
      {#if form.allow !== 'any'}
        <div class="row">
          <label for="dv-op">{t('dlgDvData')}</label>
          <select id="dv-op" class="xl-select grow" bind:value={form.operator} disabled={!usesOperator(form.allow)}>
            {#each OPERATORS as [value, label] (value)}<option {value}>{t(label)}</option>{/each}
          </select>
        </div>
        <div class="row">
          <label for="dv-v1">{t(label1)}</label>
          <input id="dv-v1" class="xl-input grow" bind:value={form.value1} spellcheck="false" />
        </div>
        {#if ranged}
          <div class="row">
            <label for="dv-v2">{t('dlgDvMaximum')}</label>
            <input id="dv-v2" class="xl-input grow" bind:value={form.value2} spellcheck="false" />
          </div>
        {/if}
        {#if form.allow === 'list'}
          <p class="hint">{t('dlgDvListHint')}</p>
        {/if}
      {/if}
      <label class="check"><input type="checkbox" bind:checked={form.ignoreBlank} disabled={form.allow === 'any'} />{t('dlgDvIgnoreBlank')}</label>
      {#if form.allow === 'list'}
        <label class="check"><input type="checkbox" bind:checked={form.inCellDropdown} />{t('dlgDvInCellDropdown')}</label>
      {/if}
      {#if existing}
        <label class="check"><input type="checkbox" bind:checked={form.applyToAll} />{t('dlgDvApplyToAll')}</label>
      {/if}
      {#if form.allow === 'any'}
        <p class="hint">{t('dlgDvAnyHint')}</p>
      {/if}
    {:else if tab === 'input'}
      <label class="check"><input type="checkbox" bind:checked={form.showInput} />{t('dlgDvShowInput')}</label>
      <div class="lbl">{t('dlgDvWhenSelected')}</div>
      <div class="row">
        <label for="dv-it">{t('dlgTitle')}</label>
        <input id="dv-it" class="xl-input grow" bind:value={form.inputTitle} maxlength="32" />
      </div>
      <label class="lbl" for="dv-im">{t('dlgDvMessage')}</label>
      <textarea id="dv-im" class="xl-input area" bind:value={form.inputMessage} maxlength="255"></textarea>
    {:else}
      <label class="check"><input type="checkbox" bind:checked={form.showError} />{t('dlgDvShowError')}</label>
      <div class="lbl">{t('dlgDvWhenInvalid')}</div>
      <div class="row">
        <label for="dv-es">{t('dlgStyle')}</label>
        <select id="dv-es" class="xl-select" bind:value={form.errorStyle}>
          {#each STYLES as [value, label] (value)}<option {value}>{t(label)}</option>{/each}
        </select>
      </div>
      <div class="row">
        <label for="dv-et">{t('dlgTitle')}</label>
        <input id="dv-et" class="xl-input grow" bind:value={form.errorTitle} maxlength="32" />
      </div>
      <label class="lbl" for="dv-em">{t('dlgDvErrorMessage')}</label>
      <textarea id="dv-em" class="xl-input area" bind:value={form.errorMessage} maxlength="255"></textarea>
    {/if}
  </div>
</Dialog>

<style>
  .panel {
    min-height: 230px;
  }
  .area {
    width: 100%;
    height: 90px;
    resize: vertical;
  }
  .lbl {
    display: block;
    margin: 6px 0 3px;
  }
</style>
