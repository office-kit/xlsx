<script lang="ts">
  import type { CellValue } from '@office-kit/xlsx/cell';
  import { CUSTOM_CODES, FORMAT_CELLS_CATEGORIES, type FormatCategory } from '../../core/number-formats.ts';
  import { t, type MessageKey } from '../../i18n/i18n.svelte.ts';
  import { CURRENCY_SYMBOLS, listFor, numberCode, type FormatSampler, type NumberState } from '../format-cells.ts';

  let { model = $bindable(), sampler, sample }: { model: NumberState; sampler: FormatSampler; sample: CellValue } = $props();

  const LABELS: Record<FormatCategory, MessageKey> = {
    general: 'nfGeneral',
    number: 'nfNumber',
    currency: 'nfCurrency',
    accounting: 'nfAccounting',
    date: 'dlgNfDate',
    time: 'nfTime',
    percentage: 'nfPercentage',
    fraction: 'nfFraction',
    scientific: 'nfScientific',
    text: 'nfText',
    special: 'dlgNfSpecial',
    custom: 'nfCustom',
  };
  const DESCRIPTIONS: Record<FormatCategory, MessageKey> = {
    general: 'dlgNfGeneralHelp',
    number: 'dlgNfNumberHelp',
    currency: 'dlgNfCurrencyHelp',
    accounting: 'dlgNfAccountingHelp',
    date: 'dlgNfDateHelp',
    time: 'dlgNfTimeHelp',
    percentage: 'dlgNfPercentageHelp',
    fraction: 'dlgNfFractionHelp',
    scientific: 'dlgNfScientificHelp',
    text: 'dlgNfTextHelp',
    special: 'dlgNfSpecialHelp',
    custom: 'dlgNfCustomHelp',
  };
  const NEGATIVE_SAMPLE = -1234.1;

  const code = $derived(numberCode(model));
  const sampleText = $derived(sample === null ? '' : (sampler.format(code, sample) ?? ''));
  const hasDecimals = $derived(['number', 'currency', 'accounting', 'percentage', 'scientific'].includes(model.category));
  const hasSymbol = $derived(model.category === 'currency' || model.category === 'accounting');
  const hasNegative = $derived(model.category === 'number' || model.category === 'currency');
  const typeList = $derived(listFor(model.category));
  const customList = $derived(CUSTOM_CODES.includes(model.customCode) || model.customCode === '' ? CUSTOM_CODES : [model.customCode, ...CUSTOM_CODES]);

  function pick(category: FormatCategory): void {
    const current = code;
    model.category = category;
    const list = listFor(category);
    if (list.length > 0 && !list.includes(model.listCode)) model.listCode = list[0] ?? '';
    // Custom starts from whatever the other categories had built, as in Excel.
    if (category === 'custom') model.customCode = current;
  }

  function negativeCode(negative: 0 | 1 | 2 | 3): string {
    return numberCode({ ...model, negative });
  }
</script>

<div class="number">
  <div class="cats">
    <div class="lbl">{t('dlgCategory')}</div>
    <div class="list" role="listbox" aria-label={t('dlgCategory')}>
      {#each FORMAT_CELLS_CATEGORIES as cat (cat)}
        <button role="option" aria-selected={model.category === cat} onclick={() => pick(cat)}>{t(LABELS[cat])}</button>
      {/each}
    </div>
  </div>
  <div class="opts">
    <fieldset class="sample">
      <legend>{t('dlgSample')}</legend>
      <div class="sample-text">{sampleText || ' '}</div>
    </fieldset>
    {#if hasDecimals}
      <div class="row">
        <label for="nf-dec">{t('dlgDecimalPlaces')}</label>
        <input id="nf-dec" class="xl-input" type="number" min="0" max="30" bind:value={model.decimals} onchange={() => (model.decimals = Math.max(0, Math.min(30, Math.trunc(model.decimals || 0))))} />
      </div>
    {/if}
    {#if hasSymbol}
      <div class="row">
        <label for="nf-sym">{t('dlgSymbol')}</label>
        <select id="nf-sym" class="xl-select" bind:value={model.symbol}>
          {#each CURRENCY_SYMBOLS as sym (sym)}
            <option value={sym}>{sym === '' ? t('dlgNone') : sym}</option>
          {/each}
        </select>
      </div>
    {/if}
    {#if model.category === 'number'}
      <label class="check"><input type="checkbox" bind:checked={model.thousands} />{t('dlgUseThousands')}</label>
    {/if}
    {#if hasNegative}
      <div class="lbl">{t('dlgNegativeNumbers')}</div>
      <div class="list short" role="listbox" aria-label={t('dlgNegativeNumbers')}>
        {#each [0, 1, 2, 3] as const as n (n)}
          {@const c = negativeCode(n)}
          <button role="option" aria-selected={model.negative === n} class:red={n === 1 || n === 3} onclick={() => (model.negative = n)}>{sampler.format(c, NEGATIVE_SAMPLE)}</button>
        {/each}
      </div>
    {/if}
    {#if typeList.length > 0}
      <div class="lbl">{t('dlgType')}</div>
      <div class="list tall" role="listbox" aria-label={t('dlgType')}>
        {#each typeList as c (c)}
          <button role="option" aria-selected={model.listCode === c} onclick={() => (model.listCode = c)}>{sampler.format(c, typeof sample === 'number' ? sample : 45292.5) ?? c}</button>
        {/each}
      </div>
    {/if}
    {#if model.category === 'custom'}
      <div class="lbl"><label for="nf-custom">{t('dlgType')}</label></div>
      <input id="nf-custom" class="xl-input code" bind:value={model.customCode} spellcheck="false" />
      <div class="list tall" role="listbox" aria-label={t('dlgType')}>
        {#each customList as c (c)}
          <button role="option" class="code" aria-selected={model.customCode === c} onclick={() => (model.customCode = c)}>{c}</button>
        {/each}
      </div>
    {/if}
    <p class="hint">{t(DESCRIPTIONS[model.category])}</p>
  </div>
</div>

<style>
  .number {
    display: flex;
    gap: 14px;
  }
  .cats {
    width: 130px;
    flex: none;
  }
  .cats .list {
    height: 300px;
  }
  .opts {
    flex: 1;
    min-width: 0;
  }
  .sample {
    margin-top: 0;
  }
  .sample-text {
    min-height: 18px;
    font-family: var(--xl-mono);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .list.short {
    height: 82px;
  }
  .list.tall {
    height: 140px;
  }
  .red {
    color: #c00000;
  }
  .code {
    font-family: var(--xl-mono);
    font-size: 11.5px;
    width: 100%;
    margin-bottom: 4px;
  }
  .lbl {
    margin: 6px 0 3px;
  }
</style>
