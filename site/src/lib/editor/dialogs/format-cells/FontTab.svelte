<script lang="ts">
  import type { UnderlineStyle } from '@office-kit/xlsx/styles';
  import { cssFontFamily } from '../../core/render-style.ts';
  import { t, type MessageKey } from '../../i18n/i18n.svelte.ts';
  import ColorPicker from '../ColorPicker.svelte';
  import { FONT_NAMES, FONT_SIZES, type FontState, type FontStyle } from '../format-cells.ts';

  let { model = $bindable() }: { model: FontState } = $props();

  const STYLES: ReadonlyArray<[FontStyle, MessageKey]> = [
    ['regular', 'dlgFontRegular'],
    ['italic', 'italic'],
    ['bold', 'bold'],
    ['boldItalic', 'dlgFontBoldItalic'],
  ];
  const UNDERLINES: ReadonlyArray<[UnderlineStyle, MessageKey]> = [
    ['none', 'dlgNone'],
    ['single', 'dlgUnderlineSingle'],
    ['double', 'dlgUnderlineDouble'],
    ['singleAccounting', 'dlgUnderlineSingleAccounting'],
    ['doubleAccounting', 'dlgUnderlineDoubleAccounting'],
  ];
  const names = $derived(FONT_NAMES.includes(model.name) ? FONT_NAMES : [model.name, ...FONT_NAMES]);
  let sizeText = $state(String(model.size));

  function commitSize(): void {
    const n = Number(sizeText);
    // Excel accepts 1–409 in half-point steps.
    if (Number.isFinite(n) && n >= 1 && n <= 409) model.size = Math.round(n * 2) / 2;
    sizeText = String(model.size);
  }
</script>

<div class="font">
  <div class="cols">
    <div class="col name">
      <label class="lbl" for="ft-name">{t('fontName')}</label>
      <input id="ft-name" class="xl-input" bind:value={model.name} spellcheck="false" />
      <div class="list" role="listbox" aria-label={t('fontName')}>
        {#each names as n (n)}
          <button role="option" aria-selected={model.name === n} style:font-family={cssFontFamily(n)} onclick={() => (model.name = n)}>{n}</button>
        {/each}
      </div>
    </div>
    <div class="col">
      <div class="lbl">{t('dlgFontStyle')}</div>
      <div class="list" role="listbox" aria-label={t('dlgFontStyle')}>
        {#each STYLES as [value, label] (value)}
          <button role="option" aria-selected={model.style === value} onclick={() => (model.style = value)}>{t(label)}</button>
        {/each}
      </div>
    </div>
    <div class="col size">
      <label class="lbl" for="ft-size">{t('size')}</label>
      <input id="ft-size" class="xl-input" bind:value={sizeText} onchange={commitSize} inputmode="decimal" />
      <div class="list" role="listbox" aria-label={t('size')}>
        {#each FONT_SIZES as s (s)}
          <button role="option" aria-selected={model.size === s} onclick={() => { model.size = s; sizeText = String(s); }}>{s}</button>
        {/each}
      </div>
    </div>
  </div>
  <div class="row">
    <label for="ft-ul">{t('underline')}</label>
    <select id="ft-ul" class="xl-select" bind:value={model.underline}>
      {#each UNDERLINES as [value, label] (value)}<option {value}>{t(label)}</option>{/each}
    </select>
    <span class="lbl right">{t('dlgColor')}</span>
    <ColorPicker bind:value={model.color} noneLabel={t('automatic')} label={t('fontColor')} />
  </div>
  <div class="cols">
    <fieldset class="effects">
      <legend>{t('dlgEffects')}</legend>
      <label class="check"><input type="checkbox" bind:checked={model.strike} />{t('strikethrough')}</label>
      <label class="check"><input type="checkbox" bind:checked={model.superscript} onchange={() => { if (model.superscript) model.subscript = false; }} />{t('dlgSuperscript')}</label>
      <label class="check"><input type="checkbox" bind:checked={model.subscript} onchange={() => { if (model.subscript) model.superscript = false; }} />{t('dlgSubscript')}</label>
    </fieldset>
    <fieldset class="preview">
      <legend>{t('dlgPreview')}</legend>
      <div
        class="preview-text"
        style:font-family={cssFontFamily(model.name)}
        style:font-size="{Math.min(model.size, 28) * (4 / 3)}px"
        style:font-weight={model.style === 'bold' || model.style === 'boldItalic' ? 700 : 400}
        style:font-style={model.style === 'italic' || model.style === 'boldItalic' ? 'italic' : 'normal'}
        style:color={model.color ? `#${model.color}` : '#000'}
        style:text-decoration={[model.underline !== 'none' ? 'underline' : '', model.strike ? 'line-through' : ''].join(' ').trim() || 'none'}
        style:text-decoration-style={model.underline.startsWith('double') ? 'double' : 'solid'}
      >
        {#if model.superscript}<sup>AaBbCcYyZz</sup>{:else if model.subscript}<sub>AaBbCcYyZz</sub>{:else}AaBbCcYyZz{/if}
      </div>
    </fieldset>
  </div>
</div>

<style>
  .cols {
    display: flex;
    gap: 10px;
  }
  .col {
    display: flex;
    flex-direction: column;
    gap: 3px;
    width: 110px;
  }
  .col.name {
    flex: 1;
  }
  .col.size {
    width: 64px;
  }
  .col .list {
    height: 120px;
  }
  .lbl.right {
    min-width: 0;
    margin-left: 12px;
  }
  .effects {
    width: 170px;
  }
  .preview {
    flex: 1;
  }
  .preview-text {
    height: 64px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: #fff;
    border: 1px solid var(--xl-border);
    overflow: hidden;
    white-space: nowrap;
  }
</style>
