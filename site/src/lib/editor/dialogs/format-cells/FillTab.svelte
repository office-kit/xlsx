<script lang="ts">
  import { getEditor } from '../../core/context.ts';
  import { applyTint } from '../../core/theme.ts';
  import type { PatternType } from '@office-kit/xlsx/styles';
  import { t, type MessageKey } from '../../i18n/i18n.svelte.ts';
  import ColorPicker from '../ColorPicker.svelte';
  import { PATTERNS, type FillState } from '../format-cells.ts';

  let { model = $bindable() }: { model: FillState } = $props();

  const ctl = getEditor();
  const PATTERN_LABELS: Record<PatternType, MessageKey> = {
    none: 'dlgNone',
    solid: 'dlgPatternSolid',
    darkGray: 'dlgPatGray75',
    mediumGray: 'dlgPatGray50',
    lightGray: 'dlgPatGray25',
    gray125: 'dlgPatGray125',
    gray0625: 'dlgPatGray0625',
    darkHorizontal: 'dlgPatHStripe',
    darkVertical: 'dlgPatVStripe',
    darkDown: 'dlgPatReverseDiagStripe',
    darkUp: 'dlgPatDiagStripe',
    darkGrid: 'dlgPatDiagCrosshatch',
    darkTrellis: 'dlgPatThickDiagCrosshatch',
    lightHorizontal: 'dlgPatThinHStripe',
    lightVertical: 'dlgPatThinVStripe',
    lightDown: 'dlgPatThinReverseDiagStripe',
    lightUp: 'dlgPatThinDiagStripe',
    lightGrid: 'dlgPatThinHCrosshatch',
    lightTrellis: 'dlgPatThinDiagCrosshatch',
  };
  const STANDARD = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];
  const TINTS = [0, 0.8, 0.6, 0.4, -0.25, -0.5];
  const themeRows = $derived(TINTS.map((tint) => ctl.doc.styles.palette.slice(0, 10).map((hex) => applyTint(hex, tint).toUpperCase())));
</script>

<div class="fill">
  <div class="bg">
    <div class="lbl">{t('dlgBackgroundColor')}</div>
    <button class="xl-btn outlined nocolor" aria-pressed={model.background === null} onclick={() => (model.background = null)}>{t('noColor')}</button>
    <div class="swatches" role="listbox" aria-label={t('dlgBackgroundColor')}>
      {#each themeRows as row, r (r)}
        {#each row as hex, c (c)}
          <button role="option" class="sw" class:gap={r === 0} aria-selected={model.background === hex} aria-label="#{hex}" title="#{hex}" style:background="#{hex}" onclick={() => (model.background = hex)}></button>
        {/each}
      {/each}
    </div>
    <div class="swatches std">
      {#each STANDARD as hex (hex)}
        <button class="sw" aria-pressed={model.background === hex} aria-label="#{hex}" title="#{hex}" style:background="#{hex}" onclick={() => (model.background = hex)}></button>
      {/each}
    </div>
  </div>
  <div class="pattern">
    <div class="row">
      <span class="lbl">{t('dlgPatternColor')}</span>
    </div>
    <ColorPicker bind:value={model.patternColor} noneLabel={t('automatic')} label={t('dlgPatternColor')} />
    <div class="row">
      <label for="fl-pattern">{t('dlgPatternStyle')}</label>
    </div>
    <select id="fl-pattern" class="xl-select" bind:value={model.pattern}>
      {#each PATTERNS as p (p)}
        <option value={p}>{t(PATTERN_LABELS[p])}</option>
      {/each}
    </select>
    <fieldset>
      <legend>{t('dlgSample')}</legend>
      <div
        class="sample"
        style:background-color={model.background ? `#${model.background}` : '#fff'}
        style:background-image={model.pattern === 'solid' ? 'none' : `repeating-linear-gradient(45deg, #${model.patternColor ?? '000000'} 0 1px, transparent 1px 4px)`}
      ></div>
    </fieldset>
  </div>
</div>

<style>
  .fill {
    display: flex;
    gap: 18px;
  }
  .nocolor {
    width: 100%;
    margin-bottom: 6px;
  }
  .swatches {
    display: grid;
    grid-template-columns: repeat(10, 18px);
    gap: 2px 4px;
  }
  .swatches.std {
    margin-top: 8px;
  }
  .sw {
    width: 18px;
    height: 16px;
    border: 1px solid rgba(0, 0, 0, 0.15);
    padding: 0;
  }
  .sw.gap {
    margin-bottom: 4px;
  }
  .sw[aria-selected='true'],
  .sw[aria-pressed='true'] {
    outline: 2px solid #f5a623;
  }
  .pattern {
    flex: 1;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .sample {
    height: 60px;
    border: 1px solid var(--xl-border);
  }
</style>
