<script lang="ts">
  import type { HorizontalAlignment, VerticalAlignment } from '@office-kit/xlsx/styles';
  import { t, type MessageKey } from '../../i18n/i18n.svelte.ts';
  import type { AlignState } from '../format-cells.ts';

  let { model = $bindable() }: { model: AlignState } = $props();

  const HORIZONTAL: ReadonlyArray<[HorizontalAlignment, MessageKey]> = [
    ['general', 'dlgAlignGeneral'],
    ['left', 'dlgAlignLeftIndent'],
    ['center', 'dlgAlignCenter'],
    ['right', 'dlgAlignRightIndent'],
    ['fill', 'dlgAlignFill'],
    ['justify', 'dlgAlignJustify'],
    ['centerContinuous', 'dlgAlignCenterAcross'],
    ['distributed', 'dlgAlignDistributedIndent'],
  ];
  const VERTICAL: ReadonlyArray<[VerticalAlignment, MessageKey]> = [
    ['top', 'dlgAlignTop'],
    ['center', 'dlgAlignCenter'],
    ['bottom', 'dlgAlignBottom'],
    ['justify', 'dlgAlignJustify'],
    ['distributed', 'dlgAlignDistributed'],
  ];
  const indentable = $derived(model.horizontal === 'left' || model.horizontal === 'right' || model.horizontal === 'distributed');

  function setDegrees(d: number): void {
    model.degrees = Math.max(-90, Math.min(90, Math.round(Number.isFinite(d) ? d : 0)));
    model.stacked = false;
  }
</script>

<div class="align">
  <div class="left">
    <fieldset>
      <legend>{t('dlgTextAlignment')}</legend>
      <div class="row">
        <label for="al-h">{t('dlgHorizontal')}</label>
        <select id="al-h" class="xl-select grow" bind:value={model.horizontal} onchange={() => { if (!indentable) model.indent = 0; }}>
          {#each HORIZONTAL as [value, label] (value)}<option {value}>{t(label)}</option>{/each}
        </select>
      </div>
      <div class="row">
        <label for="al-indent">{t('dlgIndent')}</label>
        <input id="al-indent" class="xl-input" type="number" min="0" max="250" disabled={!indentable} bind:value={model.indent} onchange={() => (model.indent = Math.max(0, Math.min(250, Math.trunc(model.indent || 0))))} />
      </div>
      <div class="row">
        <label for="al-v">{t('dlgVertical')}</label>
        <select id="al-v" class="xl-select grow" bind:value={model.vertical}>
          {#each VERTICAL as [value, label] (value)}<option {value}>{t(label)}</option>{/each}
        </select>
      </div>
    </fieldset>
    <fieldset>
      <legend>{t('dlgTextControl')}</legend>
      <label class="check"><input type="checkbox" bind:checked={model.wrap} onchange={() => { if (model.wrap) model.shrink = false; }} />{t('wrapText')}</label>
      <label class="check"><input type="checkbox" bind:checked={model.shrink} disabled={model.wrap} />{t('dlgShrinkToFit')}</label>
      <label class="check"><input type="checkbox" bind:checked={model.merge} />{t('mergeCells')}</label>
    </fieldset>
  </div>
  <fieldset class="orient">
    <legend>{t('orientation')}</legend>
    <button class="stack xl-btn outlined" aria-pressed={model.stacked} onclick={() => (model.stacked = !model.stacked)} aria-label={t('verticalText')}>
      <span>T<br />e<br />x<br />t</span>
    </button>
    <svg class="dial" viewBox="0 0 120 120" role="img" aria-label="{model.degrees}°">
      <path d="M60 10 A50 50 0 0 1 60 110" fill="none" stroke="#bbb" />
      {#each [-90, -45, 0, 45, 90] as d (d)}
        {@const rad = (d * Math.PI) / 180}
        <circle cx={60 + 50 * Math.cos(rad)} cy={60 - 50 * Math.sin(rad)} r="2.5" fill="#888" />
      {/each}
      {#if !model.stacked}
        {@const rad = (model.degrees * Math.PI) / 180}
        <line x1="60" y1="60" x2={60 + 46 * Math.cos(rad)} y2={60 - 46 * Math.sin(rad)} stroke="var(--xl-accent)" stroke-width="2" />
        <text x="68" y="63" font-size="10" transform="rotate({-model.degrees} 60 60)">{t('dlgText')}</text>
      {/if}
    </svg>
    <div class="row">
      <input id="al-deg" class="xl-input" type="number" min="-90" max="90" value={model.stacked ? 0 : model.degrees} oninput={(e) => setDegrees(e.currentTarget.valueAsNumber)} />
      <label for="al-deg">{t('dlgDegrees')}</label>
    </div>
  </fieldset>
</div>

<style>
  .align {
    display: flex;
    gap: 12px;
  }
  .left {
    flex: 1;
  }
  .orient {
    width: 150px;
    display: flex;
    flex-direction: column;
    align-items: center;
  }
  .stack {
    font-size: 10px;
    line-height: 1.05;
    padding: 4px 8px;
    margin-bottom: 4px;
  }
  .dial {
    width: 110px;
    height: 110px;
  }
</style>
