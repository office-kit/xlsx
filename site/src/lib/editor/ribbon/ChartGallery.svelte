<script lang="ts">
  // The drop-down gallery under an Insert ▸ chart button: one titled row of
  // thumbnails per chart group, then "More Charts…".
  import { CHART_GROUPS, type ChartChoice, type ChartGroup } from '../core/charts.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { CHART_LABELS, chartThumbnail, GROUP_LABELS } from './chart-ui.ts';

  let { groups, onpick, onmore }: { groups: readonly ChartGroup[]; onpick: (choice: ChartChoice) => void; onmore: () => void } = $props();

  const palette = getEditor().doc.styles.palette;
</script>

<div class="gallery">
  {#each groups as g (g)}
    <div class="head">{t(GROUP_LABELS[g])}</div>
    <div class="row">
      {#each CHART_GROUPS[g] as c (c)}
        <button class="thumb" title={t(CHART_LABELS[c])} aria-label={t(CHART_LABELS[c])} onclick={() => onpick(c)}>{@html chartThumbnail(c, 44, 36, palette)}</button>
      {/each}
    </div>
  {/each}
  <div class="xl-menu-sep"></div>
  <button class="xl-menu-item" onclick={onmore}>{t('chMoreCharts')}</button>
</div>

<style>
  .gallery {
    width: 268px;
  }
  .head {
    padding: 4px 10px 2px;
    font-size: 11.5px;
    font-weight: 600;
    color: var(--xl-text-2);
  }
  .row {
    display: flex;
    flex-wrap: wrap;
    gap: 2px;
    padding: 0 8px 4px;
  }
  .thumb {
    width: 50px;
    height: 42px;
    padding: 2px;
    border: 1px solid transparent;
    border-radius: 3px;
    background: transparent;
    cursor: pointer;
  }
  .thumb:hover,
  .thumb:focus-visible {
    border-color: var(--xl-accent);
    background: var(--xl-hover);
    outline: none;
  }
  .thumb :global(svg) {
    display: block;
    pointer-events: none;
  }
</style>
