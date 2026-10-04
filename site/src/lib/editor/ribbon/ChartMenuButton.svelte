<script lang="ts">
  // A large ribbon button whose icon is a thumbnail of its first chart type
  // and whose menu is that button's chart gallery.
  import type { ChartChoice, ChartGroup } from '../core/charts.ts';
  import { CHART_GROUPS } from '../core/charts.ts';
  import { getEditor } from '../core/context.ts';
  import Icon from '../ui/Icon.svelte';
  import Popup from '../ui/Popup.svelte';
  import ChartGallery from './ChartGallery.svelte';
  import { chartThumbnail } from './chart-ui.ts';

  let { label, title, groups, onpick, onmore }: { label: string; title: string; groups: readonly ChartGroup[]; onpick: (choice: ChartChoice) => void; onmore: (choice: ChartChoice) => void } = $props();

  const palette = getEditor().doc.styles.palette;
  const first = $derived(CHART_GROUPS[groups[0] ?? 'column'][0]);
  let open = $state(false);
  let root: HTMLButtonElement | undefined = $state();
</script>

<button class="xl-btn big" bind:this={root} {title} aria-haspopup="menu" aria-expanded={open} onclick={() => (open = !open)} onmousedown={(e) => e.preventDefault()}>
  <span class="icon">{@html chartThumbnail(first, 30, 24, palette)}</span>
  <span class="lbl">{label}<Icon name="chevron-down" size={9} /></span>
</button>
{#if open && root}
  <Popup anchor={root} onclose={() => (open = false)}>
    <ChartGallery
        {groups}
        onpick={(c) => {
          open = false;
          onpick(c);
        }}
        onmore={() => {
          open = false;
          onmore(first);
        }}
    />
  </Popup>
{/if}

<style>
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
    gap: 2px;
  }
  .icon :global(svg) {
    display: block;
  }
  .lbl {
    display: inline-flex;
    align-items: center;
    gap: 1px;
    white-space: nowrap;
  }
</style>
