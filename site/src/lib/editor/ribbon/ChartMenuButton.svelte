<script lang="ts">
  // A small icon-only ribbon button, as in Excel's Charts group: an icon plus
  // an arrow, opening that button's chart gallery.
  import type { ChartChoice, ChartGroup } from '../core/charts.ts';
  import { CHART_GROUPS } from '../core/charts.ts';
  import Icon from '../ui/Icon.svelte';
  import Popup from '../ui/Popup.svelte';
  import ChartGallery from './ChartGallery.svelte';

  let { label, title, icon, groups, onpick, onmore }: { label: string; title: string; icon: string; groups: readonly ChartGroup[]; onpick: (choice: ChartChoice) => void; onmore: (choice: ChartChoice) => void } = $props();

  const first = $derived(CHART_GROUPS[groups[0] ?? 'column'][0]);
  let open = $state(false);
  let root: HTMLButtonElement | undefined = $state();
</script>

<button class="xl-btn" bind:this={root} {title} aria-label={label} aria-haspopup="menu" aria-expanded={open} onclick={() => (open = !open)} onmousedown={(e) => e.preventDefault()}>
  <Icon name={icon} size={18} />
  <Icon name="chevron-down" size={9} />
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
  .xl-btn {
    height: 22px;
    min-height: 22px;
    gap: 1px;
    padding: 1px 3px;
  }
</style>
