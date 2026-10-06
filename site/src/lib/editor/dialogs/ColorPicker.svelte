<script lang="ts">
  // A colour well for dialogs: shows the current colour and opens the ribbon's
  // colour gallery. `value` is RRGGBB, or null for Automatic / No Color.
  import { getEditor } from '../core/context.ts';
  import ColorGrid from '../ui/ColorGrid.svelte';
  import Popup from '../ui/Popup.svelte';

  let { value = $bindable(null), noneLabel, label, onchange }: { value?: string | null; noneLabel: string; label: string; onchange?: (rgb: string | null) => void } = $props();

  const ctl = getEditor();
  let open = $state(false);
  let anchor = $state<HTMLButtonElement>();
</script>

<button class="xl-btn outlined well" bind:this={anchor} aria-label={label} aria-haspopup="menu" aria-expanded={open} onclick={() => (open = !open)}>
  <span class="swatch" class:none={value === null} style:background={value === null ? undefined : `#${value}`}></span>
  <span class="name">{value === null ? noneLabel : `#${value}`}</span>
  <span class="caret" aria-hidden="true">▾</span>
</button>
{#if open && anchor}
  <Popup {anchor} onclose={() => (open = false)}>
    <ColorGrid
      palette={ctl.doc.styles.palette}
      {noneLabel}
      onpick={(rgb) => {
        value = rgb;
        onchange?.(rgb);
        open = false;
      }}
    />
  </Popup>
{/if}

<style>
  .well {
    justify-content: flex-start;
    min-width: 140px;
  }
  .swatch {
    width: 28px;
    height: 14px;
    border: 1px solid rgba(0, 0, 0, 0.3);
  }
  .swatch.none {
    background: linear-gradient(to top right, #fff calc(50% - 1px), #c92a2a 50%, #fff calc(50% + 1px));
  }
  .name {
    flex: 1;
    text-align: left;
    color: var(--xl-text-2);
    font-family: var(--xl-mono);
    font-size: 11px;
  }
  .caret {
    color: var(--xl-text-3);
  }
</style>
