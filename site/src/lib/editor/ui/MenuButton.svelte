<script lang="ts">
  // A ribbon/menu button that opens a Popup. `split` renders the Excel split
  // button: the main half runs `onmain`, the arrow half opens the menu.
  import type { Snippet } from 'svelte';
  import Popup from './Popup.svelte';
  import Icon from './Icon.svelte';

  let {
    label,
    icon,
    title,
    split = false,
    large = false,
    onmain,
    menu,
    align = 'left',
    disabled = false,
  }: {
    label?: string;
    icon?: string;
    title?: string;
    split?: boolean;
    large?: boolean;
    onmain?: () => void;
    menu: Snippet<[() => void]>;
    align?: 'left' | 'right';
    disabled?: boolean;
  } = $props();

  let open = $state(false);
  let root: HTMLElement | undefined = $state();
  const close = () => (open = false);
</script>

<span class="mb" class:large bind:this={root}>
  {#if split}
    <button class="xl-btn main" class:large {title} {disabled} onclick={() => onmain?.()} onmousedown={(e) => e.preventDefault()}>
      {#if icon}<Icon name={icon} size={large ? 26 : 18} />{/if}
      {#if label}<span class="lbl">{label}</span>{/if}
    </button>
    <button class="xl-btn arrow" title={title} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={() => (open = !open)} onmousedown={(e) => e.preventDefault()}>
      <Icon name="chevron-down" size={11} />
    </button>
  {:else}
    <button class="xl-btn" class:large {title} aria-haspopup="menu" aria-expanded={open} {disabled} onclick={() => (open = !open)} onmousedown={(e) => e.preventDefault()}>
      {#if icon}<Icon name={icon} size={large ? 26 : 18} />{/if}
      {#if label}<span class="lbl">{label}</span>{/if}
      <Icon name="chevron-down" size={11} />
    </button>
  {/if}
</span>
{#if open && root}
  <Popup anchor={root} onclose={close} {align}>
    {@render menu(close)}
  </Popup>
{/if}

<style>
  .mb {
    display: inline-flex;
    align-items: stretch;
  }
  .mb .main {
    border-top-right-radius: 0;
    border-bottom-right-radius: 0;
    padding-right: 3px;
  }
  .mb .arrow {
    border-top-left-radius: 0;
    border-bottom-left-radius: 0;
    padding: 2px 2px;
  }
  .xl-btn.large {
    flex-direction: column;
    min-width: 44px;
    height: 64px;
    font-size: 11px;
    gap: 2px;
  }
  .mb.large {
    flex-direction: column;
  }
  .mb.large .main {
    border-radius: 4px 4px 0 0;
    height: 44px;
    padding: 2px 4px;
  }
  .mb.large .arrow {
    border-radius: 0 0 4px 4px;
    height: 20px;
  }
  .lbl {
    line-height: 1.1;
  }
</style>
