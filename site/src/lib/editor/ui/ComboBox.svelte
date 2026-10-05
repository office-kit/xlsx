<script lang="ts" module>
  /** A pickable entry, or a section heading such as Excel's "Theme Fonts". */
  export type ComboItem = { readonly value: string; readonly label?: string; readonly font?: string } | { readonly heading: string };
</script>

<script lang="ts">
  // Excel's editable combo box (Font, Font Size): type a value and press
  // Enter, or open the full list with the arrow and pick from it.
  import { tick } from 'svelte';
  import Icon from './Icon.svelte';
  import Popup from './Popup.svelte';

  let {
    value = $bindable(),
    items,
    label,
    width,
    oncommit,
  }: {
    value: string;
    items: readonly ComboItem[];
    label: string;
    width: number;
    /** Runs with the typed (Enter / change) or picked value. */
    oncommit: (value: string, how: 'typed' | 'picked') => void;
  } = $props();

  const uid = $props.id();
  let open = $state(false);
  let root: HTMLElement | undefined = $state();
  let list: HTMLDivElement | undefined = $state();
  // Enter commits and then blurs, and the blur fires `change`; commit once.
  let committedByEnter = false;

  async function toggle(): Promise<void> {
    open = !open;
    if (!open) return;
    await tick();
    // Open on the current value, as Excel scrolls its list to the selected font.
    list?.querySelector<HTMLElement>('[aria-selected="true"]')?.scrollIntoView({ block: 'center' });
  }

  function pick(v: string): void {
    open = false;
    value = v;
    oncommit(v, 'picked');
  }
</script>

<span class="combo" bind:this={root} style:width="{width}px">
  <input
    class="xl-input"
    bind:value
    aria-label={label}
    role="combobox"
    aria-controls="{uid}-list"
    aria-expanded={open}
    spellcheck="false"
    onkeydown={(e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        committedByEnter = true;
        oncommit(value, 'typed');
        (e.currentTarget as HTMLInputElement).blur();
        committedByEnter = false;
      } else if (e.key === 'ArrowDown' && e.altKey) {
        e.preventDefault();
        void toggle();
      }
    }}
    onchange={() => {
      if (!committedByEnter) oncommit(value, 'typed');
    }}
  />
  <button class="arrow" tabindex="-1" aria-label={label} aria-haspopup="listbox" aria-expanded={open} onclick={() => void toggle()} onmousedown={(e) => e.preventDefault()}>
    <Icon name="chevron-down" size={11} />
  </button>
</span>
{#if open && root}
  <Popup anchor={root} onclose={() => (open = false)}>
    <div class="list" id="{uid}-list" bind:this={list} role="listbox" aria-label={label}>
      {#each items as item, i ('heading' in item ? `h${i}` : `v${item.value}${i}`)}
        {#if 'heading' in item}
          <div class="head">{item.heading}</div>
        {:else}
          <button class="xl-menu-item" role="option" aria-selected={item.value === value} style:font-family={item.font} onclick={() => pick(item.value)} onmousedown={(e) => e.preventDefault()}>{item.label ?? item.value}</button>
        {/if}
      {/each}
    </div>
  </Popup>
{/if}

<style>
  .combo {
    display: inline-flex;
    position: relative;
  }
  .combo input {
    width: 100%;
    padding-right: 20px;
  }
  .arrow {
    position: absolute;
    right: 1px;
    top: 1px;
    bottom: 1px;
    width: 18px;
    border: 0;
    border-left: 1px solid transparent;
    background: transparent;
    color: var(--xl-text-2);
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
  }
  .arrow:hover {
    background: var(--xl-hover);
    border-left-color: var(--xl-border);
  }
  .list {
    max-height: 420px;
    overflow: auto;
    min-width: 100%;
  }
  .list [aria-selected='true'] {
    background: var(--xl-pressed);
  }
  .head {
    font-weight: 600;
    font-size: 11px;
    color: var(--xl-text-2);
    padding: 6px 10px 2px;
    background: #f7f7f7;
  }
</style>
