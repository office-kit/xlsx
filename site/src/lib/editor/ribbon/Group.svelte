<script lang="ts">
  import type { Snippet } from 'svelte';
  let { label, children }: { label: string; children: Snippet } = $props();
</script>

<div class="group" role="group" aria-label={label}>
  <div class="body">{@render children()}</div>
  <div class="label">{label}</div>
</div>

<style>
  .group {
    display: flex;
    flex-direction: column;
    padding: 2px 6px 0;
    border-right: 1px solid var(--xl-border);
    flex-shrink: 0;
  }
  .body {
    flex: 1;
    display: flex;
    align-items: center;
    gap: 4px;
  }
  .label {
    text-align: center;
    color: var(--xl-text-2);
    font-size: 10.5px;
    padding: 1px 0 2px;
    white-space: nowrap;
  }
</style>
