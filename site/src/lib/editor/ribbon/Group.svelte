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
    padding: 2px 4px 0;
    margin: 0 0 3px;
    border-right: 1px solid var(--xl-border);
    flex-shrink: 0;
  }
  .body {
    height: 70px;
    display: flex;
    align-items: flex-start;
    gap: 2px;
  }
  .label {
    height: 18px;
    line-height: 16px;
    text-align: center;
    color: var(--xl-text-2);
    font-size: 11px;
    white-space: nowrap;
  }
</style>
