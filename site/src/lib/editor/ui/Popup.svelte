<script lang="ts">
  // A floating panel anchored to an element (ribbon drop-downs, colour
  // pickers). Closes on outside pointer-down and Escape; focus returns to the
  // grid through the editor's own focus effect.
  import { onMount, type Snippet } from 'svelte';

  let { anchor, onclose, children, align = 'left' }: { anchor: HTMLElement; onclose: () => void; children: Snippet; align?: 'left' | 'right' } = $props();
  let panel: HTMLDivElement;
  let pos = $state({ left: 0, top: 0 });

  onMount(() => {
    const r = anchor.getBoundingClientRect();
    const pr = panel.getBoundingClientRect();
    let left = align === 'right' ? r.right - pr.width : r.left;
    left = Math.max(4, Math.min(left, window.innerWidth - pr.width - 4));
    let top = r.bottom + 2;
    if (top + pr.height > window.innerHeight - 4) top = Math.max(4, r.top - pr.height - 2);
    pos = { left, top };
    const down = (e: PointerEvent) => {
      if (!panel.contains(e.target as Node) && !anchor.contains(e.target as Node)) onclose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        onclose();
      }
    };
    window.addEventListener('pointerdown', down, true);
    window.addEventListener('keydown', key, true);
    return () => {
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('keydown', key, true);
    };
  });
</script>

<div class="xl-menu popup" bind:this={panel} style:left="{pos.left}px" style:top="{pos.top}px" role="menu" tabindex="-1">
  {@render children()}
</div>

<style>
  .popup {
    max-height: calc(100vh - 16px);
    overflow: auto;
  }
</style>
