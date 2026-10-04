<script lang="ts">
  // The drop-down arrow beside the active cell of a table's total row, with
  // Excel's list of SUBTOTAL functions for that column.
  import { getEditor } from '../core/context.ts';
  import { setTotalFunction, TOTAL_FUNCTIONS, totalCellAt } from '../core/tables.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const ctl = getEditor();
  const doc = ctl.doc;
  let open = $state(false);

  const target = $derived.by(() => {
    void doc.version;
    if (ctl.edit) return undefined;
    const { row, col } = doc.selection.active;
    const hit = totalCellAt(doc.ws, row, col);
    if (!hit) return undefined;
    const geo = ctl.geometry;
    const r = geo.rectOf({ r1: row, c1: col, r2: row, c2: col });
    if (r.w <= 0 || r.x + r.w <= geo.headerW || r.y + r.h <= geo.headerH || r.x >= geo.width || r.y >= geo.height) return undefined;
    return { ...hit, col, x: r.x + r.w + 1, y: r.y + r.h - 18 };
  });

  $effect(() => {
    void doc.selection;
    open = false;
  });

  function choose(fn: (typeof TOTAL_FUNCTIONS)[number]['fn'] | 'none'): void {
    const tg = target;
    open = false;
    if (tg) setTotalFunction(ctl, tg.def, tg.col, fn);
  }
</script>

{#if target}
  <button
    class="arrow"
    style:left="{target.x}px"
    style:top="{target.y}px"
    aria-label={t('dtTotalRow')}
    aria-expanded={open}
    onpointerdown={(e) => e.stopPropagation()}
    onclick={() => (open = !open)}
  >
    <svg viewBox="0 0 10 10" width="10" height="10" aria-hidden="true"><path d="M2 3.5h6L5 7z" fill="currentColor" /></svg>
  </button>
  {#if open}
    <div class="xl-menu menu" role="menu" style:left="{target.x - 120}px" style:top="{target.y + 19}px">
      <button class="xl-menu-item" role="menuitemradio" aria-checked={!target.column.totalsRowFunction} onpointerdown={(e) => e.stopPropagation()} onclick={() => choose('none')}>{t('dtTotNone')}</button>
      {#each TOTAL_FUNCTIONS as f (f.fn)}
        <button class="xl-menu-item" role="menuitemradio" aria-checked={target.column.totalsRowFunction === f.fn} onpointerdown={(e) => e.stopPropagation()} onclick={() => choose(f.fn)}>
          {t(f.label)}
        </button>
      {/each}
      <div class="xl-menu-sep"></div>
      <button
        class="xl-menu-item"
        onpointerdown={(e) => e.stopPropagation()}
        onclick={() => {
          open = false;
          ctl.openDialog('insertFunction');
        }}>{t('dtTotMore')}</button
      >
    </div>
  {/if}
{/if}

<style>
  .arrow {
    position: absolute;
    width: 16px;
    height: 18px;
    padding: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    background: #fff;
    border: 1px solid var(--xl-border-strong, #b3b3b3);
    color: #444;
    cursor: pointer;
    z-index: 3;
  }
  .menu {
    position: absolute;
    min-width: 136px;
    z-index: 20;
  }
  .xl-menu-item[aria-checked='true'] {
    font-weight: 600;
  }
</style>
