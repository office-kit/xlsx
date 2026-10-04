<script lang="ts">
  // Insert ▸ Shapes: a gallery of preset shapes. Picking one arms the draw
  // tool; the next drag on the grid draws it (a click drops the default size).
  import { getEditor } from '../core/context.ts';
  import { SHAPE_GALLERY } from '../core/shapes.ts';
  import { LINE_PRESETS, presetPath } from '../grid/shape-render.ts';
  import { en } from '../i18n/en.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import Popup from '../ui/Popup.svelte';

  const ctl = getEditor();
  let open = $state(false);
  let root: HTMLButtonElement | undefined = $state();

  const isMessageKey = (key: string): key is MessageKey => key in en;
  const shapeName = (prst: string): string => {
    const key = `prst_${prst}`;
    return isMessageKey(key) ? t(key) : prst;
  };

  function thumb(prst: string): string {
    const line = LINE_PRESETS.has(prst);
    const marker = prst === 'straightConnector1' ? ' marker-end="url(#shp-thumb-arrow)"' : '';
    return (
      '<svg width="24" height="24" viewBox="-2 -2 24 24" overflow="visible">' +
      (marker ? '<defs><marker id="shp-thumb-arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto"><path d="M0,0 L10,5 L0,10 Z" fill="#444"/></marker></defs>' : '') +
      `<path d="${presetPath(prst, 20, prst.endsWith('Callout') ? 16 : 20)}" fill="${line ? 'none' : '#fff'}" stroke="#444" stroke-width="1.2"${marker}/></svg>`
    );
  }

  function pick(prst: string): void {
    open = false;
    if (ctl.edit && !ctl.commitEdit()) return;
    ctl.selectedDrawing = null;
    ctl.shapeTool = { kind: 'shape', prst };
  }
</script>

<button class="xl-btn big" bind:this={root} title={t('shpShapes')} aria-haspopup="menu" aria-expanded={open} onclick={() => (open = !open)} onmousedown={(e) => e.preventDefault()}>
  <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="9" r="6" fill="none" stroke="currentColor" stroke-width="1.4" /><rect x="10" y="10" width="11" height="11" fill="var(--xl-accent)" fill-opacity="0.25" stroke="currentColor" stroke-width="1.4" /></svg>
  <span class="lbl">{t('shpShapes')}<Icon name="chevron-down" size={9} /></span>
</button>
{#if open && root}
  <Popup anchor={root} onclose={() => (open = false)}>
    <div class="gallery">
      {#each SHAPE_GALLERY as g (g.group)}
        <div class="head">{t(g.group)}</div>
        <div class="row">
          {#each g.shapes as prst (prst)}
            <button class="thumb" title={shapeName(prst)} aria-label={shapeName(prst)} data-prst={prst} onclick={() => pick(prst)}>{@html thumb(prst)}</button>
          {/each}
        </div>
      {/each}
    </div>
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
  .lbl {
    display: inline-flex;
    align-items: center;
    gap: 1px;
    white-space: nowrap;
  }
  .gallery {
    width: 250px;
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
    width: 30px;
    height: 30px;
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
