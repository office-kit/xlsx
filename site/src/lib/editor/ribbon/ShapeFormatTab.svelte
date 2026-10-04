<script lang="ts">
  // Shape Format, the contextual tab of a selected shape, text box or line:
  // fill and outline colours, outline weight, z-order, size and Delete.
  import { getEditor } from '../core/context.ts';
  import { reorderDrawing, setShapeFill, setShapeOutline, setShapeOutlineWeight, shapeAt } from '../core/shapes.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import ColorGrid from '../ui/ColorGrid.svelte';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import DrawingSizeGroup from './DrawingSizeGroup.svelte';
  import Group from './Group.svelte';

  let { index }: { index: number } = $props();

  const ctl = getEditor();
  const doc = ctl.doc;
  const WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

  const shape = $derived.by(() => {
    void doc.version;
    return shapeAt(doc, index);
  });
  const count = $derived.by(() => {
    void doc.version;
    return doc.ws.drawing?.items.length ?? 0;
  });

  function reorder(step: 1 | -1): void {
    ctl.selectedDrawing = reorderDrawing(doc, index, step);
  }

  function remove(): void {
    const ws = doc.ws;
    if (!ws.drawing?.items[index]) return;
    ctl.selectedDrawing = null;
    doc.transact('Delete Object', (tx) => {
      tx.sheet(ws, 'drawing');
      ws.drawing?.items.splice(index, 1);
    });
  }
</script>

<Group label={t('shpGroupStyles')}>
  <div class="stack">
    {#if !shape?.connector}
      <MenuButton icon="fill" label={t('shpFill')} title={t('shpFill')}>
        {#snippet menu(close)}
          <ColorGrid palette={doc.styles.palette} noneLabel={t('shpNoFill')} onpick={(rgb) => { close(); setShapeFill(doc, index, rgb === null ? null : `#${rgb}`); }} />
        {/snippet}
      </MenuButton>
    {/if}
    <MenuButton icon="borders" label={t('shpOutline')} title={t('shpOutline')}>
      {#snippet menu(close)}
        <ColorGrid palette={doc.styles.palette} noneLabel={t('shpNoOutline')} onpick={(rgb) => { close(); setShapeOutline(doc, index, rgb === null ? null : `#${rgb}`); }} />
        <div class="xl-menu-sep"></div>
        <div class="weights" role="group" aria-label={t('shpWeight')}>
          <span class="wl">{t('shpWeight')}</span>
          {#each WEIGHTS as w (w)}
            <button class="xl-menu-item weight" onclick={() => { close(); setShapeOutlineWeight(doc, index, w); }}>
              <span class="sample" style:border-top-width="{Math.max(1, (w * 4) / 3)}px"></span>{w} pt
            </button>
          {/each}
        </div>
      {/snippet}
    </MenuButton>
  </div>
</Group>

<Group label={t('shpGroupArrange')}>
  <div class="stack">
    <button class="xl-btn" disabled={index >= count - 1} onclick={() => reorder(1)}><Icon name="chevron-up" size={14} /><span>{t('shpBringForward')}</span></button>
    <button class="xl-btn" disabled={index <= 0} onclick={() => reorder(-1)}><Icon name="chevron-down" size={14} /><span>{t('shpSendBackward')}</span></button>
    <button class="xl-btn" onclick={remove}><Icon name="delete" size={14} /><span>{t('shpDelete')}</span></button>
  </div>
</Group>

<DrawingSizeGroup {index} />

<style>
  .stack {
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: flex-start;
    gap: 2px;
    height: 100%;
  }
  .weights {
    display: flex;
    flex-direction: column;
    min-width: 160px;
  }
  .wl {
    padding: 2px 10px;
    font-size: 11.5px;
    font-weight: 600;
    color: var(--xl-text-2);
  }
  .weight {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .sample {
    display: inline-block;
    width: 48px;
    border-top: solid #333;
  }
</style>
