<script lang="ts">
  // The Size group of the Format / Picture Format tabs: height and width of
  // the selected drawing item, in centimetres for Japanese and inches for
  // English as Excel shows them. `lockable` adds Excel's "Lock aspect ratio"
  // box (on by default, as for pictures): changing one side scales the other.
  import { untrack } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { drawingRect, resizeDrawing } from '../core/drawing-anchor.ts';
  import { i18n, t } from '../i18n/i18n.svelte.ts';
  import Group from './Group.svelte';

  let { index, lockable = false }: { index: number; lockable?: boolean } = $props();

  let lockAspect = $state(untrack(() => lockable));

  const doc = getEditor().doc;
  const PX_PER_INCH = 96;
  const unit = $derived(i18n.locale === 'ja' ? 'cm' : 'in');
  const pxPerUnit = $derived(unit === 'cm' ? PX_PER_INCH / 2.54 : PX_PER_INCH);

  const rect = $derived.by(() => {
    void doc.version;
    void doc.layoutVersion;
    return drawingRect(doc, index);
  });

  const show = (px: number | undefined): string => (px === undefined ? '' : `${(px / pxPerUnit).toFixed(2)} ${unit}`);

  function parse(text: string): number | undefined {
    const n = Number.parseFloat(text.replace(',', '.'));
    return Number.isFinite(n) && n > 0 ? n * pxPerUnit : undefined;
  }

  function commit(side: 'h' | 'w', ev: Event): void {
    const input = ev.currentTarget;
    const r = rect;
    if (!(input instanceof HTMLInputElement) || !r) return;
    const px = parse(input.value);
    if (px === undefined) {
      input.value = show(side === 'h' ? r.h : r.w);
      return;
    }
    const k = side === 'h' ? px / r.h : px / r.w;
    const w = side === 'w' ? px : lockAspect ? r.w * k : r.w;
    const h = side === 'h' ? px : lockAspect ? r.h * k : r.h;
    resizeDrawing(doc, index, w, h);
  }
</script>

<Group label={t('chGroupSize')}>
  <div class="sizes">
    <label class="row"><span>{t('chHeight')}</span><input class="xl-input" value={show(rect?.h)} onchange={(e) => commit('h', e)} onkeydown={(e) => e.stopPropagation()} /></label>
    <label class="row"><span>{t('chWidth')}</span><input class="xl-input" value={show(rect?.w)} onchange={(e) => commit('w', e)} onkeydown={(e) => e.stopPropagation()} /></label>
    {#if lockable}
      <label class="row"><input type="checkbox" bind:checked={lockAspect} />{t('chLockAspect')}</label>
    {/if}
  </div>
</Group>

<style>
  .sizes {
    display: flex;
    flex-direction: column;
    justify-content: center;
    gap: 4px;
    height: 100%;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 11.5px;
  }
  .row span {
    min-width: 28px;
  }
  .xl-input {
    width: 76px;
  }
</style>
