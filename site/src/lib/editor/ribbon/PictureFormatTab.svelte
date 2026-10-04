<script lang="ts">
  // Picture Format, the contextual tab of a selected picture: Alt Text, Reset
  // Size (back to the image's own pixel size) and Size with the aspect ratio
  // locked, as Excel does for pictures by default.
  import { getEditor } from '../core/context.ts';
  import { resizeDrawing } from '../core/drawing-anchor.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import DrawingSizeGroup from './DrawingSizeGroup.svelte';
  import Group from './Group.svelte';

  let { index }: { index: number } = $props();

  const ctl = getEditor();
  const doc = ctl.doc;

  const natural = $derived.by(() => {
    void doc.version;
    const item = doc.ws.drawing?.items[index];
    const img = item?.content.kind === 'picture' ? item.content.picture.image : undefined;
    return img && img.width > 0 && img.height > 0 ? { w: img.width, h: img.height } : undefined;
  });
</script>

<Group label={t('chGroupAdjust')}>
  <button class="xl-btn big" disabled={!natural} onclick={() => natural && resizeDrawing(doc, index, natural.w, natural.h)}><Icon name="image" size={24} /><span>{t('chResetSize')}</span></button>
</Group>

<Group label={t('chGroupAccessibility')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('altText')}><Icon name="note" size={24} /><span>{t('chAltText')}</span></button>
</Group>

<DrawingSizeGroup {index} lockable />

<style>
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
</style>
