<script lang="ts">
  // Picture Format ▸ Alt Text: the selected picture's description
  // (`<xdr:cNvPr descr>`), read by screen readers. One undo step.
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  const index = ctl.selectedDrawing;
  const item = index === null ? undefined : doc.ws.drawing?.items[index];
  const initial = item?.content.kind === 'picture' ? (item.content.picture.descr ?? '') : '';
  let text = $state(initial);

  function ok(): boolean {
    if (index === null || text === initial) return true;
    const ws = doc.ws;
    doc.transact('Alt Text', (tx) => {
      tx.sheet(ws, 'drawing');
      const live = ws.drawing?.items[index];
      if (live?.content.kind !== 'picture') return;
      if (text.trim()) live.content.picture.descr = text;
      else delete live.content.picture.descr;
    });
    return true;
  }
</script>

<Dialog title={t('chAltText')} width={380} onok={ok}>
  <p class="hint">{t('chAltTextHint')}</p>
  <!-- svelte-ignore a11y_autofocus -->
  <textarea class="xl-input" rows="5" bind:value={text} aria-label={t('chAltText')} autofocus></textarea>
</Dialog>

<style>
  textarea {
    width: 100%;
    box-sizing: border-box;
    resize: vertical;
    font: inherit;
  }
</style>
