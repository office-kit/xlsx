<script lang="ts">
  import { untrack } from 'svelte';
  import { parseRangeAddress, rangeAddress } from '../core/address.ts';
  import { getEditor } from '../core/context.ts';
  import { resizeTable, tableRange } from '../core/tables.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';
  import RefInput from './RefInput.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();
  // DialogHost re-creates a dialog for every opening, so its arguments are read once.
  const args = untrack(() => props) ?? {};

  const ctl = getEditor();
  const def = typeof args['table'] === 'number' ? ctl.doc.ws.tables[args['table']] : undefined;
  const range = def ? tableRange(def) : undefined;
  let refText = $state(range ? `=${rangeAddress(range, true)}` : '');
  let notice = $state<string | null>(null);

  function onok(): boolean {
    if (!def) return true;
    const parsed = parseRangeAddress(refText.trim().replace(/^=/, ''));
    if (!parsed || (parsed.sheet !== undefined && parsed.sheet.toLowerCase() !== ctl.doc.ws.title.toLowerCase())) {
      notice = t('invalidReference');
      return false;
    }
    const err = resizeTable(ctl, def, parsed.range);
    if (err) {
      notice = t(err);
      return false;
    }
    return true;
  }
</script>

<Dialog title={t('dtResizeTable')} width={360} modal={false} {onok} bind:notice>
  <label class="lbl" for="rt-range">{t('dtResizePrompt')}</label>
  <RefInput id="rt-range" bind:value={refText} autofocus />
  <p class="hint">{t('dtResizeNote')}</p>
</Dialog>

<style>
  .lbl {
    display: block;
    margin-bottom: 4px;
  }
</style>
