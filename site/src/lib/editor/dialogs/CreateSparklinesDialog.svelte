<script lang="ts">
  // Insert ▸ Sparklines ▸ Line / Column / Win/Loss. As in Excel, a selected
  // block becomes the Data Range and the Location Range is left to the user.
  import type { SparklineType } from '@office-kit/xlsx/worksheet';
  import { untrack } from 'svelte';
  import { rangeAddress } from '../core/address.ts';
  import { getEditor } from '../core/context.ts';
  import { currentRange } from '../core/selection.ts';
  import { createSparklines } from '../core/sparklines.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();

  const ctl = getEditor();
  const doc = ctl.doc;
  const raw = untrack(() => props?.['type']);
  const type: SparklineType = raw === 'column' || raw === 'stacked' ? raw : 'line';
  const sel = currentRange(doc.selection);
  let dataText = $state(sel.r1 === sel.r2 && sel.c1 === sel.c2 ? '' : rangeAddress(sel));
  let locationText = $state('');
  let notice = $state<string | null>(null);

  function onok(): boolean {
    const err = createSparklines(doc, type, dataText.trim().replace(/^=/, ''), locationText.trim().replace(/^=/, ''));
    if (err) {
      notice = t(err);
      return false;
    }
    return true;
  }
</script>

<Dialog title={t('spkCreateTitle')} width={360} {onok} bind:notice>
  <p class="head">{t('spkChooseData')}</p>
  <label class="row">
    <span>{t('spkDataRange')}</span>
    <input class="xl-input range" bind:value={dataText} spellcheck="false" data-autofocus />
  </label>
  <p class="head">{t('spkChooseLocation')}</p>
  <label class="row">
    <span>{t('spkLocationRange')}</span>
    <input class="xl-input range" bind:value={locationText} spellcheck="false" />
  </label>
</Dialog>

<style>
  .head {
    margin: 4px 0 6px;
    font-weight: 600;
  }
  .row {
    display: grid;
    grid-template-columns: 110px 1fr;
    align-items: center;
    margin: 0 0 10px 12px;
  }
  .range {
    width: 100%;
    font-family: var(--xl-mono);
  }
</style>
