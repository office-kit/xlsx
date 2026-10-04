<script lang="ts">
  // Insert ▸ PivotTable, and PivotTable Analyze ▸ Change Data Source (props.change
  // set): pick the source range and, when creating, where the report goes.
  import { getEditor } from '../core/context.ts';
  import { currentRegion } from '../core/navigation.ts';
  import { cellAddress, quoteSheetName } from '../core/address.ts';
  import { createPivotTable, editPivot, parseDestination, parseSource, sourceText } from '../core/pivot.ts';
  import { currentRange } from '../core/selection.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();

  const ctl = getEditor();
  const doc = ctl.doc;
  const ws = doc.ws;
  const active = ctl.activePivot;
  const changing = props?.['change'] === true && active?.kind === 'model' ? active : undefined;

  function initialSource(): string {
    if (changing) return sourceText(changing.pt.source);
    const sel = currentRange(doc.selection);
    const r = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(ws, doc.selection.active) : sel;
    return sourceText({ sheet: ws.title, ref: `${cellAddress(r.r1, r.c1)}:${cellAddress(r.r2, r.c2)}` });
  }

  let sourceInput = $state(initialSource());
  let place = $state<'new' | 'existing'>('new');
  let locationInput = $state('');
  let notice = $state<string | null>(null);

  function fail(key: MessageKey): false {
    notice = t(key);
    return false;
  }

  function onok(): boolean {
    const source = parseSource(doc.wb, sourceInput, ws.title);
    if (!source) return fail('invalidReference');
    if (changing) {
      editPivot(ctl, changing.ws, changing.index, 'Change PivotTable Data Source', (pt) => {
        pt.source = source;
      });
      return true;
    }
    let destination: { sheet: string; cell: string } | undefined;
    if (place === 'existing') {
      destination = parseDestination(doc.wb, locationInput, ws.title);
      if (!destination) return fail('invalidReference');
    }
    const err = createPivotTable(ctl, destination ? { source, destination } : { source });
    // Declined or refused: the user already saw why, so the dialog just closes.
    if (err === 'cancelled') return true;
    if (err) return fail(err);
    return true;
  }

  // Picking "Existing Worksheet" proposes the active cell, as Excel fills the box.
  $effect(() => {
    if (place === 'existing' && locationInput === '') {
      const { row, col } = doc.selection.active;
      locationInput = `${quoteSheetName(ws.title)}!${cellAddress(row, col, true)}`;
    }
  });
</script>

<Dialog title={t(changing ? 'dlgChangePivotSource' : 'dlgCreatePivot')} width={400} {onok} bind:notice>
  <div class="section">{t('pvChooseData')}</div>
  <label class="row">
    <span>{t('pvTableRange')}</span>
    <input class="xl-input ref" bind:value={sourceInput} spellcheck="false" data-autofocus />
  </label>
  {#if !changing}
    <div class="section">{t('pvChoosePlace')}</div>
    <label class="radio"><input type="radio" bind:group={place} value="new" />{t('pvNewSheet')}</label>
    <label class="radio"><input type="radio" bind:group={place} value="existing" />{t('pvExistingSheet')}</label>
    <label class="row">
      <span>{t('pvLocation')}</span>
      <input class="xl-input ref" bind:value={locationInput} spellcheck="false" disabled={place !== 'existing'} />
    </label>
  {/if}
</Dialog>

<style>
  .section {
    font-weight: 600;
    margin: 6px 0 4px;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 6px;
    margin: 2px 0 6px 12px;
  }
  .row span {
    min-width: 84px;
  }
  .ref {
    flex: 1;
    font-family: var(--xl-mono);
  }
  .radio {
    display: flex;
    align-items: center;
    gap: 6px;
    margin-left: 12px;
  }
</style>
