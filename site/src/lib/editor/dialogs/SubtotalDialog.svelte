<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { columnLabel, dataRange } from '../core/data.ts';
  import { selectRange } from '../core/selection.ts';
  import { applySubtotals, removeSubtotals, SUBTOTAL_FUNCTIONS, type SubtotalFunction } from '../core/subtotal.ts';
  import { tableAt } from '../core/tables.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  const range = dataRange(ctl);
  const inTable = tableAt(doc.ws, range.r1, range.c1) !== undefined;
  const columns = Array.from({ length: range.c2 - range.c1 + 1 }, (_, i) => range.c1 + i);
  const label = (col: number) => columnLabel(ctl, range, col, true, t('dlgColumn'));

  /** Each function's label in the list and the words its added rows get. */
  const FUNCTIONS: Record<SubtotalFunction, { label: MessageKey; group: MessageKey; grand: MessageKey }> = {
    sum: { label: 'dtFnSum', group: 'dtSubTotal', grand: 'dtSubGrandTotal' },
    count: { label: 'dtFnCount', group: 'dtSubCount', grand: 'dtSubGrandCount' },
    average: { label: 'dtFnAverage', group: 'dtSubAverage', grand: 'dtSubGrandAverage' },
    max: { label: 'dtFnMax', group: 'dtSubMax', grand: 'dtSubGrandMax' },
    min: { label: 'dtFnMin', group: 'dtSubMin', grand: 'dtSubGrandMin' },
    product: { label: 'dtFnProduct', group: 'dtSubProduct', grand: 'dtSubGrandProduct' },
    countNums: { label: 'dtFnCountNums', group: 'dtSubCount', grand: 'dtSubGrandCount' },
    stdDev: { label: 'dtFnStdDev', group: 'dtSubStdDev', grand: 'dtSubGrandStdDev' },
    stdDevp: { label: 'dtFnStdDevp', group: 'dtSubStdDevp', grand: 'dtSubGrandStdDevp' },
    var: { label: 'dtFnVar', group: 'dtSubVar', grand: 'dtSubGrandVar' },
    varp: { label: 'dtFnVarp', group: 'dtSubVarp', grand: 'dtSubGrandVarp' },
  };

  let groupBy = $state(range.c1);
  let fn = $state<SubtotalFunction>('sum');
  const checked = $state<Record<number, boolean>>(Object.fromEntries(columns.map((c) => [c, c === range.c2 && c !== range.c1])));
  let replace = $state(true);
  let pageBreaks = $state(false);
  let summaryBelow = $state(true);
  let notice = $state<string | null>(null);

  function precheck(): boolean {
    if (inTable) {
      notice = t('dtSubInTable');
      return false;
    }
    if (range.r2 <= range.r1) {
      notice = t('dtSubNoList');
      return false;
    }
    return true;
  }

  function onok(): boolean {
    if (!precheck()) return false;
    const cols = columns.filter((c) => checked[c]);
    if (cols.length === 0) {
      notice = t('dtSubNoColumns');
      return false;
    }
    const words = FUNCTIONS[fn];
    const result = applySubtotals(ctl, range, { groupBy, fn, columns: cols, replace, pageBreaks, summaryBelow }, { group: (v) => t(words.group, { v }), grand: t(words.grand) });
    doc.setSelection(selectRange(result, doc.selection.active));
    return true;
  }

  function removeAll(): void {
    if (!precheck()) return;
    const result = removeSubtotals(ctl, range);
    doc.setSelection(selectRange(result, doc.selection.active));
    ctl.closeDialog();
  }
</script>

<Dialog title={t('dtSubtotal')} width={340} {onok} bind:notice>
  {#snippet footerStart()}
    <button class="xl-btn outlined" onclick={removeAll}>{t('dtSubRemoveAll')}</button>
  {/snippet}
  <label class="lbl" for="sub-group">{t('dtSubAtEachChange')}</label>
  <select id="sub-group" class="xl-select wide" bind:value={groupBy}>
    {#each columns as c (c)}<option value={c}>{label(c)}</option>{/each}
  </select>
  <label class="lbl" for="sub-fn">{t('dtSubUseFunction')}</label>
  <select id="sub-fn" class="xl-select wide" bind:value={fn}>
    {#each SUBTOTAL_FUNCTIONS as f (f)}<option value={f}>{t(FUNCTIONS[f].label)}</option>{/each}
  </select>
  <span class="lbl">{t('dtSubAddTo')}</span>
  <div class="list cols" role="group" aria-label={t('dtSubAddTo')}>
    {#each columns as c (c)}
      <label class="check"><input type="checkbox" bind:checked={checked[c]} />{label(c)}</label>
    {/each}
  </div>
  <label class="check"><input type="checkbox" bind:checked={replace} />{t('dtSubReplace')}</label>
  <label class="check"><input type="checkbox" bind:checked={pageBreaks} />{t('dtSubPageBreak')}</label>
  <label class="check"><input type="checkbox" bind:checked={summaryBelow} />{t('dtSubSummaryBelow')}</label>
</Dialog>

<style>
  .lbl {
    display: block;
    margin: 6px 0 3px;
  }
  .wide {
    width: 100%;
  }
  .cols {
    height: 120px;
    overflow-y: auto;
    padding: 4px 6px;
    margin-bottom: 6px;
  }
</style>
