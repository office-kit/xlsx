<script lang="ts">
  import { parseRangeAddress, type CellPos } from '../core/address.ts';
  import { getEditor } from '../core/context.ts';
  import { currentRange } from '../core/selection.ts';
  import { fillDataTable } from '../core/what-if.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';
  import RefInput from './RefInput.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  // The selection is the table: its first row / column hold the inputs and formulas.
  const range = currentRange(doc.selection);

  let rowText = $state('');
  let colText = $state('');
  let notice = $state<string | null>(null);

  /** undefined for an empty box; null for one that is not a cell on this sheet. */
  function cellOn(text: string): CellPos | undefined | null {
    if (text.trim() === '') return undefined;
    const parsed = parseRangeAddress(text.trim().replace(/^=/, ''));
    if (!parsed || (parsed.sheet !== undefined && parsed.sheet.toLowerCase() !== doc.ws.title.toLowerCase())) return null;
    const r = parsed.range;
    return r.r1 === r.r2 && r.c1 === r.c2 ? { row: r.r1, col: r.c1 } : null;
  }

  function onok(): boolean {
    const rowInput = cellOn(rowText);
    const colInput = cellOn(colText);
    if (rowInput === null || colInput === null) {
      notice = t('dtDtBadInput');
      return false;
    }
    const err = fillDataTable(ctl, range, rowInput, colInput);
    if (err) {
      notice = t(err);
      return false;
    }
    return true;
  }
</script>

<Dialog title={t('dtDataTable')} width={320} modal={false} {onok} bind:notice>
  <div class="grid">
    <label for="dt-row">{t('dtDtRowInput')}</label>
    <RefInput id="dt-row" bind:value={rowText} autofocus />
    <label for="dt-col">{t('dtDtColInput')}</label>
    <RefInput id="dt-col" bind:value={colText} />
  </div>
</Dialog>

<style>
  .grid {
    display: grid;
    grid-template-columns: auto 1fr;
    gap: 6px 10px;
    align-items: center;
  }
</style>
