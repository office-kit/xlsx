<script lang="ts">
  import { getCellDisplayText } from '@office-kit/xlsx/styles';
  import { cellAddress, parseRangeAddress, type CellPos } from '../core/address.ts';
  import { getCellAt } from '../core/cells.ts';
  import { getEditor } from '../core/context.ts';
  import { goalSeek, goalSeekError, setGoalSeekValue, type GoalSeekResult } from '../core/what-if.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';
  import RefInput from './RefInput.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  const { row, col } = doc.selection.active;

  let setText = $state(cellAddress(row, col, true));
  let targetText = $state('');
  let changingText = $state('');
  let notice = $state<string | null>(null);
  /** After the search: Excel's status box, whose OK keeps the value and Cancel drops it. */
  let status = $state<{ setCell: CellPos; changing: CellPos; target: number; result: GoalSeekResult } | null>(null);

  function cellOn(text: string): CellPos | undefined {
    const parsed = parseRangeAddress(text.trim().replace(/^=/, ''));
    if (!parsed || (parsed.sheet !== undefined && parsed.sheet.toLowerCase() !== doc.ws.title.toLowerCase())) return undefined;
    const r = parsed.range;
    return r.r1 === r.r2 && r.c1 === r.c2 ? { row: r.r1, col: r.c1 } : undefined;
  }

  function shown(v: unknown, at: CellPos): string {
    if (typeof v !== 'number') return String(v ?? '');
    const cell = getCellAt(doc.ws, at.row, at.col);
    // Show the numbers in the set cell's format, as Excel's status box does.
    return cell ? getCellDisplayText(doc.wb, { ...cell, value: v }) : String(v);
  }

  function onok(): boolean {
    if (status) {
      // Excel keeps the closest value even when it found no exact solution.
      setGoalSeekValue(ctl, status.changing, status.result.value);
      return true;
    }
    const setCell = cellOn(setText);
    const changing = cellOn(changingText);
    const target = Number(targetText.trim());
    if (!setCell || !changing) {
      notice = t('dtGsNeedCell');
      return false;
    }
    if (targetText.trim() === '' || !Number.isFinite(target)) {
      notice = t('dtGsNeedNumber');
      return false;
    }
    const err = goalSeekError(ctl, setCell, changing);
    if (err) {
      notice = t(err);
      return false;
    }
    status = { setCell, changing, target, result: goalSeek(ctl, setCell, target, changing) };
    return false;
  }
</script>

<Dialog title={status ? t('dtGoalSeekStatus') : t('dtGoalSeek')} width={340} modal={false} {onok} bind:notice>
  {#if status}
    <p>{t(status.result.found ? 'dtGsFound' : 'dtGsNotFound', { cell: cellAddress(status.setCell.row, status.setCell.col) })}</p>
    <div class="grid">
      <span>{t('dtGsTargetValue')}</span><span class="num">{shown(status.target, status.setCell)}</span>
      <span>{t('dtGsCurrentValue')}</span><span class="num">{shown(status.result.result, status.setCell)}</span>
    </div>
  {:else}
    <div class="grid">
      <label for="gs-set">{t('dtGsSetCell')}</label>
      <RefInput id="gs-set" bind:value={setText} autofocus />
      <label for="gs-to">{t('dtGsToValue')}</label>
      <input id="gs-to" class="xl-input" bind:value={targetText} inputmode="decimal" />
      <label for="gs-by">{t('dtGsByChanging')}</label>
      <RefInput id="gs-by" bind:value={changingText} />
    </div>
  {/if}
</Dialog>

<style>
  .grid {
    display: grid;
    grid-template-columns: auto 1fr;
    gap: 6px 10px;
    align-items: center;
  }
  .num {
    text-align: right;
    font-variant-numeric: tabular-nums;
  }
</style>
