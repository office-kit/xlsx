<script lang="ts">
  import { makeCell } from '@office-kit/xlsx/cell';
  import { cellAddress, MAX_COL, MAX_ROW, quoteSheetName } from '../core/address.ts';
  import { usedRange } from '../core/cells.ts';
  import { pasteSpecial, type PasteSpecialOptions } from '../core/clipboard.ts';
  import { getEditor } from '../core/context.ts';
  import { currentRange } from '../core/selection.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const WHAT: ReadonlyArray<[PasteSpecialOptions['what'], MessageKey]> = [
    ['all', 'dlgPsAll'],
    ['allExceptBorders', 'dlgPsAllExceptBorders'],
    ['formulas', 'pasteFormulas'],
    ['columnWidths', 'dlgPsColumnWidths'],
    ['values', 'pasteValues'],
    ['formulasAndNumberFormats', 'dlgPsFormulasNumberFormats'],
    ['formats', 'dlgPsFormats'],
    ['valuesAndNumberFormats', 'dlgPsValuesNumberFormats'],
    ['comments', 'dlgPsComments'],
    ['validation', 'dlgPsValidation'],
  ];
  const OPERATIONS: ReadonlyArray<[PasteSpecialOptions['operation'], MessageKey]> = [
    ['none', 'dlgNone'],
    ['add', 'dlgOpAdd'],
    ['subtract', 'dlgOpSubtract'],
    ['multiply', 'dlgOpMultiply'],
    ['divide', 'dlgOpDivide'],
  ];

  const ctl = getEditor();
  let what = $state<PasteSpecialOptions['what']>('all');
  let operation = $state<PasteSpecialOptions['operation']>('none');
  let skipBlanks = $state(false);
  let transpose = $state(false);
  let notice = $state<string | null>(null);
  const hasSource = ctl.clipboard !== null;

  function onok(): boolean {
    if (!hasSource) {
      notice = t('dlgNothingToPaste');
      return false;
    }
    pasteSpecial(ctl, { what, operation, skipBlanks, transpose });
    return true;
  }

  /** Paste Link: formulas pointing back at each copied cell (absolute for a single cell, as Excel does). */
  function pasteLink(): void {
    const clip = ctl.clipboard;
    const sourceRef = clip ? ctl.doc.wb.sheets[clip.sheetIndex] : undefined;
    if (!clip || sourceRef?.kind !== 'worksheet') {
      notice = t('dlgNothingToPaste');
      return;
    }
    const used = usedRange(sourceRef.sheet);
    const src = clip.range;
    const r2 = src.r2 === MAX_ROW ? Math.max(src.r1, Math.min(src.r2, used?.r2 ?? src.r1)) : src.r2;
    const c2 = src.c2 === MAX_COL ? Math.max(src.c1, Math.min(src.c2, used?.c2 ?? src.c1)) : src.c2;
    const doc = ctl.doc;
    const ws = doc.ws;
    const origin = currentRange(doc.selection);
    const rows = Math.min(r2 - src.r1 + 1, MAX_ROW - origin.r1 + 1);
    const cols = Math.min(c2 - src.c1 + 1, MAX_COL - origin.c1 + 1);
    const prefix = clip.sheetIndex === doc.activeSheetIndex ? '' : `${quoteSheetName(sourceRef.sheet.title)}!`;
    const single = rows === 1 && cols === 1;
    doc.transact('Paste Link', (tx) => {
      tx.cells(ws, { r1: origin.r1, c1: origin.c1, r2: origin.r1 + rows - 1, c2: origin.c1 + cols - 1 });
      for (let dr = 0; dr < rows; dr++) {
        for (let dc = 0; dc < cols; dc++) {
          const row = origin.r1 + dr;
          const col = origin.c1 + dc;
          const value = { kind: 'formula' as const, t: 'normal' as const, formula: prefix + cellAddress(src.r1 + dr, src.c1 + dc, single) };
          let rowMap = ws.rows.get(row);
          const existing = rowMap?.get(col);
          if (existing) existing.value = value;
          else {
            if (!rowMap) {
              rowMap = new Map();
              ws.rows.set(row, rowMap);
            }
            rowMap.set(col, makeCell(row, col, value, ctl.defaultStyleAt(row, col)));
          }
        }
      }
    });
    ctl.selectRange({ r1: origin.r1, c1: origin.c1, r2: origin.r1 + rows - 1, c2: origin.c1 + cols - 1 });
    ctl.closeDialog();
  }
</script>

<Dialog title={t('dlgPasteSpecial')} width={440} {onok} bind:notice>
  {#snippet footerStart()}
    <button class="xl-btn outlined" disabled={!hasSource} onclick={pasteLink}>{t('dlgPasteLink')}</button>
  {/snippet}
  <fieldset>
    <legend>{t('paste')}</legend>
    <div class="grid" role="radiogroup" aria-label={t('paste')}>
      {#each WHAT as [value, label] (value)}
        <label class="check"><input type="radio" name="ps-what" {value} bind:group={what} />{t(label)}</label>
      {/each}
    </div>
  </fieldset>
  <fieldset>
    <legend>{t('dlgOperation')}</legend>
    <div class="grid" role="radiogroup" aria-label={t('dlgOperation')}>
      {#each OPERATIONS as [value, label] (value)}
        <label class="check"><input type="radio" name="ps-op" {value} bind:group={operation} />{t(label)}</label>
      {/each}
    </div>
  </fieldset>
  <div class="grid">
    <label class="check"><input type="checkbox" bind:checked={skipBlanks} />{t('dlgSkipBlanks')}</label>
    <label class="check"><input type="checkbox" bind:checked={transpose} />{t('pasteTranspose')}</label>
  </div>
</Dialog>

<style>
  .grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    column-gap: 12px;
  }
</style>
