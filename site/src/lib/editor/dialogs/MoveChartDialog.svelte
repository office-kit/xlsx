<script lang="ts">
  // Chart Design ▸ Move Chart: move the selected chart to a new chart sheet
  // or, as an object, onto another worksheet at the same anchor. One undo
  // step covering both sheets.
  import { makeDrawing } from '@office-kit/xlsx/drawing';
  import { addChartsheet } from '@office-kit/xlsx/workbook';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  const index = ctl.selectedDrawing;
  const from = doc.ws;

  const taken = (name: string) => doc.wb.sheets.some((s) => s.sheet.title.toLowerCase() === name.toLowerCase());
  function freeName(): string {
    const base = t('chChartSheetName');
    let n = 1;
    while (taken(`${base}${n}`)) n++;
    return `${base}${n}`;
  }

  const targets = doc.wb.sheets.flatMap((s, i) => (s.kind === 'worksheet' ? [{ index: i, title: s.sheet.title }] : []));
  let mode = $state<'sheet' | 'object'>('sheet');
  let name = $state(freeName());
  let target = $state(doc.activeSheetIndex);
  let notice = $state<string | null>(null);

  function ok(): boolean {
    const item = index === null ? undefined : from.drawing?.items[index];
    if (index === null || item?.content.kind !== 'chart') return true;
    const chart = item.content.chart;
    if (mode === 'sheet') {
      const title = name.trim();
      if (!title || taken(title)) {
        notice = title ? t('chSheetExists', { name: title }) : t('invalidSheetName');
        return false;
      }
      const moved = doc.transact('Move Chart', (tx) => {
        tx.workbook('sheets');
        tx.sheet(from, 'drawing');
        from.drawing?.items.splice(index, 1);
        // After the active sheet, so the sheet in view and its index stay put.
        addChartsheet(doc.wb, title, { chart, index: doc.activeSheetIndex + 1 });
        return true;
      });
      // Undefined: sheet protection refused the move. Returning false keeps
      // Dialog from closing the protection alert that replaced this dialog.
      if (!moved) return false;
      ctl.selectedDrawing = null;
      // The notice replaces this dialog; returning true would close it instead.
      ctl.openDialog('alert', { message: t('chMovedToChartsheet', { name: title }) });
      return false;
    }
    const ref = doc.wb.sheets[target];
    if (ref?.kind !== 'worksheet' || ref.sheet === from) return true;
    const to = ref.sheet;
    const moved = doc.transact('Move Chart', (tx) => {
      tx.sheet(from, 'drawing');
      tx.sheet(to, 'drawing');
      from.drawing?.items.splice(index, 1);
      to.drawing ??= makeDrawing();
      to.drawing.items.push(item);
      return true;
    });
    if (!moved) return false;
    ctl.selectedDrawing = null;
    doc.activateSheet(target);
    return true;
  }
</script>

<Dialog title={t('chMoveChartTitle')} width={420} onok={ok} bind:notice>
  <p class="hint">{t('chMoveChoose')}</p>
  <div class="row">
    <label class="check"><input type="radio" name="mc-mode" checked={mode === 'sheet'} onchange={() => (mode = 'sheet')} />{t('chNewSheet')}</label>
    <input class="xl-input grow" bind:value={name} disabled={mode !== 'sheet'} maxlength="31" spellcheck="false" aria-label={t('chNewSheet')} />
  </div>
  <div class="row">
    <label class="check"><input type="radio" name="mc-mode" checked={mode === 'object'} onchange={() => (mode = 'object')} />{t('chObjectIn')}</label>
    <select class="xl-input grow" bind:value={target} disabled={mode !== 'object'} aria-label={t('chObjectIn')}>
      {#each targets as s (s.index)}
        <option value={s.index}>{s.title}</option>
      {/each}
    </select>
  </div>
</Dialog>
