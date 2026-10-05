<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { columnLabel, dataRange, guessHeader, removeDuplicates } from '../core/data.ts';
  import { selectRange } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  const range = dataRange(ctl);
  const columns = Array.from({ length: range.c2 - range.c1 + 1 }, (_, i) => range.c1 + i);
  let hasHeader = $state(guessHeader(ctl.doc.ws, range));
  const checked = $state<Record<number, boolean>>(Object.fromEntries(columns.map((c) => [c, true])));
  let notice = $state<string | null>(null);

  function onok(): boolean {
    const cols = columns.filter((c) => checked[c]);
    if (cols.length === 0) {
      notice = t('dlgDupNoColumns');
      return false;
    }
    const rows = range.r2 - range.r1 + 1 - (hasHeader ? 1 : 0);
    ctl.doc.setSelection(selectRange(range, ctl.doc.selection.active));
    const removed = removeDuplicates(ctl, range, cols, hasHeader);
    if (removed === 'outlined') {
      notice = t('dlgDupOutlined');
      return false;
    }
    // Sheet protection refused it and has already told the user why.
    if (removed === undefined) return true;
    const detail = removed === 0 ? t('dlgDupNone') : t('dlgDupResult', { n: removed, m: rows - removed });
    ctl.dialog = { kind: 'alert', props: { message: 'removeDuplicates', detail } };
    return true;
  }

  function setAll(on: boolean): void {
    for (const c of columns) checked[c] = on;
  }
</script>

<Dialog title={t('removeDuplicates')} width={380} {onok} bind:notice>
  <p class="hint">{t('dlgDupIntro')}</p>
  <div class="toolbar">
    <button class="xl-btn outlined" onclick={() => setAll(true)}>{t('dlgSelectAll')}</button>
    <button class="xl-btn outlined" onclick={() => setAll(false)}>{t('dlgUnselectAll')}</button>
    <label class="check header"><input type="checkbox" bind:checked={hasHeader} />{t('dlgMyDataHasHeaders')}</label>
  </div>
  <div class="list cols">
    {#each columns as c (c)}
      <label class="check"><input type="checkbox" bind:checked={checked[c]} />{columnLabel(ctl, range, c, hasHeader, t('dlgColumn'))}</label>
    {/each}
  </div>
</Dialog>

<style>
  .toolbar {
    display: flex;
    gap: 6px;
    align-items: center;
    margin-bottom: 6px;
  }
  .header {
    margin-left: auto;
  }
  .cols {
    height: 170px;
    padding: 2px 6px;
  }
</style>
