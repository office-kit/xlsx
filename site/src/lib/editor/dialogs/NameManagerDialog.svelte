<script lang="ts">
  import { fromStorageFormula, type CalcScalar } from '../calc/index.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  type Filter = 'all' | 'sheet' | 'workbook' | 'errors' | 'noErrors';

  const ctl = getEditor();
  const doc = ctl.doc;
  let filter = $state<Filter>('all');
  let selected = $state<number | null>(null);

  function show(v: CalcScalar): { text: string; error: boolean } {
    if (v === null) return { text: '', error: false };
    if (typeof v === 'object') return { text: v.code, error: true };
    if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', error: false };
    return { text: String(v), error: false };
  }

  const rows = $derived.by(() => {
    void doc.version;
    const sheetTitle = doc.ws.title;
    return doc.wb.definedNames
      .map((dn, index) => ({ dn, index }))
      .filter(({ dn }) => !dn.hidden)
      .map(({ dn, index }) => {
        const value = show(doc.calc.evaluate(dn.value.replace(/^=/, ''), dn.scope === undefined ? sheetTitle : (doc.wb.sheets[dn.scope]?.sheet.title ?? sheetTitle), 1, 1));
        return {
          index,
          name: dn.name.replace(/^_xlnm\./, ''),
          value,
          refersTo: `=${fromStorageFormula(dn.value.replace(/^=/, ''))}`,
          scope: dn.scope === undefined ? t('dlgWorkbook') : (doc.wb.sheets[dn.scope]?.sheet.title ?? ''),
          comment: dn.comment ?? '',
          sheetScoped: dn.scope !== undefined,
        };
      })
      .filter((r) => (filter === 'sheet' ? r.sheetScoped : filter === 'workbook' ? !r.sheetScoped : filter === 'errors' ? r.value.error : filter === 'noErrors' ? !r.value.error : true))
      .sort((a, b) => a.name.localeCompare(b.name));
  });

  function remove(): void {
    const index = selected;
    if (index === null) return;
    doc.transact('Delete Name', (tx) => {
      tx.structural = true;
      tx.workbook('definedNames');
      doc.wb.definedNames = doc.wb.definedNames.filter((_, i) => i !== index);
    });
    selected = null;
  }
</script>

<Dialog title={t('nameManager')} width={680} showCancel={false} okLabel={t('dlgClose')}>
  <div class="toolbar">
    <button class="xl-btn outlined" onclick={() => ctl.openDialog('defineName', { returnTo: 'nameManager' })}>{t('dlgNewEllipsis')}</button>
    <button class="xl-btn outlined" disabled={selected === null} onclick={() => ctl.openDialog('defineName', { edit: selected, returnTo: 'nameManager' })}>{t('dlgEditEllipsis')}</button>
    <button class="xl-btn outlined" disabled={selected === null} onclick={remove}>{t('delete')}</button>
    <label class="filter">
      {t('dlgFilter')}
      <select class="xl-select" bind:value={filter}>
        <option value="all">{t('dlgFilterAll')}</option>
        <option value="sheet">{t('dlgFilterSheet')}</option>
        <option value="workbook">{t('dlgFilterWorkbook')}</option>
        <option value="errors">{t('dlgFilterErrors')}</option>
        <option value="noErrors">{t('dlgFilterNoErrors')}</option>
      </select>
    </label>
  </div>
  <div class="table list" role="listbox" aria-label={t('nameManager')}>
    <div class="head" aria-hidden="true">
      <span>{t('dlgName')}</span><span>{t('dlgValue')}</span><span>{t('dlgRefersTo')}</span><span>{t('dlgScope')}</span><span>{t('dlgComment')}</span>
    </div>
    {#each rows as r (r.index)}
      <button role="option" aria-selected={selected === r.index} onclick={() => (selected = r.index)} ondblclick={() => ctl.openDialog('defineName', { edit: r.index, returnTo: 'nameManager' })}>
        <span>{r.name}</span><span class:err={r.value.error}>{r.value.text}</span><span>{r.refersTo}</span><span>{r.scope}</span><span>{r.comment}</span>
      </button>
    {:else}
      <div class="hint empty">{t('noNames')}</div>
    {/each}
  </div>
  {#if selected !== null}
    {@const dn = doc.wb.definedNames[selected]}
    <div class="row">
      <span class="lbl">{t('dlgRefersTo')}</span>
      <input class="xl-input grow mono" readonly value={dn ? `=${fromStorageFormula(dn.value.replace(/^=/, ''))}` : ''} aria-label={t('dlgRefersTo')} />
    </div>
  {/if}
</Dialog>

<style>
  .toolbar {
    display: flex;
    gap: 6px;
    align-items: center;
    margin-bottom: 8px;
  }
  .filter {
    margin-left: auto;
    display: flex;
    gap: 6px;
    align-items: center;
  }
  .table {
    height: 240px;
  }
  .head,
  .table button {
    display: grid;
    grid-template-columns: 1.2fr 1fr 1.6fr 0.8fr 1fr;
    gap: 8px;
  }
  .head {
    padding: 3px 6px;
    font-weight: 600;
    border-bottom: 1px solid var(--xl-border);
    background: #f7f7f7;
    position: sticky;
    top: 0;
  }
  .table span {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .err {
    color: var(--xl-danger);
  }
  .empty {
    padding: 8px;
  }
  .mono {
    font-family: var(--xl-mono);
  }
</style>
