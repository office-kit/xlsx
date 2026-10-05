<script lang="ts">
  import { untrack } from 'svelte';
  import type { DefinedName } from '@office-kit/xlsx/workbook';
  import { CalcParseError, fromStorageFormula, parseFormula, toStorageFormula } from '../calc/index.ts';
  import { getEditor } from '../core/context.ts';
  import { selectionRefText, validateName } from '../core/names.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();
  // DialogHost re-creates a dialog for every opening, so its arguments are read once.
  const args = untrack(() => props) ?? {};

  const ctl = getEditor();
  const doc = ctl.doc;
  const editIndex = typeof args['edit'] === 'number' ? args['edit'] : undefined;
  const editing = editIndex === undefined ? undefined : doc.wb.definedNames[editIndex];
  const backToManager = args['returnTo'] === 'nameManager';
  const sheets = doc.wb.sheets.map((s, index) => ({ index, title: s.sheet.title }));

  let name = $state(editing?.name ?? '');
  /** -1 is workbook scope; otherwise a sheet index. */
  let scope = $state(editing?.scope ?? -1);
  let comment = $state(editing?.comment ?? '');
  let refersTo = $state(editing ? `=${fromStorageFormula(editing.value.replace(/^=/, ''))}` : `=${selectionRefText(ctl)}`);
  let notice = $state<string | null>(null);

  function parses(formula: string): boolean {
    try {
      parseFormula(formula);
      return true;
    } catch (e) {
      if (e instanceof CalcParseError) return false;
      throw e;
    }
  }

  function onok(): boolean {
    const trimmed = name.trim();
    if (validateName(trimmed)) {
      notice = t('invalidName');
      return false;
    }
    const target = scope < 0 ? undefined : scope;
    const clash = doc.wb.definedNames.some((d, i) => i !== editIndex && d.scope === target && d.name.toLowerCase() === trimmed.toLowerCase());
    if (clash) {
      notice = t('duplicateName');
      return false;
    }
    const value = toStorageFormula(refersTo.trim().replace(/^=/, ''));
    if (!value || !parses(value)) {
      notice = t('formulaError');
      return false;
    }
    const dn: DefinedName = { name: trimmed, value, ...(target !== undefined ? { scope: target } : {}), ...(comment.trim() ? { comment: comment.trim() } : {}) };
    doc.transact(editing ? 'Edit Name' : 'Define Name', (tx) => {
      tx.structural = true;
      tx.workbook('definedNames');
      doc.wb.definedNames = editIndex === undefined ? [...doc.wb.definedNames, dn] : doc.wb.definedNames.map((d, i) => (i === editIndex ? dn : d));
    });
    if (backToManager) ctl.openDialog('nameManager');
    return true;
  }

  function oncancel(): void {
    if (backToManager) ctl.openDialog('nameManager');
  }
</script>

<Dialog title={t(editing ? 'dlgEditName' : 'dlgNewName')} width={420} {onok} {oncancel} bind:notice>
  <div class="row">
    <label for="dn-name">{t('dlgName')}</label>
    <input id="dn-name" class="xl-input grow" bind:value={name} spellcheck="false" />
  </div>
  <div class="row">
    <label for="dn-scope">{t('dlgScope')}</label>
    <select id="dn-scope" class="xl-select grow" bind:value={scope} disabled={editing !== undefined}>
      <option value={-1}>{t('dlgWorkbook')}</option>
      {#each sheets as s (s.index)}<option value={s.index}>{s.title}</option>{/each}
    </select>
  </div>
  <div class="row top">
    <label for="dn-comment">{t('dlgComment')}</label>
    <textarea id="dn-comment" class="xl-input grow area" bind:value={comment}></textarea>
  </div>
  <div class="row">
    <label for="dn-ref">{t('dlgRefersTo')}</label>
    <input id="dn-ref" class="xl-input grow mono" bind:value={refersTo} spellcheck="false" />
  </div>
</Dialog>

<style>
  .top {
    align-items: flex-start;
  }
  .area {
    height: 64px;
    resize: vertical;
  }
  .mono {
    font-family: var(--xl-mono);
  }
</style>
