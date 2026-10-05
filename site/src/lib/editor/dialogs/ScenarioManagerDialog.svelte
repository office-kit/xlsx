<script lang="ts">
  // Scenario Manager with its sub-pages — Add / Edit Scenario, Scenario
  // Values and Scenario Summary — as pages of one dialog; Cancel on a
  // sub-page returns to the list, as Excel's stacked dialogs do.
  import { cellAddress } from '../core/address.ts';
  import { getEditor } from '../core/context.ts';
  import {
    changingCells,
    currentValues,
    deleteScenario,
    saveScenario,
    scenarioNameError,
    scenarioSummary,
    showScenario,
    type ScenarioDraft,
  } from '../core/what-if.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';
  import RefInput from './RefInput.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;

  const scenarios = $derived.by(() => {
    void doc.version;
    return doc.ws.scenarios?.scenarios ?? [];
  });

  type Page = 'list' | 'edit' | 'values' | 'summary';
  let page = $state<Page>('list');
  let selected = $state(0);
  let editing = $state<number | undefined>(undefined);
  let draft = $state<ScenarioDraft>({ name: '', cells: '', comment: '', locked: true, hidden: false });
  let refs = $state<string[]>([]);
  let values = $state<string[]>([]);
  let resultText = $state('');
  let notice = $state<string | null>(null);

  const current = $derived(scenarios[selected]);

  function startAdd(): void {
    const { row, col } = doc.selection.active;
    editing = undefined;
    draft = { name: '', cells: cellAddress(row, col), comment: '', locked: true, hidden: false };
    page = 'edit';
  }

  function startEdit(): void {
    const s = current;
    if (!s) return;
    editing = selected;
    draft = { name: s.name, cells: s.inputCells.map((c) => c.ref).join(','), comment: s.comment ?? '', locked: s.locked === true, hidden: s.hidden === true };
    page = 'edit';
  }

  function onok(): boolean {
    switch (page) {
      case 'list':
        return true;
      case 'edit': {
        const nameErr = scenarioNameError(ctl, draft.name, editing);
        if (nameErr) {
          notice = t(nameErr);
          return false;
        }
        const cells = changingCells(draft.cells);
        if (typeof cells === 'string') {
          notice = t(cells);
          return false;
        }
        const prior = editing === undefined ? undefined : scenarios[editing];
        const priorValues = new Map(prior?.inputCells.map((c) => [c.ref, c.val]));
        const live = currentValues(ctl, cells);
        refs = cells;
        values = cells.map((ref, i) => priorValues.get(ref) ?? live[i] ?? '');
        page = 'values';
        return false;
      }
      case 'values':
        saveScenario(ctl, draft, refs, values, editing);
        selected = editing ?? scenarios.length - 1;
        page = 'list';
        return false;
      case 'summary': {
        const result = resultText.trim() === '' ? [] : changingCells(resultText);
        if (typeof result === 'string') {
          notice = t(result);
          return false;
        }
        scenarioSummary(ctl, result, {
          title: t('dtScSummaryTitle'),
          current: t('dtScCurrentValues'),
          changing: t('dtScChangingCellsLabel'),
          result: t('dtScResultCellsLabel'),
          sheetName: t('dtScSummaryTitle'),
        });
        return true;
      }
    }
  }

  function oncancel(): boolean {
    if (page === 'list') return true;
    page = 'list';
    return false;
  }

  const absolute = (ref: string) => ref.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');

  const titles: Record<Page, () => string> = {
    list: () => t('dtScenarioManager'),
    edit: () => t(editing === undefined ? 'dtScAdd' : 'dtScEdit'),
    values: () => t('dtScValues'),
    summary: () => t('dtScSummary'),
  };
</script>

<Dialog title={titles[page]()} width={page === 'list' ? 440 : 360} modal={false} {onok} {oncancel} {...page === 'list' ? { okLabel: t('dlgClose') } : {}} showCancel={page !== 'list'} bind:notice>
  {#if page === 'list'}
    <div class="row top">
      <div class="list items" role="listbox" aria-label={t('dtScenarios')}>
        {#each scenarios as s, i (i)}
          <button role="option" aria-selected={i === selected} onclick={() => (selected = i)} ondblclick={() => showScenario(ctl, i)}>{s.name}</button>
        {:else}
          <p class="hint empty">{t('dtScNone')}</p>
        {/each}
      </div>
      <div class="buttons">
        <button class="xl-btn outlined" onclick={startAdd}>{t('dtScAddButton')}</button>
        <button class="xl-btn outlined" disabled={!current} onclick={() => deleteScenario(ctl, selected)}>{t('dtScDelete')}</button>
        <button class="xl-btn outlined" disabled={!current} onclick={startEdit}>{t('dtScEditButton')}</button>
        <button class="xl-btn outlined" disabled={scenarios.length === 0} onclick={() => (page = 'summary')}>{t('dtScSummaryButton')}</button>
        <button class="xl-btn outlined" disabled={!current} onclick={() => showScenario(ctl, selected)}>{t('dtScShow')}</button>
      </div>
    </div>
    <div class="info">
      <span>{t('dtScChangingCells')}</span><span class="mono">{current?.inputCells.map((c) => c.ref).join(',') ?? ''}</span>
      <span>{t('dtScComment')}</span><span>{current?.comment ?? ''}</span>
    </div>
  {:else if page === 'edit'}
    <label class="lbl" for="sc-name">{t('dtScName')}</label>
    <input id="sc-name" class="xl-input wide" bind:value={draft.name} data-autofocus />
    <label class="lbl" for="sc-cells">{t('dtScChangingCells')}</label>
    <RefInput id="sc-cells" bind:value={draft.cells} />
    <p class="hint">{t('dtScCellsHint')}</p>
    <label class="lbl" for="sc-comment">{t('dtScComment')}</label>
    <textarea id="sc-comment" class="xl-input wide" rows="3" bind:value={draft.comment}></textarea>
    <fieldset>
      <legend>{t('dtScProtection')}</legend>
      <label class="check"><input type="checkbox" bind:checked={draft.locked} />{t('dtScPreventChanges')}</label>
      <label class="check"><input type="checkbox" bind:checked={draft.hidden} />{t('dtScHide')}</label>
    </fieldset>
  {:else if page === 'values'}
    <p>{t('dtScValuesPrompt')}</p>
    <div class="values">
      {#each refs as ref, i (ref)}
        <label for="sc-v{i}">{i + 1}: <span class="mono">{absolute(ref)}</span></label>
        <input id="sc-v{i}" class="xl-input" bind:value={values[i]} data-autofocus={i === 0 ? '' : undefined} />
      {/each}
    </div>
  {:else}
    <fieldset>
      <legend>{t('dtScReportType')}</legend>
      <label class="check"><input type="radio" name="sc-report" checked />{t('dtScSummaryType')}</label>
      <label class="check dim"><input type="radio" name="sc-report" disabled />{t('dtScPivotType')}</label>
    </fieldset>
    <label class="lbl" for="sc-result">{t('dtScResultCells')}</label>
    <RefInput id="sc-result" bind:value={resultText} autofocus />
  {/if}
</Dialog>

<style>
  .top {
    align-items: flex-start;
  }
  .items {
    flex: 1;
    height: 150px;
    overflow-y: auto;
  }
  .empty {
    padding: 6px;
    white-space: normal;
  }
  .buttons {
    display: flex;
    flex-direction: column;
    gap: 4px;
    min-width: 96px;
  }
  .info {
    display: grid;
    grid-template-columns: auto 1fr;
    gap: 4px 10px;
    margin-top: 8px;
  }
  .lbl {
    display: block;
    margin: 6px 0 3px;
  }
  .wide {
    width: 100%;
    box-sizing: border-box;
  }
  .values {
    display: grid;
    grid-template-columns: auto 1fr;
    gap: 4px 10px;
    align-items: center;
    max-height: 260px;
    overflow-y: auto;
  }
  .mono {
    font-family: var(--xl-mono);
  }
  .dim {
    color: var(--xl-text-2);
  }
</style>
