<script lang="ts">
  import { untrack } from 'svelte';
  // Find and Replace: one modeless dialog with a Find and a Replace mode, as in
  // Excel. Enter runs Find Next; the dialog stays open while the user works.
  import type { Worksheet } from '@office-kit/xlsx/worksheet';
  import { cellAddress } from '../core/address.ts';
  import { getEditor } from '../core/context.ts';
  import { findAll, replaceAll, replaceCurrent, type FindOptions } from '../core/find.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();
  // DialogHost re-creates a dialog for every opening, so its arguments are read once.
  const args = untrack(() => props) ?? {};

  interface Hit {
    readonly sheetIndex: number;
    readonly sheet: string;
    readonly row: number;
    readonly col: number;
    readonly text: string;
  }

  const ctl = getEditor();
  const fs = ctl.findState;
  let mode = $state<'find' | 'replace'>(args['kind'] === 'replace' ? 'replace' : 'find');
  let results = $state.raw<Hit[]>([]);
  let selected = $state(-1);
  let notice = $state<string | null>(null);

  function options(): FindOptions {
    return { query: fs.query, matchCase: fs.matchCase, wholeCell: fs.wholeCell, byColumns: fs.byColumns, lookIn: mode === 'replace' ? 'formulas' : fs.lookIn };
  }

  /** Worksheets in search order: the active sheet first, then the rest in tab order. */
  function scope(): Array<{ ws: Worksheet; index: number }> {
    const sheets = ctl.doc.wb.sheets;
    const active = ctl.doc.activeSheetIndex;
    const out: Array<{ ws: Worksheet; index: number }> = [];
    const count = fs.inWorkbook ? sheets.length : 1;
    for (let k = 0; k < count; k++) {
      const index = (active + k) % sheets.length;
      const ref = sheets[index];
      if (ref?.kind === 'worksheet' && ref.state === 'visible') out.push({ ws: ref.sheet, index });
    }
    return out;
  }

  function search(): Hit[] {
    const sheets = scope();
    const indexOf = new Map(sheets.map((s) => [s.ws, s.index]));
    return findAll(
      ctl,
      options(),
      sheets.map((s) => s.ws),
    ).map((m) => ({ sheetIndex: indexOf.get(m.ws) ?? 0, sheet: m.ws.title, row: m.cell.row, col: m.cell.col, text: m.text }));
  }

  function goTo(hit: Hit): void {
    ctl.doc.activateSheet(hit.sheetIndex);
    ctl.selectCell({ row: hit.row, col: hit.col });
    ctl.reveal(hit.row, hit.col);
  }

  function findNextHit(): boolean {
    if (!fs.query) return false;
    const hits = search();
    if (hits.length === 0) {
      notice = t('notFound');
      return false;
    }
    const sheet = ctl.doc.activeSheetIndex;
    const a = ctl.doc.selection.active;
    const after = (h: Hit) => h.sheetIndex === sheet && (fs.byColumns ? h.col > a.col || (h.col === a.col && h.row > a.row) : h.row > a.row || (h.row === a.row && h.col > a.col));
    // Hits are ordered active sheet first, so the first non-active-sheet hit is the next sheet's.
    const next = hits.find(after) ?? hits.find((h) => h.sheetIndex !== sheet) ?? hits[0];
    if (next) goTo(next);
    return true;
  }

  function onok(): boolean {
    findNextHit();
    return false;
  }

  function findAllHits(): void {
    if (!fs.query) return;
    results = search();
    selected = -1;
    if (results.length === 0) notice = t('notFound');
  }

  function replaceOne(): void {
    if (!fs.query) return;
    if (search().length === 0) {
      notice = t('notFound');
      return;
    }
    const result = replaceCurrent(ctl, options(), fs.replace);
    if (result.invalidFormula !== undefined) notice = `${t('formulaError')} ${result.invalidFormula}`;
  }

  function replaceEverything(): void {
    if (!fs.query) return;
    const result = replaceAll(
      ctl,
      options(),
      fs.replace,
      scope().map((s) => s.ws),
    );
    results = [];
    if (result.invalidFormula !== undefined) notice = `${t('formulaError')} ${result.invalidFormula}`;
    else notice = result.replaced === 0 ? t('notFound') : t('dlgReplacedCount', { n: result.replaced });
  }
</script>

<Dialog title={t('dlgFindReplace')} width={480} modal={false} {onok} bind:notice>
  <div class="seg" role="tablist">
    <button role="tab" aria-selected={mode === 'find'} onclick={() => (mode = 'find')}>{t('dlgFind')}</button>
    <button role="tab" aria-selected={mode === 'replace'} onclick={() => (mode = 'replace')}>{t('dlgReplace')}</button>
  </div>
  <div class="row">
    <label for="fr-what">{t('dlgFindWhat')}</label>
    <input id="fr-what" class="xl-input grow" bind:value={fs.query} spellcheck="false" />
  </div>
  {#if mode === 'replace'}
    <div class="row">
      <label for="fr-with">{t('dlgReplaceWith')}</label>
      <input id="fr-with" class="xl-input grow" bind:value={fs.replace} spellcheck="false" />
    </div>
  {/if}
  <div class="opts">
    <div>
      <div class="row">
        <label for="fr-within">{t('dlgWithin')}</label>
        <select id="fr-within" class="xl-select" value={fs.inWorkbook ? 'workbook' : 'sheet'} onchange={(e) => (fs.inWorkbook = e.currentTarget.value === 'workbook')}>
          <option value="sheet">{t('dlgSheet')}</option>
          <option value="workbook">{t('dlgWorkbook')}</option>
        </select>
      </div>
      <div class="row">
        <label for="fr-by">{t('dlgSearch')}</label>
        <select id="fr-by" class="xl-select" value={fs.byColumns ? 'columns' : 'rows'} onchange={(e) => (fs.byColumns = e.currentTarget.value === 'columns')}>
          <option value="rows">{t('dlgByRows')}</option>
          <option value="columns">{t('dlgByColumns')}</option>
        </select>
      </div>
      <div class="row">
        <label for="fr-in">{t('dlgLookIn')}</label>
        <select id="fr-in" class="xl-select" bind:value={fs.lookIn} disabled={mode === 'replace'}>
          <option value="formulas">{t('pasteFormulas')}</option>
          <option value="values">{t('pasteValues')}</option>
        </select>
      </div>
    </div>
    <div>
      <label class="check"><input type="checkbox" bind:checked={fs.matchCase} />{t('dlgMatchCase')}</label>
      <label class="check"><input type="checkbox" bind:checked={fs.wholeCell} />{t('dlgMatchEntire')}</label>
    </div>
  </div>
  {#if results.length > 0}
    <div class="results list" role="listbox" aria-label={t('dlgFindAll')}>
      <div class="head" aria-hidden="true"><span>{t('dlgSheet')}</span><span>{t('dlgCell')}</span><span>{t('dlgValue')}</span></div>
      {#each results as hit, i (`${hit.sheetIndex}:${hit.row}:${hit.col}`)}
        <button
          role="option"
          aria-selected={selected === i}
          onclick={() => {
            selected = i;
            goTo(hit);
          }}><span>{hit.sheet}</span><span>{cellAddress(hit.row, hit.col, true)}</span><span>{hit.text}</span></button
        >
      {/each}
    </div>
    <div class="hint">{t('dlgCellsFound', { n: results.length })}</div>
  {/if}

  {#snippet footer()}
    {#if mode === 'replace'}
      <button class="xl-btn outlined" disabled={!fs.query} onclick={replaceEverything}>{t('dlgReplaceAll')}</button>
      <button class="xl-btn outlined" disabled={!fs.query} onclick={replaceOne}>{t('dlgReplace')}</button>
    {/if}
    <span class="spacer"></span>
    <button class="xl-btn outlined" onclick={() => ctl.closeDialog()}>{t('dlgClose')}</button>
    <button class="xl-btn outlined" disabled={!fs.query} onclick={findAllHits}>{t('dlgFindAll')}</button>
    <button class="xl-btn primary" disabled={!fs.query} onclick={findNextHit}>{t('findNext')}</button>
  {/snippet}
</Dialog>

<style>
  .opts {
    display: flex;
    gap: 24px;
    justify-content: space-between;
  }
  .results {
    max-height: 160px;
    margin-top: 6px;
  }
  .results .head,
  .results button {
    display: grid;
    grid-template-columns: 100px 70px 1fr;
    gap: 6px;
  }
  .results .head {
    padding: 3px 6px;
    font-weight: 600;
    border-bottom: 1px solid var(--xl-border);
    position: sticky;
    top: 0;
    background: #f7f7f7;
  }
  .results span {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .spacer {
    flex: 1;
  }
</style>
