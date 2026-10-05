<script lang="ts">
  // Chart Design ▸ Select Data (Excel's "Select Data Source"). Changing the
  // chart data range rebuilds the chart over the new block, keeping its type
  // and formatting; otherwise the series list is applied as edited: names,
  // value / X / size refs, added and removed series, and the category labels.
  // Chartex charts only take a data range.
  import type { Worksheet } from '@office-kit/xlsx/worksheet';
  import { parseRangeAddress, quoteSheetName, rangeAddress } from '../core/address.ts';
  import { resolveRef } from '../core/chart-data.ts';
  import { editChart, isXYChart, seriesRefs, setSeriesRefs, type SeriesRefs } from '../core/chart-edit.ts';
  import { chartChoiceOf, chartSource, rebuildChart } from '../core/charts.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  const index = ctl.selectedDrawing;
  const item = index === null ? undefined : doc.ws.drawing?.items[index];
  const chart = item?.content.kind === 'chart' ? item.content.chart : undefined;
  const source = chart ? chartSource(chart) : undefined;
  const classic = chart?.space !== undefined;
  const xy = chart ? isXYChart(chart) : false;
  const bubble = chart?.space?.plotArea.chart.kind === 'bubble';
  const initial = chart ? seriesRefs(chart) : { series: [], categories: '' };

  const initialRange = source ? `=${quoteSheetName(source.sheet)}!${rangeAddress(source.range, true)}` : '';
  let rangeText = $state(initialRange);
  let byColumns = $state(source?.byColumns ?? true);
  let rows = $state<SeriesRefs[]>(initial.series.map((s) => ({ ...s })));
  let categories = $state(initial.categories);
  let current = $state(0);
  let notice = $state<string | null>(null);
  const row = $derived(rows[current]);

  const worksheets = new Map<string, Worksheet>();
  for (const s of doc.wb.sheets) if (s.kind === 'worksheet') worksheets.set(s.sheet.title, s.sheet);

  /** A ref as typed (with or without `=`, with or without a sheet) qualified with the chart's sheet. */
  function qualify(text: string): string | undefined {
    const raw = text.trim().replace(/^=/, '');
    const parsed = parseRangeAddress(raw);
    if (!parsed) return undefined;
    const sheet = parsed.sheet ?? source?.sheet ?? doc.ws.title;
    return `${quoteSheetName(sheet)}!${rangeAddress(parsed.range, true)}`;
  }

  function seriesLabel(s: SeriesRefs, i: number): string {
    if (!s.name.startsWith('=')) return s.name || t('chNewSeries', { n: i + 1 });
    const values = resolveRef(s.name.slice(1), doc.calc, worksheets);
    return values?.map((v) => (v === null || typeof v === 'object' ? '' : String(v))).join(' ') || s.name;
  }

  function addSeries(): void {
    rows = [...rows, { name: '', values: '', x: '', size: '' }];
    current = rows.length - 1;
  }

  function removeSeries(): void {
    rows = rows.filter((_, i) => i !== current);
    current = Math.max(0, Math.min(current, rows.length - 1));
  }

  function rebuildOver(text: string, columns: boolean): boolean {
    const ref = qualify(text);
    const parsed = ref ? parseRangeAddress(ref) : undefined;
    const choice = chart && chartChoiceOf(chart);
    if (!parsed?.sheet || !choice || index === null) {
      notice = t('chInvalidRef', { ref: text });
      return false;
    }
    const sheet = parsed.sheet;
    const src = { sheet, range: parsed.range, byColumns: columns };
    editChart(doc, index, 'Select Data', (live) => {
      const next = rebuildChart(live, choice, src, (r, c) => doc.calc.cellValue(sheet, r, c));
      if (next.space) live.space = next.space;
      if (next.cxSpace) live.cxSpace = next.cxSpace;
    });
    return true;
  }

  function ok(): boolean {
    if (index === null || !chart) return true;
    if (rangeText.trim() !== initialRange || byColumns !== (source?.byColumns ?? true)) return rebuildOver(rangeText, byColumns);
    const edited = JSON.stringify(rows) !== JSON.stringify(initial.series) || categories !== initial.categories;
    if (!classic || !edited) return true;
    if (rows.length === 0) {
      notice = t('chNeedSeries');
      return false;
    }
    const fixed: SeriesRefs[] = [];
    for (const r of rows) {
      const values = qualify(r.values);
      const x = r.x.trim() ? qualify(r.x) : '';
      const size = r.size.trim() ? qualify(r.size) : '';
      const nameRef = r.name.startsWith('=') ? qualify(r.name) : undefined;
      const bad = !values ? r.values : x === undefined ? r.x : size === undefined ? r.size : r.name.startsWith('=') && !nameRef ? r.name : undefined;
      if (bad !== undefined || !values) {
        notice = t('chInvalidRef', { ref: bad ?? r.values });
        return false;
      }
      fixed.push({ name: nameRef ? `=${nameRef}` : r.name, values, x: x ?? '', size: size ?? '' });
    }
    const cat = categories.trim() ? qualify(categories) : '';
    if (cat === undefined) {
      notice = t('chInvalidRef', { ref: categories });
      return false;
    }
    editChart(doc, index, 'Select Data', (live) => setSeriesRefs(live, fixed, cat, (ref) => resolveRef(ref, doc.calc, worksheets)));
    return true;
  }
</script>

<Dialog title={t('chSelectDataTitle')} width={560} onok={ok} bind:notice>
  <div class="row">
    <label for="sd-range">{t('chChartDataRange')}</label>
    <input id="sd-range" class="xl-input grow" bind:value={rangeText} spellcheck="false" />
    <button class="xl-btn outlined" onclick={() => (byColumns = !byColumns)}>{t('chSwitchRowCol')}</button>
  </div>
  {#if classic}
    <div class="cols">
      <fieldset>
        <legend>{t('chLegendSeries')}</legend>
        <div class="list" role="listbox" aria-label={t('chLegendSeries')}>
          {#each rows as s, i (i)}
            <button class="entry" role="option" aria-selected={current === i} onclick={() => (current = i)}>{seriesLabel(s, i)}</button>
          {/each}
        </div>
        <div class="buttons">
          <button class="xl-btn outlined" onclick={addSeries}>{t('chAddSeries')}</button>
          <button class="xl-btn outlined" disabled={rows.length <= 1} onclick={removeSeries}>{t('chRemoveSeries')}</button>
        </div>
      </fieldset>
      <fieldset>
        {#if row}
          <label class="field">{t('chSeriesName')}<input class="xl-input" bind:value={row.name} spellcheck="false" /></label>
          {#if xy}
            <label class="field">{t('chXValues')}<input class="xl-input" bind:value={row.x} spellcheck="false" /></label>
          {/if}
          <label class="field">{t('chSeriesValues')}<input class="xl-input" bind:value={row.values} spellcheck="false" /></label>
          {#if bubble}
            <label class="field">{t('chBubbleSizes')}<input class="xl-input" bind:value={row.size} spellcheck="false" /></label>
          {/if}
        {/if}
      </fieldset>
    </div>
    {#if !xy}
      <div class="row">
        <label for="sd-cat">{t('chCategoryLabels')}</label>
        <input id="sd-cat" class="xl-input grow" bind:value={categories} spellcheck="false" />
      </div>
    {/if}
  {/if}
</Dialog>

<style>
  .cols {
    display: flex;
    gap: 10px;
    margin: 8px 0;
  }
  .cols fieldset {
    flex: 1;
    min-width: 0;
  }
  .list {
    display: flex;
    flex-direction: column;
    height: 120px;
    overflow: auto;
    border: 1px solid var(--xl-border);
    background: #fff;
  }
  .entry {
    text-align: left;
    border: 0;
    background: transparent;
    font: inherit;
    padding: 2px 6px;
    cursor: pointer;
  }
  .entry[aria-selected='true'] {
    background: var(--xl-accent);
    color: #fff;
  }
  .buttons {
    display: flex;
    gap: 6px;
    margin-top: 6px;
  }
  .field {
    display: flex;
    flex-direction: column;
    gap: 2px;
    margin-bottom: 6px;
  }
</style>
