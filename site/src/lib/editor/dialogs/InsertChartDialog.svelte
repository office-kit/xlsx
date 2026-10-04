<script lang="ts">
  // Insert Chart (Recommended Charts / More Charts…) and Change Chart Type:
  // a list of chart groups, the group's types, and a live preview of the
  // selected type over the data. Inserting places the chart beside the data;
  // changing rebuilds the selected chart over its own source block and keeps
  // its formatting. Either way it is one undo step.
  import { untrack } from 'svelte';
  import type { Worksheet } from '@office-kit/xlsx/worksheet';
  import { parseRangeAddress, rangeAddress, type Range } from '../core/address.ts';
  import { chartData } from '../core/chart-data.ts';
  import { editChart } from '../core/chart-edit.ts';
  import { analyzeSource, buildChart, CHART_CHOICES, CHART_GROUPS, chartChoiceOf, chartGroupOf, chartRangeOf, chartSource, insertChart, rebuildChart, type ChartChoice, type ChartGroup, type ChartSource } from '../core/charts.ts';
  import { getEditor } from '../core/context.ts';
  import { usedRange } from '../core/cells.ts';
  import { currentRange } from '../core/selection.ts';
  import { renderChartSvg } from '../grid/chart-render.ts';
  import { CHART_LABELS, GROUP_LABELS, OFFERED_GROUPS, shrinkSvg } from '../ribbon/chart-ui.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();

  const ctl = getEditor();
  const doc = ctl.doc;

  // DialogHost remounts the dialog for every open, so the props are read once.
  const changeIndex = untrack(() => (props?.['mode'] === 'change' ? ctl.selectedDrawing : null));
  const target = changeIndex === null ? undefined : doc.ws.drawing?.items[changeIndex];
  const targetChart = target?.content.kind === 'chart' ? target.content.chart : undefined;
  const source: ChartSource | undefined = targetChart ? chartSource(targetChart) : undefined;
  const changing = changeIndex !== null;

  const isChoice = (v: unknown): v is ChartChoice => CHART_CHOICES.some((c) => c === v);
  const requested = untrack(() => props?.['type']);
  const initialChoice: ChartChoice = (targetChart && chartChoiceOf(targetChart)) ?? (isChoice(requested) ? requested : 'columnClustered');

  let choice = $state<ChartChoice>(initialChoice);
  let group = $state<ChartGroup>(chartGroupOf(initialChoice));
  const sheet = source?.sheet ?? doc.ws.title;
  const initialRange = source?.range ?? chartRangeOf(doc) ?? currentRange(doc.selection);
  let rangeText = $state(rangeAddress(initialRange, true));
  let notice = $state<string | null>(changing && !source ? t('chNoSource') : null);

  const range = $derived.by((): Range | undefined => {
    if (source) return source.range;
    const parsed = parseRangeAddress(rangeText.replace(/^=/, ''))?.range;
    const used = usedRange(doc.ws);
    if (!parsed || !used) return undefined;
    const r = { r1: parsed.r1, c1: parsed.c1, r2: Math.min(parsed.r2, used.r2), c2: Math.min(parsed.c2, used.c2) };
    return r.r1 <= r.r2 && r.c1 <= r.c2 ? r : undefined;
  });
  const valueAt = (r: number, c: number) => doc.calc.cellValue(sheet, r, c);
  let byColumns = $state(source?.byColumns ?? analyzeSource(initialRange, valueAt).byColumns);

  const worksheets = new Map<string, Worksheet>();
  for (const s of doc.wb.sheets) if (s.kind === 'worksheet') worksheets.set(s.sheet.title, s.sheet);

  function preview(c: ChartChoice, w: number, h: number, withText: boolean): string {
    const r = range;
    if (!r) return '';
    const chart = targetChart && source ? rebuildChart(targetChart, c, { ...source, byColumns }, valueAt) : buildChart(c, sheet, r, valueAt, byColumns);
    const data = chartData(chart, doc.calc, worksheets, doc.styles.palette);
    if (!data) return '';
    const view = withText ? data : { ...data, title: undefined, legend: undefined, catAxis: false, valAxis: false, catTitle: undefined, valTitle: undefined };
    return withText ? renderChartSvg(view, w, h, doc.styles.palette) : shrinkSvg(renderChartSvg(view, w * 2, h * 2, doc.styles.palette), w, h);
  }

  function ok(): boolean {
    const r = range;
    if (changing) {
      if (!source || changeIndex === null) return false;
      editChart(doc, changeIndex, 'Change Chart Type', (chart) => {
        const next = rebuildChart(chart, choice, source, valueAt);
        // Replace both parts: a classic chart may become chartex and back.
        if (next.space) chart.space = next.space;
        else delete chart.space;
        if (next.cxSpace) chart.cxSpace = next.cxSpace;
        else delete chart.cxSpace;
      });
      return true;
    }
    if (!r) {
      notice = parseRangeAddress(rangeText.replace(/^=/, '')) ? t('chNoData') : t('cfInvalidRange', { ref: rangeText });
      return false;
    }
    const index = insertChart(doc, choice, r, byColumns);
    if (index === undefined) return true;
    ctl.selectedDrawing = index;
    ctl.ribbonTab = 'chartDesign';
    return true;
  }
</script>

<Dialog title={changing ? t('chChangeType') : t('chInsertChart')} width={720} onok={ok} okDisabled={changing && !source} bind:notice>
  <div class="layout">
    <div class="groups" role="listbox" aria-label={t('chInsertChart')}>
      {#each OFFERED_GROUPS as g (g)}
        <button class="group" role="option" aria-selected={group === g} onclick={() => { group = g; choice = CHART_GROUPS[g][0]; }}>{t(GROUP_LABELS[g])}</button>
      {/each}
    </div>
    <div class="main">
      <div class="types" role="listbox" aria-label={t(GROUP_LABELS[group])}>
        {#each CHART_GROUPS[group] as c (c)}
          <button class="thumb" role="option" aria-selected={choice === c} title={t(CHART_LABELS[c])} onclick={() => (choice = c)} ondblclick={() => { choice = c; if (ok()) ctl.closeDialog(); }}>
            <span class="svg">{@html preview(c, 64, 44, false)}</span>
          </button>
        {/each}
      </div>
      <div class="name">{t(CHART_LABELS[choice])}</div>
      <div class="big">{@html preview(choice, 420, 240, true)}</div>
      {#if !changing}
        <div class="row">
          <label for="ch-range">{t('chDataRange')}</label>
          <input id="ch-range" class="xl-input grow" bind:value={rangeText} spellcheck="false" />
        </div>
      {/if}
      <div class="row">
        <span class="lbl">{t('chSeriesIn')}</span>
        <label class="check"><input type="radio" name="ch-orient" checked={byColumns} disabled={changing} onchange={() => (byColumns = true)} />{t('chColumns')}</label>
        <label class="check"><input type="radio" name="ch-orient" checked={!byColumns} disabled={changing} onchange={() => (byColumns = false)} />{t('chRows')}</label>
      </div>
    </div>
  </div>
</Dialog>

<style>
  .layout {
    display: flex;
    gap: 12px;
  }
  .groups {
    display: flex;
    flex-direction: column;
    width: 170px;
    flex-shrink: 0;
    border-right: 1px solid var(--xl-border);
    padding-right: 6px;
  }
  .group {
    text-align: left;
    border: 0;
    border-radius: 4px;
    background: transparent;
    font: inherit;
    padding: 4px 8px;
    color: var(--xl-text);
    cursor: pointer;
  }
  .group:hover {
    background: var(--xl-hover);
  }
  .group[aria-selected='true'] {
    background: var(--xl-selected, #dbe8f7);
    font-weight: 600;
  }
  .main {
    flex: 1;
    min-width: 0;
  }
  .types {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
  }
  .thumb {
    padding: 2px;
    border: 1px solid var(--xl-border);
    border-radius: 4px;
    background: #fff;
    cursor: pointer;
  }
  .thumb[aria-selected='true'] {
    border-color: var(--xl-accent);
    box-shadow: 0 0 0 1px var(--xl-accent);
  }
  .svg {
    display: block;
    width: 64px;
    height: 44px;
    pointer-events: none;
  }
  .name {
    margin: 8px 0 4px;
    font-weight: 600;
  }
  .big {
    width: 420px;
    height: 240px;
    margin-bottom: 8px;
  }
</style>
