<script lang="ts">
  // Chart Design, the contextual tab of a selected chart: chart elements,
  // Quick Layout, colours and styles (galleries preview the selected chart
  // with the change applied), data and type changes, and Move Chart. Each
  // command is one undo step on the chart.
  import type { ChartReference } from '@office-kit/xlsx/drawing';
  import type { Worksheet } from '@office-kit/xlsx/worksheet';
  import { chartData, dmlColor } from '../core/chart-data.ts';
  import {
    applyChartStyle,
    applyColorSet,
    applyQuickLayout,
    chartElements,
    CHART_STYLES,
    COLOR_SETS,
    editChart,
    QUICK_LAYOUTS,
    setAxisTitle,
    setAxisVisible,
    setChartTitle,
    setDataLabels,
    setGridline,
    setLegend,
    type Axis,
    type Gridline,
    type LabelChoice,
    type LegendChoice,
    type TitleMode,
  } from '../core/chart-edit.ts';
  import { chartChoiceOf, chartSource, rebuildChart } from '../core/charts.ts';
  import { getEditor } from '../core/context.ts';
  import { renderChartSvg } from '../grid/chart-render.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';
  import { shrinkSvg } from './chart-ui.ts';

  let { index }: { index: number } = $props();

  const ctl = getEditor();
  const doc = ctl.doc;

  const chart = $derived.by((): ChartReference | undefined => {
    void doc.version;
    const item = doc.ws.drawing?.items[index];
    return item?.content.kind === 'chart' ? item.content.chart : undefined;
  });
  const el = $derived(chart ? chartElements(chart) : undefined);
  const pie = $derived(['pie', 'pie3D', 'doughnut', 'ofPie'].includes(chart?.space?.plotArea.chart.kind ?? ''));
  const source = $derived(chart ? chartSource(chart) : undefined);

  function edit(label: string, fn: (c: ChartReference) => void): void {
    editChart(doc, index, label, fn);
  }

  function preview(fn: (c: ChartReference) => void, w: number, h: number): string {
    if (!chart) return '';
    const copy = structuredClone(chart);
    fn(copy);
    const worksheets = new Map<string, Worksheet>();
    for (const s of doc.wb.sheets) if (s.kind === 'worksheet') worksheets.set(s.sheet.title, s.sheet);
    const data = chartData(copy, doc.calc, worksheets, doc.styles.palette);
    // Drawn at 2.5× and scaled down, so the whole layout fits the tile.
    return data ? shrinkSvg(renderChartSvg(data, w * 2.5, h * 2.5, doc.styles.palette), w, h) : '';
  }

  function switchRowColumn(): void {
    const c = chart;
    const src = source;
    const choice = c && chartChoiceOf(c);
    if (!src || !choice) {
      ctl.openDialog('alert', { message: 'chNoSource' });
      return;
    }
    edit('Switch Row/Column', (live) => {
      const next = rebuildChart(live, choice, { ...src, byColumns: !src.byColumns }, (r, col) => doc.calc.cellValue(src.sheet, r, col));
      if (next.space) live.space = next.space;
      if (next.cxSpace) live.cxSpace = next.cxSpace;
    });
  }

  const TITLES: ReadonlyArray<[TitleMode, MessageKey]> = [
    ['none', 'chNone'],
    ['above', 'chTitleAbove'],
    ['overlay', 'chTitleOverlay'],
  ];
  const LEGENDS: ReadonlyArray<[LegendChoice, MessageKey]> = [
    ['none', 'chNone'],
    ['r', 'chLegendRight'],
    ['t', 'chLegendTop'],
    ['l', 'chLegendLeft'],
    ['b', 'chLegendBottom'],
  ];
  const LABELS: ReadonlyArray<[LabelChoice, MessageKey]> = [
    ['none', 'chNone'],
    ['center', 'chLabelCenter'],
    ['insideEnd', 'chLabelInsideEnd'],
    ['insideBase', 'chLabelInsideBase'],
    ['outsideEnd', 'chLabelOutsideEnd'],
  ];
  const PIE_LABELS: ReadonlyArray<[LabelChoice, MessageKey]> = [
    ['none', 'chNone'],
    ['center', 'chLabelCenter'],
    ['insideEnd', 'chLabelInsideEnd'],
    ['outsideEnd', 'chLabelOutsideEnd'],
    ['bestFit', 'chLabelBestFit'],
  ];
  const AXES: ReadonlyArray<[Axis, MessageKey]> = [
    ['cat', 'chPrimaryHorizontal'],
    ['val', 'chPrimaryVertical'],
  ];
  const GRIDS: ReadonlyArray<[Gridline, MessageKey]> = [
    ['valMajor', 'chGridValMajor'],
    ['catMajor', 'chGridCatMajor'],
    ['valMinor', 'chGridValMinor'],
  ];

  const swatch = (c: (typeof COLOR_SETS)[number]['colors'][number]) => dmlColor(c, doc.styles.palette) ?? '#000000';
</script>

{#snippet item(label: string, checked: boolean, run: () => void, close: () => void)}
  <button class="xl-menu-item" role="menuitemcheckbox" aria-checked={checked} onclick={() => { run(); close(); }}>
    <span class="tick">{#if checked}<Icon name="check" size={12} />{/if}</span>{label}
  </button>
{/snippet}

{#if chart && el}
  <Group label={t('chGroupLayouts')}>
    <MenuButton large icon="chart" label={t('chAddElement')} title={t('chAddElement')}>
      {#snippet menu(close)}
        <div class="elements">
          {#if el.hasAxes}
            <div class="head">{t('chElAxes')}</div>
            {#each AXES as [a, key] (a)}
              {@render item(t(key), el.axes[a], () => edit('Axes', (c) => setAxisVisible(c, a, !el.axes[a])), close)}
            {/each}
            <div class="head">{t('chElAxisTitles')}</div>
            {@render item(t('chPrimaryHorizontal'), el.catTitle, () => edit('Axis Titles', (c) => setAxisTitle(c, 'cat', !el.catTitle)), close)}
            {@render item(t('chPrimaryVertical'), el.valTitle, () => edit('Axis Titles', (c) => setAxisTitle(c, 'val', !el.valTitle)), close)}
          {/if}
          <div class="head">{t('chElChartTitle')}</div>
          {#each TITLES as [mode, key] (mode)}
            {@render item(t(key), el.title === mode, () => edit('Chart Title', (c) => setChartTitle(c, mode, mode === 'none' ? undefined : el.titleText || t('chDefaultTitle'))), close)}
          {/each}
          <div class="head">{t('chElDataLabels')}</div>
          {#each pie ? PIE_LABELS : LABELS as [pos, key] (pos)}
            {@render item(t(key), el.labels === pos, () => edit('Data Labels', (c) => setDataLabels(c, pos)), close)}
          {/each}
          {#if el.hasAxes}
            <div class="head">{t('chElGridlines')}</div>
            {#each GRIDS as [g, key] (g)}
              {@render item(t(key), el.gridlines[g], () => edit('Gridlines', (c) => setGridline(c, g, !el.gridlines[g])), close)}
            {/each}
          {/if}
          <div class="head">{t('chElLegend')}</div>
          {#each LEGENDS as [pos, key] (pos)}
            {@render item(t(key), el.legend === pos, () => edit('Legend', (c) => setLegend(c, pos)), close)}
          {/each}
        </div>
      {/snippet}
    </MenuButton>
    <MenuButton large icon="grid" label={t('chQuickLayout')} title={t('chQuickLayout')}>
      {#snippet menu(close)}
        <div class="gallery">
          {#each QUICK_LAYOUTS as n (n)}
            <button class="tile" title={t('chLayoutN', { n })} onclick={() => { edit('Quick Layout', (c) => applyQuickLayout(c, n)); close(); }}>{@html preview((c) => applyQuickLayout(c, n), 96, 64)}</button>
          {/each}
        </div>
      {/snippet}
    </MenuButton>
  </Group>

  <Group label={t('chGroupStyles')}>
    <MenuButton large icon="fill" label={t('chChangeColors')} title={t('chChangeColors')}>
      {#snippet menu(close)}
        <div class="colors">
          <div class="head">{t('chColorful')}</div>
          {#each COLOR_SETS.slice(0, 4) as set, i (set.id)}
            <button class="xl-menu-item set" title={t('chColorfulN', { n: i + 1 })} onclick={() => { edit('Change Colors', (c) => applyColorSet(c, set)); close(); }}>
              {#each set.colors as col, j (j)}<span class="sw" style:background={swatch(col)}></span>{/each}
            </button>
          {/each}
          <div class="head">{t('chMonochromatic')}</div>
          {#each COLOR_SETS.slice(4) as set, i (set.id)}
            <button class="xl-menu-item set" title={t('chMonoN', { n: i + 1 })} onclick={() => { edit('Change Colors', (c) => applyColorSet(c, set)); close(); }}>
              {#each set.colors as col, j (j)}<span class="sw" style:background={swatch(col)}></span>{/each}
            </button>
          {/each}
        </div>
      {/snippet}
    </MenuButton>
    <div class="styles">
      {#each CHART_STYLES as n (n)}
        <button class="style" title={t('chStyleN', { n })} onclick={() => edit('Chart Style', (c) => applyChartStyle(c, n))}>{@html preview((c) => applyChartStyle(c, n), 64, 44)}</button>
      {/each}
    </div>
  </Group>

  <Group label={t('chGroupData')}>
    <button class="xl-btn big" disabled={!source} onclick={switchRowColumn}><Icon name="text-columns" size={24} /><span>{t('chSwitchRowCol')}</span></button>
    <button class="xl-btn big" onclick={() => ctl.openDialog('chartSelectData')}><Icon name="table" size={24} /><span>{t('chSelectData')}</span></button>
  </Group>

  <Group label={t('chGroupType')}>
    <button class="xl-btn big" onclick={() => ctl.openDialog('insertChart', { mode: 'change' })}><Icon name="chart" size={24} /><span>{t('chChangeType')}</span></button>
  </Group>

  <Group label={t('chGroupLocation')}>
    <button class="xl-btn big" onclick={() => ctl.openDialog('moveChart')}><Icon name="sheet" size={24} /><span>{t('chMoveChart')}</span></button>
  </Group>
{/if}

<style>
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
  .head {
    padding: 6px 10px 2px;
    font-size: 11.5px;
    font-weight: 600;
    color: var(--xl-text-2);
  }
  .tick {
    display: inline-flex;
    width: 14px;
  }
  .gallery {
    display: grid;
    grid-template-columns: repeat(3, auto);
    gap: 4px;
    padding: 4px 8px;
  }
  .tile,
  .style {
    padding: 1px;
    border: 1px solid var(--xl-border);
    border-radius: 3px;
    background: #fff;
    cursor: pointer;
  }
  .tile:hover,
  .style:hover {
    border-color: var(--xl-accent);
  }
  .tile :global(svg),
  .style :global(svg) {
    display: block;
    pointer-events: none;
  }
  .styles {
    display: grid;
    grid-template-columns: repeat(3, auto);
    gap: 2px;
    align-content: center;
    height: 100%;
  }
  .set {
    gap: 2px;
  }
  .sw {
    display: inline-block;
    width: 18px;
    height: 12px;
  }
</style>
