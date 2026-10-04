<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import HomeTab from './HomeTab.svelte';
  import InsertTab from './InsertTab.svelte';
  import PageLayoutTab from './PageLayoutTab.svelte';
  import FormulasTab from './FormulasTab.svelte';
  import DataTab from './DataTab.svelte';
  import ReviewTab from './ReviewTab.svelte';
  import ViewTab from './ViewTab.svelte';
  import ChartDesignTab from './ChartDesignTab.svelte';
  import ChartFormatTab from './ChartFormatTab.svelte';
  import PictureFormatTab from './PictureFormatTab.svelte';
  import ShapeFormatTab from './ShapeFormatTab.svelte';
  import SparklineTab from './SparklineTab.svelte';
  import TableDesignTab from './TableDesignTab.svelte';
  import { tableAt } from '../core/tables.ts';
  import PivotAnalyzeTab from './PivotAnalyzeTab.svelte';
  import PivotDesignTab from './PivotDesignTab.svelte';

  const ctl = getEditor();
  const TABS: Array<{ id: string; label: MessageKey }> = [
    { id: 'home', label: 'tabHome' },
    { id: 'insert', label: 'tabInsert' },
    { id: 'pageLayout', label: 'tabPageLayout' },
    { id: 'formulas', label: 'tabFormulas' },
    { id: 'data', label: 'tabData' },
    { id: 'review', label: 'tabReview' },
    { id: 'view', label: 'tabView' },
  ];

  // Contextual tabs follow the selected drawing item, after View as on Excel for Mac.
  const selectedKind = $derived.by(() => {
    void ctl.doc.version;
    const index = ctl.selectedDrawing;
    return index === null ? undefined : ctl.doc.ws.drawing?.items[index]?.content.kind;
  });
  // The Sparkline tab follows the active cell, like Excel's.
  const activeSparkline = $derived.by(() => {
    const { row, col } = ctl.doc.selection.active;
    return ctl.sparklines.get(`${row}:${col}`);
  });
  // The Table tab follows the active cell too, unless a drawing is selected.
  const activeTable = $derived.by(() => {
    void ctl.doc.version;
    if (selectedKind !== undefined) return -1;
    const { row, col } = ctl.doc.selection.active;
    const hit = tableAt(ctl.doc.ws, row, col);
    return hit ? ctl.doc.ws.tables.indexOf(hit.def) : -1;
  });
  const tableTabs: Array<{ id: string; label: MessageKey }> = $derived(activeTable >= 0 ? [{ id: 'tableDesign', label: 'tabTableDesign' }] : []);
  const contextual: Array<{ id: string; label: MessageKey }> = $derived(
    selectedKind === 'chart'
      ? [
          { id: 'chartDesign', label: 'tabChartDesign' },
          { id: 'chartFormat', label: 'tabChartFormat' },
        ]
      : selectedKind === 'picture'
        ? [{ id: 'pictureFormat', label: 'tabPictureFormat' }]
        : selectedKind === 'shape'
          ? [{ id: 'shapeFormat', label: 'tabShapeFormat' }]
          : activeSparkline
            ? [{ id: 'sparkline', label: 'tabSparkline' }]
            : ctl.activePivot
              ? [
                  { id: 'pivotAnalyze', label: 'tabPivotAnalyze' },
                  { id: 'pivotDesign', label: 'tabPivotDesign' },
                ]
              : [],
  );
  const CONTEXTUAL_IDS = new Set(['chartDesign', 'chartFormat', 'pictureFormat', 'shapeFormat', 'sparkline', 'pivotAnalyze', 'pivotDesign']);
  $effect(() => {
    if (CONTEXTUAL_IDS.has(ctl.ribbonTab) && ![...contextual, ...tableTabs].some((tab) => tab.id === ctl.ribbonTab)) ctl.ribbonTab = 'home';
  });

  function select(id: string) {
    if (ctl.ribbonTab === id && !ctl.ribbonCollapsed) ctl.ribbonCollapsed = true;
    else {
      ctl.ribbonTab = id;
      ctl.ribbonCollapsed = false;
    }
  }
</script>

<div class="ribbon">
  <div class="tabs" role="tablist">
    {#each [...TABS, ...contextual, ...tableTabs] as tab (tab.id)}
      <button
        role="tab"
        class="tab"
        class:active={ctl.ribbonTab === tab.id}
        class:contextual={CONTEXTUAL_IDS.has(tab.id)}
        aria-selected={ctl.ribbonTab === tab.id}
        onclick={() => select(tab.id)}
        ondblclick={() => (ctl.ribbonCollapsed = !ctl.ribbonCollapsed)}
        onmousedown={(e) => e.preventDefault()}>{t(tab.label)}</button
      >
    {/each}
    <span class="spacer"></span>
    <button class="xl-btn collapse" title={ctl.ribbonCollapsed ? t('showRibbon') : t('collapseRibbon')} onclick={() => (ctl.ribbonCollapsed = !ctl.ribbonCollapsed)}>
      <Icon name={ctl.ribbonCollapsed ? 'chevron-down' : 'chevron-up'} size={14} />
    </button>
  </div>
  {#if !ctl.ribbonCollapsed}
    <div class="panel" role="tabpanel">
      {#if ctl.ribbonTab === 'home'}<HomeTab />
      {:else if ctl.ribbonTab === 'insert'}<InsertTab />
      {:else if ctl.ribbonTab === 'pageLayout'}<PageLayoutTab />
      {:else if ctl.ribbonTab === 'formulas'}<FormulasTab />
      {:else if ctl.ribbonTab === 'data'}<DataTab />
      {:else if ctl.ribbonTab === 'review'}<ReviewTab />
      {:else if ctl.ribbonTab === 'view'}<ViewTab />
      {:else if ctl.ribbonTab === 'chartDesign' && ctl.selectedDrawing !== null}<ChartDesignTab index={ctl.selectedDrawing} />
      {:else if ctl.ribbonTab === 'chartFormat' && ctl.selectedDrawing !== null}<ChartFormatTab index={ctl.selectedDrawing} />
      {:else if ctl.ribbonTab === 'pictureFormat' && ctl.selectedDrawing !== null}<PictureFormatTab index={ctl.selectedDrawing} />
      {:else if ctl.ribbonTab === 'shapeFormat' && ctl.selectedDrawing !== null}<ShapeFormatTab index={ctl.selectedDrawing} />
      {:else if ctl.ribbonTab === 'sparkline' && activeSparkline}<SparklineTab groupIndex={activeSparkline.groupIndex} />
      {:else if ctl.ribbonTab === 'tableDesign' && activeTable >= 0}<TableDesignTab tableIndex={activeTable} />
      {:else if ctl.ribbonTab === 'pivotAnalyze' && ctl.activePivot}<PivotAnalyzeTab />
      {:else if ctl.ribbonTab === 'pivotDesign' && ctl.activePivot}<PivotDesignTab />
      {/if}
    </div>
  {/if}
</div>

<style>
  .ribbon {
    background: var(--xl-ribbon);
    border-bottom: 1px solid var(--xl-border);
    flex-shrink: 0;
  }
  .tabs {
    display: flex;
    align-items: flex-end;
    gap: 2px;
    padding: 0 8px;
    height: 30px;
  }
  .tab {
    border: 0;
    background: transparent;
    font: inherit;
    font-size: 12.5px;
    padding: 6px 10px 6px;
    cursor: pointer;
    color: var(--xl-text);
    border-bottom: 3px solid transparent;
  }
  .tab:hover {
    background: var(--xl-hover);
  }
  .tab.active {
    color: var(--xl-accent);
    font-weight: 600;
    border-bottom-color: var(--xl-accent);
  }
  .tab.contextual {
    color: var(--xl-accent);
  }
  .spacer {
    flex: 1;
  }
  .collapse {
    margin-bottom: 3px;
  }
  .panel {
    display: flex;
    align-items: stretch;
    height: 100px;
    padding: 4px 4px 0;
    overflow-x: auto;
    overflow-y: hidden;
    /* Narrow windows scroll the ribbon sideways (trackpad/shift-wheel) without a bar eating its height. */
    scrollbar-width: none;
    background: #fff;
    border-top: 1px solid var(--xl-border);
  }
</style>
