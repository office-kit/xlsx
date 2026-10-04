<script lang="ts">
  import * as A from '../core/actions.ts';
  import { newComment } from '../core/comments.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import Group from './Group.svelte';
  import ShapesButton from './ShapesButton.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import ChartMenuButton from './ChartMenuButton.svelte';
  import { INSERT_MENUS } from './chart-ui.ts';
  import { chartRangeOf, insertChart, type ChartChoice } from '../core/charts.ts';

  const ctl = getEditor();

  // Picking from a chart gallery inserts straight away over the selection, as in Excel.
  function insertChartNow(choice: ChartChoice): void {
    const range = chartRangeOf(ctl.doc);
    if (!range) {
      ctl.openDialog('insertChart', { type: choice });
      return;
    }
    if (ctl.edit && !ctl.commitEdit()) return;
    const index = insertChart(ctl.doc, choice, range);
    if (index === undefined) return;
    ctl.selectedDrawing = index;
    ctl.ribbonTab = 'chartDesign';
  }
  let pictureInput: HTMLInputElement;

  const SPARKLINE_MENU = [
    { type: 'line', label: 'spkLine' },
    { type: 'column', label: 'spkColumn' },
    { type: 'stacked', label: 'spkWinLoss' },
  ] as const;

  function armTextBox(): void {
    if (ctl.edit && !ctl.commitEdit()) return;
    ctl.selectedDrawing = null;
    ctl.shapeTool = { kind: 'textBox' };
  }
</script>

<Group label={t('groupTables')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('createPivotTable')}>
    <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.4" /><path d="M3 8h18M8 3v18" stroke="currentColor" stroke-width="1.4" /><path d="M12 12h5v5" fill="none" stroke="var(--xl-accent)" stroke-width="1.6" /><path d="M15 15l2 2 2-2" fill="none" stroke="var(--xl-accent)" stroke-width="1.4" /></svg>
    <span>{t('pivotTable')}</span>
  </button>
  <button class="xl-btn big" onclick={() => ctl.openDialog('createTable')}><Icon name="table" size={24} /><span>{t('table')}</span></button>
</Group>

<Group label={t('groupIllustrations')}>
  <button class="xl-btn big" onclick={() => pictureInput.click()}><Icon name="image" size={24} /><span>{t('pictures')}</span></button>
  <input bind:this={pictureInput} type="file" accept="image/png,image/jpeg,image/gif" hidden onchange={(e) => { const f = (e.currentTarget as HTMLInputElement).files?.[0]; if (f) void A.insertPicture(ctl, f); (e.currentTarget as HTMLInputElement).value = ''; }} />
  <ShapesButton />
  <button class="xl-btn big" title={t('shpTextBox')} onclick={armTextBox} onmousedown={(e) => e.preventDefault()}>
    <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" fill="none" stroke="currentColor" stroke-width="1.4" /><path d="M7 8h10M12 8v9" stroke="currentColor" stroke-width="1.6" /></svg>
    <span>{t('shpTextBox')}</span>
  </button>
</Group>

<Group label={t('groupCharts')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('insertChart', {})}><Icon name="chart" size={24} /><span>{t('recommendedCharts')}</span></button>
  {#each INSERT_MENUS as m (m.label)}
    <ChartMenuButton label={t(m.label)} title={t(m.tip)} groups={m.groups} onpick={insertChartNow} onmore={(c) => ctl.openDialog('insertChart', { type: c })} />
  {/each}
</Group>

<Group label={t('spkSparklines')}>
  <MenuButton large icon="chart-line" label={t('spkSparklines')} title={t('spkSparklines')}>
    {#snippet menu(close)}
      {#each SPARKLINE_MENU as m (m.type)}
        <button class="xl-menu-item" onclick={() => { close(); ctl.openDialog('createSparklines', { type: m.type }); }}>{t(m.label)}</button>
      {/each}
    {/snippet}
  </MenuButton>
</Group>

<Group label={t('groupLinks')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('hyperlink')}><Icon name="link" size={24} /><span>{t('link')}</span></button>
</Group>

<Group label={t('groupComments')}>
  <button class="xl-btn big" onclick={() => newComment(ctl)}><Icon name="note" size={24} /><span>{t('cmtNewComment')}</span></button>
  <button class="xl-btn big" onclick={() => ctl.openDialog('note')}><Icon name="note" size={24} /><span>{t('newNote')}</span></button>
</Group>

<Group label={t('dtGroupText')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('pageSetup', { tab: 'headerFooter' })}><Icon name="page-layout" size={24} /><span>{t('dtHeaderFooter')}</span></button>
</Group>

<Group label={t('groupSymbols')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('symbol')}><Icon name="symbol" size={24} /><span>{t('symbol')}</span></button>
</Group>

<style>
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
</style>
