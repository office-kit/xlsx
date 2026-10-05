<script lang="ts">
  import * as A from '../core/actions.ts';
  import { newComment } from '../core/comments.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import Group from './Group.svelte';
  import ShapesButton from './ShapesButton.svelte';
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

  const SPARKLINES = [
    { type: 'line', label: 'spkLine', icon: 'sparkline-line' },
    { type: 'column', label: 'spkColumn', icon: 'sparkline-column' },
    { type: 'stacked', label: 'spkWinLoss', icon: 'sparkline-winloss' },
  ] as const;

  function armTextBox(): void {
    if (ctl.edit && !ctl.commitEdit()) return;
    ctl.selectedDrawing = null;
    ctl.shapeTool = { kind: 'textBox' };
  }
</script>

<Group label={t('groupTables')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('createPivotTable')}>
    <svg width="30" height="30" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.4" /><path d="M3 8h18M8 3v18" stroke="currentColor" stroke-width="1.4" /><path d="M12 12h5v5" fill="none" stroke="var(--xl-accent)" stroke-width="1.6" /><path d="M15 15l2 2 2-2" fill="none" stroke="var(--xl-accent)" stroke-width="1.4" /></svg>
    <span>{t('pivotTable')}</span>
  </button>
  <button class="xl-btn big" onclick={() => ctl.openDialog('createTable')}><Icon name="table" size={30} /><span>{t('table')}</span></button>
</Group>

<Group label={t('groupIllustrations')}>
  <button class="xl-btn big" onclick={() => pictureInput.click()}><Icon name="image" size={30} /><span>{t('pictures')}</span></button>
  <input bind:this={pictureInput} type="file" accept="image/png,image/jpeg,image/gif" hidden onchange={(e) => { const f = (e.currentTarget as HTMLInputElement).files?.[0]; if (f) void A.insertPicture(ctl, f); (e.currentTarget as HTMLInputElement).value = ''; }} />
  <ShapesButton />
</Group>

<Group label={t('groupCharts')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('insertChart', {})}><Icon name="recommended-charts" size={30} /><span>{t('recommendedCharts')}</span></button>
  <div class="chart-grid">
    {#each INSERT_MENUS as m (m.label)}
      <span style:grid-column={m.col} style:grid-row={m.row}>
        <ChartMenuButton label={t(m.label)} title={t(m.tip)} icon={m.icon} groups={m.groups} onpick={insertChartNow} onmore={(c) => ctl.openDialog('insertChart', { type: c })} />
      </span>
    {/each}
  </div>
</Group>

<Group label={t('spkSparklines')}>
  {#each SPARKLINES as m (m.type)}
    <button class="xl-btn big" onclick={() => ctl.openDialog('createSparklines', { type: m.type })}><Icon name={m.icon} size={30} /><span>{t(m.label)}</span></button>
  {/each}
</Group>

<Group label={t('groupLinks')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('hyperlink')}><Icon name="link" size={30} /><span>{t('link')}</span></button>
</Group>

<Group label={t('groupComments')}>
  <button class="xl-btn big" onclick={() => newComment(ctl)}><Icon name="comment-new" size={30} /><span>{t('insComment')}</span></button>
</Group>

<Group label={t('dtGroupText')}>
  <button class="xl-btn big" title={t('shpTextBox')} onclick={armTextBox} onmousedown={(e) => e.preventDefault()}>
    <svg width="30" height="30" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" fill="none" stroke="currentColor" stroke-width="1.4" /><path d="M7 8h10M12 8v9" stroke="currentColor" stroke-width="1.6" /></svg>
    <span>{t('shpTextBox')}</span>
  </button>
  <button class="xl-btn big" onclick={() => ctl.openDialog('pageSetup', { tab: 'headerFooter' })}><Icon name="header-footer" size={30} /><span>{t('dtHeaderFooter')}</span></button>
</Group>

<Group label={t('groupSymbols')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('symbol')}><Icon name="symbol" size={30} /><span>{t('symbol')}</span></button>
</Group>

<style>
  .chart-grid {
    display: grid;
    grid-template-rows: repeat(3, 22px);
    gap: 1px 2px;
    align-self: flex-start;
  }
</style>
