<script lang="ts">
  // Review ▸ Workbook Statistics.
  import { getEditor } from '../core/context.ts';
  import { COUNT_KEYS, workbookStatistics, type SheetCounts } from '../core/workbook-stats.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  const stats = workbookStatistics(ctl.doc.wb, ctl.doc.ws);
  const LABELS: Record<keyof SheetCounts, MessageKey> = {
    cells: 'wsCells',
    tables: 'wsTables',
    pivotTables: 'wsPivotTables',
    formulas: 'wsFormulas',
    charts: 'wsCharts',
    images: 'wsImages',
    comments: 'wsComments',
    notes: 'wsNotes',
  };
</script>

<Dialog title={t('workbookStatistics')} width={360} showCancel={false}>
  <h3>{t('wsCurrentSheet')}: {ctl.doc.ws.title}</h3>
  <dl>
    <dt>{t('wsEndOfSheet')}</dt>
    <dd>{stats.endOfSheet}</dd>
    {#each COUNT_KEYS as k (k)}<dt>{t(LABELS[k])}</dt><dd>{stats.sheet[k].toLocaleString()}</dd>{/each}
  </dl>
  <h3>{t('wsWorkbook')}</h3>
  <dl>
    <dt>{t('wsSheets')}</dt>
    <dd>{stats.sheets}</dd>
    {#each COUNT_KEYS as k (k)}<dt>{t(LABELS[k])}</dt><dd>{stats.workbook[k].toLocaleString()}</dd>{/each}
  </dl>
</Dialog>

<style>
  h3 {
    margin: 4px 0;
    font-size: 12px;
  }
  dl {
    display: grid;
    grid-template-columns: 1fr auto;
    gap: 2px 12px;
    margin: 0 0 10px;
  }
  dt {
    color: var(--xl-text-2);
  }
  dd {
    margin: 0;
    text-align: right;
  }
</style>
