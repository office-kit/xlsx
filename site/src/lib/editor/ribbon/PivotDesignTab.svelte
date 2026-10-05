<script lang="ts">
  // Design, the PivotTable's second contextual tab: subtotals, grand totals
  // and report layout. Each choice is one undo step.
  import type { PivotLayout, PivotSubtotals, PivotTable } from '@office-kit/xlsx/worksheet';
  import { getEditor } from '../core/context.ts';
  import { editPivot } from '../core/pivot.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';

  const ctl = getEditor();
  const model = $derived(ctl.activePivot?.kind === 'model' ? ctl.activePivot : undefined);

  const SUBTOTALS: Array<{ id: PivotSubtotals; label: MessageKey }> = [
    { id: 'off', label: 'pvSubOff' },
    { id: 'bottom', label: 'pvSubBottom' },
    { id: 'top', label: 'pvSubTop' },
  ];
  const GRAND: Array<{ rows: boolean; cols: boolean; label: MessageKey }> = [
    { rows: false, cols: false, label: 'pvGTOff' },
    { rows: true, cols: true, label: 'pvGTOn' },
    { rows: true, cols: false, label: 'pvGTRows' },
    { rows: false, cols: true, label: 'pvGTColumns' },
  ];
  const LAYOUTS: Array<{ id: PivotLayout; label: MessageKey }> = [
    { id: 'compact', label: 'pvCompact' },
    { id: 'outline', label: 'pvOutline' },
    { id: 'tabular', label: 'pvTabular' },
  ];

  function apply(label: string, fn: (pt: PivotTable) => void): void {
    const m = model;
    if (m) editPivot(ctl, m.ws, m.index, label, fn);
  }
</script>

<Group label={t('pvGroupLayout')}>
  <MenuButton large icon="group" label={t('pvSubtotals')} title={t('pvSubtotals')} disabled={!model}>
    {#snippet menu(close)}
      {#each SUBTOTALS as s (s.id)}
        <button class="xl-menu-item" aria-pressed={model?.pt.subtotals === s.id} onclick={() => { close(); apply('Subtotals', (pt) => { pt.subtotals = s.id; }); }}>{t(s.label)}</button>
      {/each}
    {/snippet}
  </MenuButton>
  <MenuButton large icon="sum" label={t('pvGrandTotals')} title={t('pvGrandTotals')} disabled={!model}>
    {#snippet menu(close)}
      {#each GRAND as g (g.label)}
        <button
          class="xl-menu-item"
          aria-pressed={model?.pt.rowGrandTotals === g.cols && model?.pt.columnGrandTotals === g.rows}
          onclick={() => { close(); apply('Grand Totals', (pt) => { pt.rowGrandTotals = g.cols; pt.columnGrandTotals = g.rows; }); }}>{t(g.label)}</button
        >
      {/each}
    {/snippet}
  </MenuButton>
  <MenuButton large icon="page-layout" label={t('pvReportLayout')} title={t('pvReportLayout')} disabled={!model}>
    {#snippet menu(close)}
      {#each LAYOUTS as l (l.id)}
        <button class="xl-menu-item" aria-pressed={model?.pt.layout === l.id} onclick={() => { close(); apply('Report Layout', (pt) => { pt.layout = l.id; }); }}>{t(l.label)}</button>
      {/each}
    {/snippet}
  </MenuButton>
</Group>
