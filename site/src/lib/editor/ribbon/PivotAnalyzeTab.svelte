<script lang="ts">
  // PivotTable Analyze, the contextual tab of a cell inside a PivotTable.
  // A pivot the editor can only show (kept verbatim from the file) gets the
  // Field List toggle and nothing that would change it.
  import { getEditor } from '../core/context.ts';
  import { editPivot } from '../core/pivot.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Group from './Group.svelte';

  const ctl = getEditor();
  const active = $derived(ctl.activePivot);
  const model = $derived(active?.kind === 'model' ? active : undefined);
  const name = $derived(active?.kind === 'model' ? active.pt.name : (active?.summary.name ?? ''));

  function refresh(): void {
    const m = model;
    if (m) editPivot(ctl, m.ws, m.index, 'Refresh PivotTable', () => {});
  }
</script>

<Group label={t('pvGroupPivot')}>
  <div class="name">
    <span>{t('pvPivotName')}</span>
    <input class="xl-input" value={name} readonly aria-label={t('pvPivotName')} />
  </div>
</Group>

<Group label={t('pvGroupData')}>
  <button class="xl-btn big" disabled={!model} onclick={refresh}>
    <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true"><path d="M19 12a7 7 0 1 1-2.05-4.95" fill="none" stroke="currentColor" stroke-width="1.6" /><path d="M19 4v4h-4" fill="none" stroke="var(--xl-accent)" stroke-width="1.6" /></svg>
    <span>{t('pvRefresh')}</span>
  </button>
  <button class="xl-btn big" disabled={!model} onclick={() => ctl.openDialog('createPivotTable', { change: true })}>
    <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="5" width="13" height="14" fill="none" stroke="currentColor" stroke-width="1.4" /><path d="M3 10h13M8 5v14" stroke="currentColor" stroke-width="1.2" /><path d="M14 16l6-6 2 2-6 6h-2z" fill="var(--xl-accent)" /></svg>
    <span>{t('pvChangeSource')}</span>
  </button>
</Group>

<Group label={t('pvGroupShow')}>
  <button class="xl-btn big" aria-pressed={ctl.pivotFieldList} onclick={() => (ctl.pivotFieldList = !ctl.pivotFieldList)}>
    <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="3" width="16" height="18" fill="none" stroke="currentColor" stroke-width="1.4" /><path d="M7 8h2M11 8h6M7 12h2M11 12h6M7 16h2M11 16h6" stroke="currentColor" stroke-width="1.4" /></svg>
    <span>{t('pvFieldList')}</span>
  </button>
</Group>

<style>
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
  .name {
    display: flex;
    flex-direction: column;
    gap: 4px;
    font-size: 11px;
  }
  .name input {
    width: 120px;
  }
</style>
