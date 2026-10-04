<script lang="ts">
  import * as A from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { quickSort } from '../core/data.ts';
  import { clearAllFilters, reapplyFilters } from '../core/filter.ts';
  import { clearAdvancedFilter, hasAdvancedFilter } from '../core/advanced-filter.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';

  const ctl = getEditor();
  const hasFilter = $derived.by(() => {
    void ctl.doc.version;
    return ctl.doc.ws.autoFilter !== undefined;
  });
  const hasAdvanced = $derived.by(() => {
    void ctl.doc.version;
    return hasAdvancedFilter(ctl.doc.ws);
  });
</script>

<Group label={t('groupSortFilter')}>
  <div class="col">
    <button class="xl-btn" title={t('sortAZ')} onclick={() => quickSort(ctl, false)}><Icon name="sort-asc" />{t('sortAZShort')}</button>
    <button class="xl-btn" title={t('sortZA')} onclick={() => quickSort(ctl, true)}><Icon name="sort-desc" />{t('sortZAShort')}</button>
  </div>
  <button class="xl-btn big" onclick={() => ctl.openDialog('sort')}><Icon name="sort" size={24} /><span>{t('sort')}</span></button>
  <button class="xl-btn big" aria-pressed={hasFilter} onclick={() => ctl.toggleFilter()}><Icon name="filter" size={24} /><span>{t('filter')}</span></button>
  <div class="col">
    <button
      class="xl-btn"
      disabled={!hasFilter && !hasAdvanced}
      onclick={() => {
        clearAllFilters(ctl);
        clearAdvancedFilter(ctl);
      }}>{t('clearFilter')}</button
    >
    <button class="xl-btn" disabled={!hasFilter} onclick={() => reapplyFilters(ctl)}>{t('reapply')}</button>
    <button class="xl-btn" onclick={() => ctl.openDialog('advancedFilter')}>{t('dtAdvanced')}</button>
  </div>
</Group>

<Group label={t('groupDataTools')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('textToColumns')}><Icon name="text-columns" size={24} /><span>{t('textToColumns')}</span></button>
  <button class="xl-btn big" onclick={() => A.flashFillActive(ctl)}><Icon name="fill-down" size={24} /><span>{t('flashFill')}</span></button>
  <button class="xl-btn big" onclick={() => ctl.openDialog('removeDuplicates')}><Icon name="duplicates" size={24} /><span>{t('removeDuplicates')}</span></button>
  <MenuButton large icon="validation" label={t('dataValidation')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { ctl.openDialog('dataValidation'); close(); }}>{t('dataValidationEllipsis')}</button>
      <button class="xl-menu-item" onclick={() => { A.circleInvalid(ctl); close(); }}>{t('circleInvalidData')}</button>
      <button class="xl-menu-item" onclick={() => { A.clearInvalidCircles(ctl); close(); }}>{t('clearValidationCircles')}</button>
    {/snippet}
  </MenuButton>
  <button class="xl-btn big" onclick={() => ctl.openDialog('consolidate')}><Icon name="insert" size={24} /><span>{t('dtConsolidate')}</span></button>
</Group>

<Group label={t('dtGroupForecast')}>
  <MenuButton large icon="calc" label={t('dtWhatIf')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { ctl.openDialog('scenarioManager'); close(); }}>{t('dtScenarioManagerEllipsis')}</button>
      <button class="xl-menu-item" onclick={() => { ctl.openDialog('goalSeek'); close(); }}>{t('dtGoalSeekEllipsis')}</button>
      <button class="xl-menu-item" onclick={() => { ctl.openDialog('dataTable'); close(); }}>{t('dtDataTableEllipsis')}</button>
    {/snippet}
  </MenuButton>
</Group>

<Group label={t('groupOutline')}>
  <button class="xl-btn big" onclick={() => A.groupSelection(ctl, true)}><Icon name="group" size={24} /><span>{t('group')}</span></button>
  <button class="xl-btn big" onclick={() => A.groupSelection(ctl, false)}><Icon name="ungroup" size={24} /><span>{t('ungroup')}</span></button>
  <button class="xl-btn big" onclick={() => ctl.openDialog('subtotal')}><Icon name="sum" size={24} /><span>{t('dtSubtotal')}</span></button>
  <div class="col">
    <button class="xl-btn" onclick={() => A.showDetail(ctl, true)}>{t('showDetail')}</button>
    <button class="xl-btn" onclick={() => A.showDetail(ctl, false)}>{t('hideDetail')}</button>
  </div>
</Group>

<style>
  .col {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
</style>
