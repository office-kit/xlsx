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
    <button class="xl-btn" title={t('sortAZ')} onclick={() => quickSort(ctl, false)}><Icon name="sort-asc" /></button>
    <button class="xl-btn" title={t('sortZA')} onclick={() => quickSort(ctl, true)}><Icon name="sort-desc" /></button>
  </div>
  <button class="xl-btn big" onclick={() => ctl.openDialog('sort')}><Icon name="sort" size={30} /><span>{t('sort')}</span></button>
  <button class="xl-btn big" aria-pressed={hasFilter} onclick={() => ctl.toggleFilter()}><Icon name="filter" size={30} /><span>{t('filter')}</span></button>
  <div class="col">
    <button
      class="xl-btn"
      disabled={!hasFilter && !hasAdvanced}
      onclick={() => {
        clearAllFilters(ctl);
        clearAdvancedFilter(ctl);
      }}><Icon name="filter-clear" />{t('clearFilter')}</button
    >
    <button class="xl-btn" disabled={!hasFilter} onclick={() => reapplyFilters(ctl)}><Icon name="filter-reapply" />{t('reapply')}</button>
    <button class="xl-btn" onclick={() => ctl.openDialog('advancedFilter')}><Icon name="filter-advanced" />{t('dtAdvanced')}</button>
  </div>
</Group>

<Group label={t('groupDataTools')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('textToColumns')}><Icon name="text-columns" size={30} /><span>{t('textToColumns')}</span></button>
  <div class="col">
    <button class="xl-btn" title={t('flashFill')} onclick={() => A.flashFillActive(ctl)}><Icon name="flash-fill" /></button>
    <button class="xl-btn" title={t('removeDuplicates')} onclick={() => ctl.openDialog('removeDuplicates')}><Icon name="duplicates" /></button>
    <MenuButton icon="validation" title={t('dataValidation')}>
      {#snippet menu(close)}
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('dataValidation'); close(); }}>{t('dataValidationEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => { A.circleInvalid(ctl); close(); }}>{t('circleInvalidData')}</button>
        <button class="xl-menu-item" onclick={() => { A.clearInvalidCircles(ctl); close(); }}>{t('clearValidationCircles')}</button>
      {/snippet}
    </MenuButton>
  </div>
  <div class="col">
    <button class="xl-btn" title={t('dtConsolidate')} onclick={() => ctl.openDialog('consolidate')}><Icon name="consolidate" /></button>
  </div>
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
  <div class="col">
    <button class="xl-btn" onclick={() => A.groupSelection(ctl, true)}><Icon name="group" />{t('group')}</button>
    <button class="xl-btn" onclick={() => A.groupSelection(ctl, false)}><Icon name="ungroup" />{t('ungroup')}</button>
    <button class="xl-btn" onclick={() => ctl.openDialog('subtotal')}><Icon name="sum" />{t('dtSubtotal')}</button>
  </div>
  <div class="col">
    <button class="xl-btn" title={t('showDetail')} onclick={() => A.showDetail(ctl, true)}><Icon name="show-detail" /></button>
    <button class="xl-btn" title={t('hideDetail')} onclick={() => A.showDetail(ctl, false)}><Icon name="hide-detail" /></button>
  </div>
</Group>
