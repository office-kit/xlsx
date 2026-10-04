<script lang="ts">
  import * as A from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';
  import { FUNCTION_CATALOG } from '../calc/index.ts';

  const ctl = getEditor();
  const CATEGORIES: Array<{ id: (typeof FUNCTION_CATALOG)[number]['category']; label: MessageKey; icon: string }> = [
    { id: 'Financial', label: 'catFinancial', icon: 'currency' },
    { id: 'Logical', label: 'catLogical', icon: 'check' },
    { id: 'Text', label: 'catText', icon: 'textbox' },
    { id: 'Date & Time', label: 'catDateTime', icon: 'calc' },
    { id: 'Lookup & Reference', label: 'catLookup', icon: 'search' },
    { id: 'Math & Trig', label: 'catMath', icon: 'function' },
  ];
  const OTHER: Array<{ id: (typeof FUNCTION_CATALOG)[number]['category']; label: MessageKey }> = [
    { id: 'Statistical', label: 'catStatistical' },
    { id: 'Engineering', label: 'catEngineering' },
    { id: 'Information', label: 'catInformation' },
    { id: 'Database', label: 'catDatabase' },
  ];
  const byCategory = (cat: string) => FUNCTION_CATALOG.filter((f) => f.category === cat).map((f) => f.name);
  const manual = $derived.by(() => {
    void ctl.doc.version;
    return ctl.doc.wb.calcProperties?.calcMode === 'manual';
  });
</script>

<Group label={t('groupFunctionLibrary')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('insertFunction')}><Icon name="function" size={24} /><span>{t('insertFunction')}</span></button>
  <MenuButton large split icon="sum" label={t('autoSum')} onmain={() => A.autoSum(ctl)}>
    {#snippet menu(close)}
      {#each [['SUM', 'fnSum'], ['AVERAGE', 'fnAverage'], ['COUNT', 'fnCount'], ['MAX', 'fnMax'], ['MIN', 'fnMin']] as const as [fn, key] (fn)}
        <button class="xl-menu-item" onclick={() => { A.autoSum(ctl, fn); close(); }}>{t(key)}</button>
      {/each}
    {/snippet}
  </MenuButton>
  {#each CATEGORIES as cat (cat.id)}
    <MenuButton large icon={cat.icon} label={t(cat.label)}>
      {#snippet menu(close)}
        <div class="fn-list">
          {#each byCategory(cat.id) as name (name)}
            <button class="xl-menu-item" onclick={() => { A.insertFunctionCall(ctl, name); close(); }}>{name}</button>
          {/each}
        </div>
      {/snippet}
    </MenuButton>
  {/each}
  <MenuButton large icon="more" label={t('moreFunctionsShort')}>
    {#snippet menu(close)}
      {#each OTHER as cat (cat.id)}
        <div class="menu-head">{t(cat.label)}</div>
        <div class="fn-list short">
          {#each byCategory(cat.id) as name (name)}
            <button class="xl-menu-item" onclick={() => { A.insertFunctionCall(ctl, name); close(); }}>{name}</button>
          {/each}
        </div>
      {/each}
    {/snippet}
  </MenuButton>
</Group>

<Group label={t('groupDefinedNames')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('nameManager')}><Icon name="name" size={24} /><span>{t('nameManager')}</span></button>
  <div class="col">
    <button class="xl-btn" onclick={() => ctl.openDialog('defineName')}>{t('defineName')}</button>
    <MenuButton label={t('useInFormula')}>
      {#snippet menu(close)}
        {#each ctl.doc.wb.definedNames.filter((d) => !d.name.startsWith('_xlnm.') && !d.hidden) as dn (dn.name + (dn.scope ?? ''))}
          <button class="xl-menu-item" onclick={() => { A.insertNameIntoFormula(ctl, dn.name); close(); }}>{dn.name}</button>
        {:else}
          <div class="empty">{t('noNames')}</div>
        {/each}
      {/snippet}
    </MenuButton>
    <button class="xl-btn" onclick={() => A.createNamesFromSelection(ctl)}>{t('createFromSelection')}</button>
  </div>
</Group>

<Group label={t('groupFormulaAuditing')}>
  <div class="col">
    <button class="xl-btn" onclick={() => A.traceCells(ctl, 'precedents')}><Icon name="trace" />{t('tracePrecedents')}</button>
    <button class="xl-btn" onclick={() => A.traceCells(ctl, 'dependents')}><Icon name="trace" />{t('traceDependents')}</button>
    <button class="xl-btn" onclick={() => A.traceCells(ctl, 'clear')}>{t('removeArrows')}</button>
  </div>
  <div class="col">
    <button class="xl-btn" aria-pressed={ctl.showFormulas} onclick={() => A.setViewFlag(ctl, 'showFormulas', !ctl.showFormulas)}><Icon name="show-formulas" />{t('showFormulas')}</button>
    <button class="xl-btn" onclick={() => A.errorChecking(ctl)}>{t('errorChecking')}</button>
    <button class="xl-btn" onclick={() => ctl.openDialog('evaluateFormula')}>{t('evaluateFormula')}</button>
  </div>
  <button class="xl-btn big" aria-pressed={ctl.watchWindow} onclick={() => (ctl.watchWindow = !ctl.watchWindow)}><Icon name="watch" size={24} /><span>{t('watchWindow')}</span></button>
</Group>

<Group label={t('groupCalculation')}>
  <MenuButton large icon="calc" label={t('calculationOptions')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { A.setCalcMode(ctl, 'auto'); close(); }}>{manual ? '' : '✓ '}{t('calcAutomatic')}</button>
      <button class="xl-menu-item" onclick={() => { A.setCalcMode(ctl, 'manual'); close(); }}>{manual ? '✓ ' : ''}{t('calcManual')}</button>
    {/snippet}
  </MenuButton>
  <div class="col">
    <button class="xl-btn" onclick={() => ctl.doc.recalculate()}>{t('calculateNow')}</button>
    <button class="xl-btn" onclick={() => ctl.doc.recalculate()}>{t('calculateSheet')}</button>
  </div>
</Group>

<style>
  .col {
    display: flex;
    flex-direction: column;
    gap: 2px;
    align-items: flex-start;
  }
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
  .fn-list {
    max-height: 60vh;
    overflow: auto;
    min-width: 180px;
  }
  .fn-list.short {
    max-height: 160px;
  }
  .menu-head {
    font-weight: 600;
    font-size: 11px;
    color: var(--xl-text-2);
    padding: 6px 10px 2px;
    background: #f7f7f7;
  }
  .empty {
    padding: 6px 12px;
    color: var(--xl-text-3);
  }
</style>
