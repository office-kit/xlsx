<script lang="ts">
  import * as A from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';
  import { FUNCTION_CATALOG } from '../calc/index.ts';
  import { recentFunctions } from '../core/recent-functions.svelte.ts';

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
  const known = new Set(FUNCTION_CATALOG.map((f) => f.name));
  const recent = $derived(recentFunctions.names.filter((n) => known.has(n)));
  const manual = $derived.by(() => {
    void ctl.doc.version;
    return ctl.doc.wb.calcProperties?.calcMode === 'manual';
  });
</script>

{#snippet category(cat: (typeof CATEGORIES)[number])}
  <MenuButton icon={cat.icon} label={t(cat.label)}>
    {#snippet menu(close)}
      <div class="fn-list">
        {#each byCategory(cat.id) as name (name)}
          <button class="xl-menu-item" onclick={() => { A.insertFunctionCall(ctl, name); close(); }}>{name}</button>
        {/each}
      </div>
    {/snippet}
  </MenuButton>
{/snippet}

<Group label={t('groupFunctionLibrary')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('insertFunction')}><Icon name="function" size={30} /><span>{t('insertFunction')}</span></button>
  <div class="col">
    <MenuButton icon="sum" label={t('autoSum')}>
      {#snippet menu(close)}
        {#each [['SUM', 'fnSum'], ['AVERAGE', 'fnAverage'], ['COUNT', 'fnCount'], ['MAX', 'fnMax'], ['MIN', 'fnMin']] as const as [fn, key] (fn)}
          <button class="xl-menu-item" onclick={() => { A.autoSum(ctl, fn); close(); }}>{t(key)}</button>
        {/each}
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('insertFunction'); close(); }}>{t('moreFunctions')}</button>
      {/snippet}
    </MenuButton>
    <MenuButton icon="recent" label={t('recentlyUsed')}>
      {#snippet menu(close)}
        <div class="fn-list">
          {#each recent as name (name)}
            <button class="xl-menu-item" onclick={() => { A.insertFunctionCall(ctl, name); close(); }}>{name}</button>
          {/each}
        </div>
      {/snippet}
    </MenuButton>
    {#each CATEGORIES.slice(0, 1) as cat (cat.id)}{@render category(cat)}{/each}
  </div>
  <div class="col">
    {#each CATEGORIES.slice(1, 4) as cat (cat.id)}{@render category(cat)}{/each}
  </div>
  <div class="col">
    {#each CATEGORIES.slice(4) as cat (cat.id)}{@render category(cat)}{/each}
    <MenuButton icon="more" label={t('moreFunctionsShort')}>
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
  </div>
</Group>

<Group label={t('groupDefinedNames')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('nameManager')}><Icon name="name" size={30} /><span>{t('nameManager')}</span></button>
  <div class="col">
    <button class="xl-btn" onclick={() => ctl.openDialog('defineName')}><Icon name="define-name" />{t('defineName')}</button>
    <MenuButton icon="use-in-formula" label={t('useInFormula')}>
      {#snippet menu(close)}
        {#each ctl.doc.wb.definedNames.filter((d) => !d.name.startsWith('_xlnm.') && !d.hidden) as dn (dn.name + (dn.scope ?? ''))}
          <button class="xl-menu-item" onclick={() => { A.insertNameIntoFormula(ctl, dn.name); close(); }}>{dn.name}</button>
        {:else}
          <div class="empty">{t('noNames')}</div>
        {/each}
      {/snippet}
    </MenuButton>
    <button class="xl-btn" onclick={() => A.createNamesFromSelection(ctl)}><Icon name="create-from-selection" />{t('createFromSelection')}</button>
  </div>
</Group>

<Group label={t('groupFormulaAuditing')}>
  <div class="col">
    <button class="xl-btn" onclick={() => A.traceCells(ctl, 'precedents')}><Icon name="trace" />{t('tracePrecedents')}</button>
    <button class="xl-btn" onclick={() => A.traceCells(ctl, 'dependents')}><Icon name="trace-dependents" />{t('traceDependents')}</button>
    <button class="xl-btn" onclick={() => A.traceCells(ctl, 'clear')}><Icon name="remove-arrows" />{t('removeArrows')}</button>
  </div>
  <div class="col">
    <button class="xl-btn" aria-pressed={ctl.showFormulas} onclick={() => A.setViewFlag(ctl, 'showFormulas', !ctl.showFormulas)}><Icon name="show-formulas" />{t('showFormulas')}</button>
    <button class="xl-btn" onclick={() => A.errorChecking(ctl)}><Icon name="error-check" />{t('errorChecking')}</button>
    <button class="xl-btn" onclick={() => ctl.openDialog('evaluateFormula')}><Icon name="evaluate" />{t('evaluateFormula')}</button>
  </div>
  <button class="xl-btn big" aria-pressed={ctl.watchWindow} onclick={() => (ctl.watchWindow = !ctl.watchWindow)}><Icon name="watch" size={30} /><span>{t('watchWindow')}</span></button>
</Group>

<Group label={t('groupCalculation')}>
  <MenuButton large icon="calc" label={t('calculationOptions')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { A.setCalcMode(ctl, 'auto'); close(); }}>{manual ? '' : '✓ '}{t('calcAutomatic')}</button>
      <button class="xl-menu-item" onclick={() => { A.setCalcMode(ctl, 'manual'); close(); }}>{manual ? '✓ ' : ''}{t('calcManual')}</button>
    {/snippet}
  </MenuButton>
  <div class="col">
    <button class="xl-btn" title={t('calculateNow')} onclick={() => ctl.doc.recalculate()}><Icon name="calc-now" /></button>
    <button class="xl-btn" title={t('calculateSheet')} onclick={() => ctl.doc.recalculate()}><Icon name="calc-sheet" /></button>
  </div>
</Group>

<style>
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
