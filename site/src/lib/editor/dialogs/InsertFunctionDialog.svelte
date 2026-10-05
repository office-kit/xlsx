<script lang="ts">
  import { FUNCTION_CATALOG, type FunctionCategory } from '../calc/index.ts';
  import { insertFunctionCall } from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const CATEGORIES: ReadonlyArray<[FunctionCategory, MessageKey]> = [
    ['Financial', 'catFinancial'],
    ['Date & Time', 'catDateTime'],
    ['Math & Trig', 'catMath'],
    ['Statistical', 'catStatistical'],
    ['Lookup & Reference', 'catLookup'],
    ['Database', 'catDatabase'],
    ['Text', 'catText'],
    ['Logical', 'catLogical'],
    ['Information', 'catInformation'],
    ['Engineering', 'catEngineering'],
  ];

  const ctl = getEditor();
  let query = $state('');
  let category = $state<FunctionCategory | 'all'>('all');
  let selected = $state<string | null>(FUNCTION_CATALOG[0]?.name ?? null);
  let listEl = $state<HTMLDivElement>();

  const matches = $derived.by(() => {
    const q = query.trim().toLowerCase();
    return FUNCTION_CATALOG.filter((f) => (category === 'all' || f.category === category) && (!q || f.name.toLowerCase().includes(q) || f.description.toLowerCase().includes(q)));
  });
  // The selection falls back to the first match when filtering hides it, as the Formula Builder does.
  const current = $derived(matches.find((f) => f.name === selected) ?? matches[0]);

  function onok(): boolean {
    if (!current) return false;
    insertFunctionCall(ctl, current.name);
    return true;
  }

  function onListKey(e: KeyboardEvent): void {
    if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
    e.preventDefault();
    const i = matches.findIndex((f) => f.name === current?.name);
    const next = matches[Math.max(0, Math.min(matches.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1)))];
    if (!next) return;
    selected = next.name;
    listEl?.querySelector<HTMLElement>(`[data-fn="${next.name}"]`)?.focus();
  }
</script>

<Dialog title={t('insertFunction')} width={480} {onok} okLabel={t('dlgInsertFunctionOk')} okDisabled={!current}>
  <div class="row">
    <label for="if-q">{t('dlgSearchFunction')}</label>
    <input id="if-q" class="xl-input grow" bind:value={query} placeholder={t('dlgSearchFunctionHint')} spellcheck="false" onkeydown={onListKey} />
  </div>
  <div class="row">
    <label for="if-cat">{t('dlgCategory')}</label>
    <select id="if-cat" class="xl-select grow" bind:value={category}>
      <option value="all">{t('dlgAllCategories')}</option>
      {#each CATEGORIES as [value, label] (value)}<option {value}>{t(label)}</option>{/each}
    </select>
  </div>
  <div class="lbl">{t('dlgSelectFunction')}</div>
  <div class="list fns" role="listbox" aria-label={t('dlgSelectFunction')} bind:this={listEl} tabindex="-1" onkeydown={onListKey}>
    {#each matches as f (f.name)}
      <button role="option" data-fn={f.name} aria-selected={current?.name === f.name} onclick={() => (selected = f.name)} ondblclick={() => { selected = f.name; if (onok()) ctl.closeDialog(); }}>{f.name}</button>
    {:else}
      <div class="hint empty">{t('dlgNoFunctions')}</div>
    {/each}
  </div>
  {#if current}
    <p class="syntax">{current.syntax}</p>
    <p class="hint">{current.description}</p>
  {/if}
</Dialog>

<style>
  .fns {
    height: 180px;
  }
  .syntax {
    font-weight: 600;
    font-family: var(--xl-mono);
    font-size: 12px;
    margin: 8px 0 2px;
  }
  .empty {
    padding: 8px;
  }
</style>
