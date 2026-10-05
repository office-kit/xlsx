<script lang="ts">
  // The AutoFilter drop-down: sort, a condition (text or number filter), and
  // the searchable value checklist. Changes apply on OK, as on Windows Excel.
  import { onMount, untrack } from 'svelte';
  import type { CustomFilterCondition, FilterColumn } from '@office-kit/xlsx/worksheet';
  import type { CellPos } from '../core/address.ts';
  import { getCellAt } from '../core/cells.ts';
  import { getEditor } from '../core/context.ts';
  import { sortRange } from '../core/data.ts';
  import { activeCriteria, filterOwnerAt, filterValues, setColumnFilter, valuesColumn } from '../core/filter.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';

  let { at: openedAt }: { at: CellPos } = $props();
  // The menu works on the header it was opened from; GridOverlays remounts it for another header.
  const at = untrack(() => openedAt);
  const ctl = getEditor();
  const doc = ctl.doc;
  const owner = filterOwnerAt(doc.ws, at.row, at.col);

  // A long list is unusable and slow to render; Excel itself stops at 10,000.
  const LIST_LIMIT = 1000;

  const values = owner ? filterValues(ctl, owner, at.col) : [];
  const current = owner ? activeCriteria(ctl, owner).get(at.col) : undefined;
  let checked = $state(new Set(current?.values ?? values));
  let search = $state('');

  type Op = 'none' | 'eq' | 'ne' | 'begins' | 'ends' | 'contains' | 'notContains' | 'gt' | 'ge' | 'lt' | 'le' | 'between' | 'top10' | 'aboveAvg' | 'belowAvg';
  const numeric = $derived.by(() => {
    if (!owner) return false;
    let nums = 0;
    let texts = 0;
    for (let r = owner.range.r1 + 1; r <= Math.min(owner.range.r2, owner.range.r1 + 500); r++) {
      const v = getCellAt(doc.ws, r, at.col)?.value;
      const n = numberOf(v);
      if (n !== undefined) nums++;
      else if (typeof v === 'string') texts++;
    }
    return nums > texts;
  });
  const TEXT_OPS: Array<[Op, MessageKey]> = [
    ['eq', 'fEquals'],
    ['ne', 'fNotEquals'],
    ['begins', 'fBeginsWith'],
    ['ends', 'fEndsWith'],
    ['contains', 'fContains'],
    ['notContains', 'fNotContains'],
  ];
  const NUMBER_OPS: Array<[Op, MessageKey]> = [
    ['eq', 'fEquals'],
    ['ne', 'fNotEquals'],
    ['gt', 'fGreaterThan'],
    ['ge', 'fGreaterOrEqual'],
    ['lt', 'fLessThan'],
    ['le', 'fLessOrEqual'],
    ['between', 'fBetween'],
    ['top10', 'fTop10'],
    ['aboveAvg', 'fAboveAverage'],
    ['belowAvg', 'fBelowAverage'],
  ];
  let op = $state<Op>('none');
  let v1 = $state('');
  let v2 = $state('');

  function numberOf(v: unknown): number | undefined {
    if (typeof v === 'number') return v;
    if (v && typeof v === 'object' && 'kind' in v && v.kind === 'formula' && 'cachedValue' in v && typeof v.cachedValue === 'number') return v.cachedValue;
    return undefined;
  }

  const shown = $derived.by(() => {
    const q = search.trim().toLowerCase();
    const list = q ? values.filter((v) => v.toLowerCase().includes(q)) : values;
    return list.slice(0, LIST_LIMIT);
  });
  const allChecked = $derived(shown.every((v) => checked.has(v)));

  function toggle(v: string) {
    const next = new Set(checked);
    if (next.has(v)) next.delete(v);
    else next.add(v);
    checked = next;
  }

  function toggleAll() {
    const next = new Set(checked);
    for (const v of shown) {
      if (allChecked) next.delete(v);
      else next.add(v);
    }
    checked = next;
  }

  /** The Text / Number filter as Excel saves it (custom conditions, Top 10, above / below average). */
  function conditionColumn(): FilterColumn | undefined {
    if (op === 'none' || !owner) return undefined;
    const colId = at.col - owner.range.c1;
    const one = (operator: CustomFilterCondition['operator'], val: string): FilterColumn => ({
      kind: 'custom',
      colId,
      conditions: [operator === undefined ? { val } : { operator, val }],
    });
    switch (op) {
      case 'eq':
        return one(undefined, v1);
      case 'ne':
        return one('notEqual', v1);
      case 'begins':
        return one(undefined, `${v1}*`);
      case 'ends':
        return one(undefined, `*${v1}`);
      case 'contains':
        return one(undefined, `*${v1}*`);
      case 'notContains':
        return one('notEqual', `*${v1}*`);
      case 'gt':
        return one('greaterThan', v1);
      case 'ge':
        return one('greaterThanOrEqual', v1);
      case 'lt':
        return one('lessThan', v1);
      case 'le':
        return one('lessThanOrEqual', v1);
      case 'between':
        return { kind: 'custom', colId, and: true, conditions: [{ operator: 'greaterThanOrEqual', val: v1 }, { operator: 'lessThanOrEqual', val: v2 }] };
      case 'top10': {
        const n = Number(v1);
        return { kind: 'top10', colId, val: v1 !== '' && Number.isFinite(n) ? Math.max(1, Math.floor(n)) : 10 };
      }
      case 'aboveAvg':
        return { kind: 'dynamic', colId, type: 'aboveAverage' };
      case 'belowAvg':
        return { kind: 'dynamic', colId, type: 'belowAverage' };
    }
  }

  function apply() {
    if (!owner) return close();
    const condition = conditionColumn();
    const everything = values.every((v) => checked.has(v));
    // A saved condition (custom, Top 10, colour) has no checkbox form; untouched boxes keep it.
    if (!condition && everything && current && !current.values) return close();
    // A column holds one kind of filter: a condition wins over the value list, as in Excel's dialog.
    setColumnFilter(ctl, owner, at.col, condition ?? (everything ? undefined : valuesColumn(owner, at.col, checked)));
    close();
  }

  function clearColumn() {
    if (owner) setColumnFilter(ctl, owner, at.col, undefined);
    close();
  }

  function sort(descending: boolean) {
    if (owner) sortRange(ctl, owner.range, [{ col: at.col, descending }], true);
    close();
  }

  function close() {
    ctl.filterMenu = null;
  }

  let el: HTMLDivElement;
  let pos = $state({ x: 0, y: 0 });
  onMount(() => {
    const r = ctl.geometry.rectOf({ r1: at.row, c1: at.col, r2: at.row, c2: at.col });
    const host = el.offsetParent?.getBoundingClientRect();
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    const maxX = (host?.width ?? window.innerWidth) - w - 4;
    const maxY = (host?.height ?? window.innerHeight) - h - 4;
    pos = { x: Math.max(4, Math.min(r.x, maxX)), y: Math.max(4, Math.min(r.y + r.h, maxY)) };
    const down = (e: PointerEvent) => {
      if (!el.contains(e.target as Node)) close();
    };
    window.addEventListener('pointerdown', down, true);
    return () => window.removeEventListener('pointerdown', down, true);
  });
</script>

<div
  class="filter-menu xl-menu"
  bind:this={el}
  style:left="{pos.x}px"
  style:top="{pos.y}px"
  role="dialog"
  tabindex="-1"
  onpointerdown={(e) => e.stopPropagation()}
  onkeydown={(e) => {
    e.stopPropagation();
    if (e.key === 'Escape') close();
    else if (e.key === 'Enter' && !(e.target instanceof HTMLButtonElement)) apply();
  }}
>
  <button class="xl-menu-item" onclick={() => sort(false)}>{t('sortAZ')}</button>
  <button class="xl-menu-item" onclick={() => sort(true)}>{t('sortZA')}</button>
  <div class="xl-menu-sep"></div>
  <div class="row">
    <label for="filter-op">{numeric ? t('fNumberFilters') : t('fTextFilters')}</label>
    <select id="filter-op" class="xl-input" bind:value={op}>
      <option value="none">{t('fChooseOne')}</option>
      {#each numeric ? NUMBER_OPS : TEXT_OPS as [value, label] (value)}
        <option {value}>{t(label)}</option>
      {/each}
    </select>
  </div>
  {#if op !== 'none' && op !== 'aboveAvg' && op !== 'belowAvg'}
    <div class="row">
      <input class="xl-input" bind:value={v1} placeholder={op === 'top10' ? '10' : ''} aria-label="Value" />
      {#if op === 'between'}
        <span>–</span>
        <input class="xl-input" bind:value={v2} aria-label="Second value" />
      {/if}
    </div>
  {/if}
  <div class="xl-menu-sep"></div>
  <input class="xl-input search" type="search" placeholder={t('search')} bind:value={search} />
  <div class="list">
    <label class="item"><input type="checkbox" checked={allChecked} onchange={toggleAll} />{t('fSelectAll')}</label>
    {#each shown as v (v)}
      <label class="item"><input type="checkbox" checked={checked.has(v)} onchange={() => toggle(v)} />{v === '' ? t('fBlanks') : v}</label>
    {/each}
    {#if values.length > LIST_LIMIT && !search}
      <div class="more">{t('fListTruncated', { n: LIST_LIMIT })}</div>
    {/if}
  </div>
  <div class="actions">
    <button class="xl-btn" disabled={!current} onclick={clearColumn}>{t('fClear')}</button>
    <span class="spacer"></span>
    <button class="xl-btn primary" onclick={apply}>{t('ok')}</button>
    <button class="xl-btn" onclick={close}>{t('cancel')}</button>
  </div>
</div>

<style>
  .filter-menu {
    position: absolute;
    z-index: 20;
    width: 260px;
    padding: 4px 0;
  }
  .row {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 3px 10px;
    font-size: 12px;
  }
  .row .xl-input {
    flex: 1;
    min-width: 0;
  }
  .search {
    margin: 4px 10px;
    width: calc(100% - 20px);
  }
  .list {
    max-height: 220px;
    overflow: auto;
    margin: 0 10px;
    border: 1px solid var(--xl-border, #d1d1d1);
    padding: 2px 0;
  }
  .item {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 1px 6px;
    font-size: 12px;
    white-space: nowrap;
  }
  .more {
    padding: 2px 6px;
    font-size: 11px;
    color: var(--xl-text-3);
  }
  .actions {
    display: flex;
    gap: 6px;
    padding: 8px 10px 4px;
  }
  .spacer {
    flex: 1;
  }
</style>
