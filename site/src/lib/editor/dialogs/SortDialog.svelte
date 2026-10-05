<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { columnLabel, dataRange, guessHeader, sortRange, type SortKey } from '../core/data.ts';
  import { CUSTOM_LISTS } from '../core/fill.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  interface Level {
    id: number;
    /** Column (or row, sorting left to right) the level sorts on. */
    line: number;
    order: 'asc' | 'desc' | 'custom';
    list: number;
  }

  /** Excel allows 64 sort levels. */
  const MAX_LEVELS = 64;

  const ctl = getEditor();
  const range = dataRange(ctl);
  let hasHeader = $state(guessHeader(ctl.doc.ws, range));
  let caseSensitive = $state(false);
  let leftToRight = $state(false);
  let showOptions = $state(false);
  let nextId = 1;
  const active = ctl.doc.selection.active;
  let levels = $state<Level[]>([{ id: nextId++, line: Math.min(Math.max(active.col, range.c1), range.c2), order: 'asc', list: 0 }]);
  let selected = $state(0);
  let notice = $state<string | null>(null);

  const lines = $derived.by(() => {
    const out: Array<{ value: number; label: string }> = [];
    if (leftToRight) {
      for (let r = range.r1; r <= range.r2; r++) out.push({ value: r, label: `${t('dlgRow')} ${r}` });
    } else {
      for (let c = range.c1; c <= range.c2; c++) out.push({ value: c, label: columnLabel(ctl, range, c, hasHeader, t('dlgColumn')) });
    }
    return out;
  });

  function setOrientation(ltr: boolean): void {
    leftToRight = ltr;
    const first = ltr ? range.r1 : range.c1;
    levels = levels.map((l) => ({ ...l, line: first }));
  }

  function addLevel(): void {
    if (levels.length >= MAX_LEVELS) return;
    const used = new Set(levels.map((l) => l.line));
    const free = lines.find((o) => !used.has(o.value))?.value ?? lines[0]?.value ?? range.c1;
    levels.push({ id: nextId++, line: free, order: 'asc', list: 0 });
    selected = levels.length - 1;
  }

  function copyLevel(): void {
    const src = levels[selected];
    if (!src || levels.length >= MAX_LEVELS) return;
    levels.splice(selected + 1, 0, { ...src, id: nextId++ });
    selected++;
  }

  function deleteLevel(): void {
    if (levels.length === 0) return;
    levels.splice(selected, 1);
    selected = Math.max(0, Math.min(selected, levels.length - 1));
  }

  function move(step: -1 | 1): void {
    const to = selected + step;
    const a = levels[selected];
    const b = levels[to];
    if (!a || !b) return;
    levels[selected] = b;
    levels[to] = a;
    selected = to;
  }

  function onok(): boolean {
    const seen = new Set<number>();
    for (const l of levels) {
      if (seen.has(l.line)) {
        notice = t('dlgSortDuplicate', { name: lines.find((o) => o.value === l.line)?.label ?? '' });
        return false;
      }
      seen.add(l.line);
    }
    if (levels.length === 0) return true;
    const keys: SortKey[] = levels.map((l) => {
      const list = l.order === 'custom' ? CUSTOM_LISTS[l.list] : undefined;
      const descending = l.order === 'desc';
      return list ? { col: l.line, descending, caseSensitive, list } : { col: l.line, descending, caseSensitive };
    });
    sortRange(ctl, range, keys, hasHeader, leftToRight ? 'columns' : 'rows');
    return true;
  }
</script>

<Dialog title={t('sort')} width={600} {onok} bind:notice>
  <div class="toolbar">
    <button class="xl-btn outlined" disabled={levels.length >= MAX_LEVELS} onclick={addLevel}>+ {t('dlgAddLevel')}</button>
    <button class="xl-btn outlined" disabled={levels.length === 0} onclick={deleteLevel}>− {t('dlgDeleteLevel')}</button>
    <button class="xl-btn outlined" disabled={levels.length === 0 || levels.length >= MAX_LEVELS} onclick={copyLevel}>{t('dlgCopyLevel')}</button>
    <button class="xl-btn outlined" disabled={selected <= 0} onclick={() => move(-1)} aria-label={t('dlgMoveUp')}>▲</button>
    <button class="xl-btn outlined" disabled={selected >= levels.length - 1} onclick={() => move(1)} aria-label={t('dlgMoveDown')}>▼</button>
    <button class="xl-btn outlined" aria-pressed={showOptions} onclick={() => (showOptions = !showOptions)}>{t('dlgOptionsEllipsis')}</button>
    <label class="check header"><input type="checkbox" bind:checked={hasHeader} disabled={leftToRight} />{t('dlgMyDataHasHeaders')}</label>
  </div>
  <div class="levels">
    <div class="head" aria-hidden="true">
      <span></span><span>{leftToRight ? t('dlgRow') : t('dlgColumn')}</span><span>{t('dlgSortOn')}</span><span>{t('dlgOrder')}</span>
    </div>
    {#each levels as level, i (level.id)}
      <div class="level" class:selected={selected === i} role="group" aria-label={i === 0 ? t('dlgSortBy') : t('dlgThenBy')} onfocusin={() => (selected = i)} onpointerdown={() => (selected = i)}>
        <span class="by">{i === 0 ? t('dlgSortBy') : t('dlgThenBy')}</span>
        <select class="xl-select" bind:value={level.line} aria-label={leftToRight ? t('dlgRow') : t('dlgColumn')}>
          {#each lines as o (o.value)}<option value={o.value}>{o.label}</option>{/each}
        </select>
        <select class="xl-select" disabled aria-label={t('dlgSortOn')}><option>{t('dlgCellValues')}</option></select>
        <div class="order">
          <select class="xl-select" bind:value={level.order} aria-label={t('dlgOrder')}>
            <option value="asc">{t('dlgAtoZ')}</option>
            <option value="desc">{t('dlgZtoA')}</option>
            <option value="custom">{t('dlgCustomList')}</option>
          </select>
          {#if level.order === 'custom'}
            <select class="xl-select" bind:value={level.list} aria-label={t('dlgCustomList')}>
              {#each CUSTOM_LISTS as list, k (k)}<option value={k}>{list.join(', ')}</option>{/each}
            </select>
          {/if}
        </div>
      </div>
    {/each}
  </div>
  {#if showOptions}
    <fieldset>
      <legend>{t('dlgSortOptions')}</legend>
      <label class="check"><input type="checkbox" bind:checked={caseSensitive} />{t('dlgCaseSensitive')}</label>
      <div class="row" role="radiogroup" aria-label={t('orientation')}>
        <span class="lbl">{t('orientation')}</span>
        <label class="check"><input type="radio" name="sort-orient" checked={!leftToRight} onchange={() => setOrientation(false)} />{t('dlgTopToBottom')}</label>
        <label class="check"><input type="radio" name="sort-orient" checked={leftToRight} onchange={() => setOrientation(true)} />{t('dlgLeftToRight')}</label>
      </div>
    </fieldset>
  {/if}
</Dialog>

<style>
  .toolbar {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    align-items: center;
    margin-bottom: 8px;
  }
  .header {
    margin-left: auto;
  }
  .levels {
    border: 1px solid var(--xl-border-strong);
    background: #fff;
    border-radius: 3px;
    min-height: 120px;
    max-height: 220px;
    overflow: auto;
  }
  .head,
  .level {
    display: grid;
    grid-template-columns: 70px 1fr 120px 1fr;
    gap: 6px;
    align-items: start;
    padding: 4px 6px;
  }
  .head {
    font-weight: 600;
    border-bottom: 1px solid var(--xl-border);
    background: #f7f7f7;
  }
  .level.selected {
    background: var(--xl-selected);
  }
  .by {
    padding-top: 4px;
  }
  .order {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .xl-select {
    min-width: 0;
  }
</style>
