<script lang="ts">
  // The PivotTable Fields pane docked right of the grid, as in Excel: the field
  // list with check boxes over the four areas (Filters, Columns, Rows, Values).
  // Fields drag from the list into an area, between areas, and out of the
  // areas to remove them; a chip's menu does the same by click and holds the
  // value field settings. Every change is one undo step through editPivot.
  import type { PivotAggregate } from '@office-kit/xlsx/worksheet';
  import { getPivotFieldItems } from '@office-kit/xlsx/worksheet';
  import { getEditor } from '../core/context.ts';
  import { defaultAreaFor, editPivot, itemLabel, placeField, placementOf, removeFromArea, type PivotArea } from '../core/pivot.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Icon from './Icon.svelte';
  import Popup from './Popup.svelte';

  const ctl = getEditor();

  const AREAS: Array<{ id: PivotArea; label: MessageKey }> = [
    { id: 'filters', label: 'pvFilters' },
    { id: 'columns', label: 'pvColumns' },
    { id: 'rows', label: 'pvRows' },
    { id: 'values', label: 'pvValues' },
  ];
  const AGGREGATES: Array<{ id: PivotAggregate; label: MessageKey }> = [
    { id: 'sum', label: 'pvSum' },
    { id: 'count', label: 'pvCount' },
    { id: 'average', label: 'pvAverage' },
    { id: 'max', label: 'pvMax' },
    { id: 'min', label: 'pvMin' },
  ];
  const MOVE_TO: Record<PivotArea, MessageKey> = {
    filters: 'pvMoveToFilters',
    columns: 'pvMoveToColumns',
    rows: 'pvMoveToRows',
    values: 'pvMoveToValues',
  };

  const active = $derived(ctl.activePivot);

  const aggregateLabel = (a: PivotAggregate): string => t(AGGREGATES.find((x) => x.id === a)?.label ?? 'pvSum');
  // Captions the file keeps unset; Excel derives them ("Sum of Sales", 「合計 / Sales」).
  const defaultCaption = (a: PivotAggregate, field: string): string => t('pvValueCaption', { agg: aggregateLabel(a), field });
  const model = $derived(active?.kind === 'model' ? active : undefined);

  interface Chip {
    area: PivotArea;
    position: number;
    field: number;
    label: string;
  }

  /** Display rows of every area. The Σ Values chip stands for the value fields' place among the columns. */
  const chips = $derived.by((): Record<PivotArea, Chip[]> => {
    const out: Record<PivotArea, Chip[]> = { filters: [], columns: [], rows: [], values: [] };
    if (active?.kind === 'readonly') {
      const s = active.summary;
      const ro = (area: PivotArea, names: string[]) => names.map((label, position) => ({ area, position, field: -1, label }));
      return { filters: ro('filters', s.filters), columns: ro('columns', s.columns), rows: ro('rows', s.rows), values: ro('values', s.values) };
    }
    if (!model) return out;
    const { pt, fields } = model;
    const name = (f: number) => fields[f] ?? `#${f + 1}`;
    out.filters = pt.filters.map((f, position) => ({ area: 'filters', position, field: f.field, label: name(f.field) }));
    out.columns = pt.columns.map((field, position) => ({ area: 'columns', position, field, label: name(field) }));
    out.rows = pt.rows.map((field, position) => ({ area: 'rows', position, field, label: name(field) }));
    out.values = pt.values.map((v, position) => ({
      area: 'values',
      position,
      field: v.field,
      label: v.name ?? defaultCaption(v.aggregate, name(v.field)),
    }));
    return out;
  });

  const fieldNames = $derived(active?.kind === 'readonly' ? active.summary.fields : (model?.fields ?? []));
  const placedNames = $derived(
    new Set(active?.kind === 'readonly' ? [...active.summary.rows, ...active.summary.columns, ...active.summary.filters] : []),
  );

  function isChecked(i: number): boolean {
    if (model) return placementOf(model.pt, i) !== undefined;
    const name = fieldNames[i] ?? '';
    return placedNames.has(name) || (active?.kind === 'readonly' && active.summary.values.some((v) => v.endsWith(name)));
  }

  function edit(label: string, fn: (pt: NonNullable<typeof model>['pt']) => void): void {
    const m = model;
    if (!m) return;
    editPivot(ctl, m.ws, m.index, label, fn);
  }

  function toggleField(field: number, on: boolean): void {
    const m = model;
    if (!m) return;
    if (on) {
      const area = defaultAreaFor(ctl.doc.wb, m.pt, field);
      edit('Add PivotTable Field', (pt) => placeField(pt, field, area));
    } else {
      edit('Remove PivotTable Field', (pt) => {
        pt.rows = pt.rows.filter((f) => f !== field);
        pt.columns = pt.columns.filter((f) => f !== field);
        pt.filters = pt.filters.filter((f) => f.field !== field);
        pt.values = pt.values.filter((v) => v.field !== field);
      });
    }
  }

  // ---- drag and drop ----------------------------------------------------------

  type DragSource = { from: 'list'; field: number } | { from: 'chip'; chip: Chip };
  let dragging: DragSource | null = null;
  let dropArea = $state<PivotArea | null>(null);
  let dropBefore = $state<number | null>(null);

  function startDrag(e: DragEvent, src: DragSource): void {
    if (!model) return;
    dragging = src;
    e.dataTransfer?.setData('text/plain', src.from === 'list' ? String(src.field) : src.chip.label);
    if (e.dataTransfer) e.dataTransfer.effectAllowed = 'move';
  }

  function over(e: DragEvent, area: PivotArea, before: number | null): void {
    if (!dragging) return;
    e.preventDefault();
    e.stopPropagation();
    dropArea = area;
    dropBefore = before;
  }

  function endDrag(): void {
    dragging = null;
    dropArea = null;
    dropBefore = null;
  }

  function drop(e: DragEvent, area: PivotArea): void {
    e.preventDefault();
    const src = dragging;
    const before = dropBefore;
    endDrag();
    if (!src) return;
    moveTo(src, area, before ?? undefined);
  }

  /** Put the dragged field into `area` before position `before` (end when undefined). */
  function moveTo(src: DragSource, area: PivotArea, before?: number): void {
    const field = src.from === 'list' ? src.field : src.chip.field;
    edit('Move PivotTable Field', (pt) => {
      let at = before;
      // A value keeps its function when moved; a field newly dropped on Values
      // sums if it holds numbers and counts otherwise, as in Excel.
      const kept = src.from === 'chip' && src.chip.area === 'values' ? pt.values[src.chip.position]?.aggregate : undefined;
      const aggregate = kept ?? (defaultAreaFor(ctl.doc.wb, pt, field) === 'values' ? 'sum' : 'count');
      if (src.from === 'chip') {
        // Within one area, taking the chip out first shifts the later slots up.
        if (src.chip.area === area && at !== undefined && src.chip.position < at) at--;
        removeFromArea(pt, src.chip.area, src.chip.position);
      }
      placeField(pt, field, area, at, aggregate);
    });
  }

  function dropOnList(e: DragEvent): void {
    e.preventDefault();
    const src = dragging;
    endDrag();
    if (src?.from !== 'chip') return;
    edit('Remove PivotTable Field', (pt) => removeFromArea(pt, src.chip.area, src.chip.position));
  }

  // ---- chip menu ----------------------------------------------------------------

  let menu = $state<{ chip: Chip; anchor: HTMLElement } | null>(null);
  let settings = $state<{ position: number; aggregate: PivotAggregate; name: string; anchor: HTMLElement } | null>(null);

  function openMenu(e: MouseEvent, chip: Chip): void {
    if (!model) return;
    menu = { chip, anchor: e.currentTarget as HTMLElement };
  }

  function menuMove(delta: -1 | 1): void {
    const m = menu;
    menu = null;
    if (!m) return;
    const { chip } = m;
    edit('Move PivotTable Field', (pt) => {
      const list: unknown[] = chip.area === 'rows' ? pt.rows : chip.area === 'columns' ? pt.columns : chip.area === 'filters' ? pt.filters : pt.values;
      const to = chip.position + delta;
      if (to < 0 || to >= list.length) return;
      const [item] = list.splice(chip.position, 1);
      list.splice(to, 0, item);
    });
  }

  function menuMoveTo(area: PivotArea): void {
    const m = menu;
    menu = null;
    if (m) moveTo({ from: 'chip', chip: m.chip }, area);
  }

  function menuRemove(): void {
    const m = menu;
    menu = null;
    if (m) edit('Remove PivotTable Field', (pt) => removeFromArea(pt, m.chip.area, m.chip.position));
  }

  function openSettings(): void {
    const m = menu;
    menu = null;
    const v = m && model?.pt.values[m.chip.position];
    if (!m || !v) return;
    settings = { position: m.chip.position, aggregate: v.aggregate, name: m.chip.label, anchor: m.anchor };
  }

  function applySettings(): void {
    const s = settings;
    settings = null;
    if (!s) return;
    edit('Value Field Settings', (pt) => {
      const v = pt.values[s.position];
      if (!v) return;
      const field = model?.fields[v.field] ?? '';
      const name = s.name.trim();
      // A caption left at its default stays unset, so it follows the new function
      // ("Sum of X" becomes "Count of X") as it does in Excel.
      const isDefault = name === '' || name === defaultCaption(v.aggregate, field) || /^(Sum|Count|Average|Max|Min) of /.test(name);
      v.aggregate = s.aggregate;
      if (isDefault) delete v.name;
      else v.name = name;
    });
  }

  function filterItems(field: number): Array<{ label: string; value: string | number | boolean | null }> {
    if (!model) return [];
    return getPivotFieldItems(ctl.doc.wb, model.pt.source, field).map((value) => ({ label: itemLabel(value), value }));
  }

  function setFilter(position: number, value: string | number | boolean | null | undefined): void {
    menu = null;
    edit('Filter PivotTable', (pt) => {
      const f = pt.filters[position];
      if (!f) return;
      if (value === undefined) delete f.selected;
      else f.selected = value;
    });
  }
</script>

{#if active}
  <aside class="pane" aria-label={t('pvFields')}>
    <header>
      <span class="title">{t('pvFields')}</span>
      <button class="xl-btn icon" title={t('pvClosePane')} onclick={() => (ctl.pivotFieldList = false)}><Icon name="close" size={14} /></button>
    </header>
    {#if active.kind === 'readonly'}
      <p class="note">{t('pvReadOnly')}</p>
    {/if}
    <div class="caption">{t('pvChooseFields')}</div>
    <ul class="fields" ondragover={(e) => { if (dragging?.from === 'chip') e.preventDefault(); }} ondrop={dropOnList}>
      {#each fieldNames as name, i (i)}
        <li draggable={!!model} ondragstart={(e) => startDrag(e, { from: 'list', field: i })} ondragend={endDrag}>
          <label>
            <input type="checkbox" checked={isChecked(i)} disabled={!model} onchange={(e) => toggleField(i, e.currentTarget.checked)} />
            <span>{name}</span>
          </label>
        </li>
      {/each}
    </ul>
    <div class="caption">{t('pvDragHint')}</div>
    <div class="areas">
      {#each AREAS as area (area.id)}
        <section
          class="area"
          class:target={dropArea === area.id}
          data-area={area.id}
          aria-label={t(area.label)}
          ondragover={(e) => over(e, area.id, null)}
          ondragleave={(e) => { if (e.currentTarget === e.target) dropArea = null; }}
          ondrop={(e) => drop(e, area.id)}
        >
          <div class="area-title">{t(area.label)}</div>
          <ul>
            {#each chips[area.id] as chip (chip.position)}
              <li
                class="chip"
                class:insert={dropArea === area.id && dropBefore === chip.position}
                draggable={!!model}
                ondragstart={(e) => startDrag(e, { from: 'chip', chip })}
                ondragend={endDrag}
                ondragover={(e) => over(e, area.id, chip.position)}
              >
                <button class="chip-btn" disabled={!model} onclick={(e) => openMenu(e, chip)}>
                  <span class="chip-label">{chip.label}</span>
                  {#if model}<Icon name="chevron-down" size={12} />{/if}
                </button>
              </li>
            {/each}
            {#if area.id === 'columns' && (model ? model.pt.values.length > 1 : active.kind === 'readonly' && active.summary.columns.length === 0 && active.summary.values.length > 1)}
              <li class="chip sigma"><span class="chip-label">{t('pvValuesField')}</span></li>
            {/if}
          </ul>
        </section>
      {/each}
    </div>
  </aside>
{/if}

{#if menu}
  {@const chip = menu.chip}
  <Popup anchor={menu.anchor} onclose={() => (menu = null)}>
    <button class="xl-menu-item" disabled={chip.position === 0} onclick={() => menuMove(-1)}>{t('pvMoveUp')}</button>
    <button class="xl-menu-item" disabled={chip.position === chips[chip.area].length - 1} onclick={() => menuMove(1)}>{t('pvMoveDown')}</button>
    <div class="xl-menu-sep"></div>
    {#each AREAS as area (area.id)}
      {#if area.id !== chip.area}
        <button class="xl-menu-item" onclick={() => menuMoveTo(area.id)}>{t(MOVE_TO[area.id])}</button>
      {/if}
    {/each}
    <div class="xl-menu-sep"></div>
    <button class="xl-menu-item" onclick={menuRemove}>{t('pvRemoveField')}</button>
    {#if chip.area === 'values'}
      <div class="xl-menu-sep"></div>
      <button class="xl-menu-item" onclick={openSettings}>{t('pvValueSettings')}</button>
    {/if}
    {#if chip.area === 'filters'}
      <div class="xl-menu-sep"></div>
      <div class="menu-caption">{t('pvShowItem')}</div>
      <div class="items">
        <button class="xl-menu-item" aria-pressed={model?.pt.filters[chip.position]?.selected === undefined} onclick={() => setFilter(chip.position, undefined)}>{t('pvAll')}</button>
        {#each filterItems(chip.field) as item, k (k)}
          <button class="xl-menu-item" aria-pressed={model?.pt.filters[chip.position]?.selected === item.value} onclick={() => setFilter(chip.position, item.value)}>{item.label}</button>
        {/each}
      </div>
    {/if}
  </Popup>
{/if}

{#if settings}
  <Popup anchor={settings.anchor} onclose={() => (settings = null)}>
    <div class="settings">
      <label class="name">{t('pvCustomName')}<input class="xl-input" bind:value={settings.name} /></label>
      <div class="menu-caption">{t('pvSummarizeBy')}</div>
      <select class="xl-input" size="5" bind:value={settings.aggregate} aria-label={t('pvSummarizeBy')}>
        {#each AGGREGATES as a (a.id)}
          <option value={a.id}>{t(a.label)}</option>
        {/each}
      </select>
      <div class="buttons">
        <button class="xl-btn" onclick={() => (settings = null)}>{t('cancel')}</button>
        <button class="xl-btn primary" onclick={applySettings}>{t('ok')}</button>
      </div>
    </div>
  </Popup>
{/if}

<style>
  .pane {
    display: flex;
    flex-direction: column;
    width: 280px;
    flex: none;
    min-height: 0;
    background: var(--xl-panel);
    border-left: 1px solid var(--xl-border);
    padding: 6px 10px 10px;
    gap: 4px;
    overflow: hidden;
  }
  header {
    display: flex;
    align-items: center;
  }
  .title {
    flex: 1;
    font-size: 14px;
    font-weight: 600;
  }
  .note {
    margin: 0;
    padding: 6px;
    background: var(--xl-accent-soft);
    border-radius: var(--xl-radius);
    color: var(--xl-text-2);
  }
  .caption {
    color: var(--xl-text-2);
    margin-top: 4px;
  }
  .fields {
    list-style: none;
    margin: 0;
    padding: 2px;
    border: 1px solid var(--xl-border);
    flex: 1 1 40%;
    min-height: 80px;
    overflow: auto;
  }
  .fields li {
    padding: 1px 2px;
    cursor: grab;
  }
  .fields label {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .areas {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px;
    flex: 1 1 50%;
    min-height: 0;
  }
  .area {
    display: flex;
    flex-direction: column;
    border: 1px solid var(--xl-border);
    min-height: 90px;
    overflow: hidden;
  }
  .area.target {
    border-color: var(--xl-accent);
    background: var(--xl-accent-soft);
  }
  .area-title {
    padding: 2px 6px;
    color: var(--xl-text-2);
    border-bottom: 1px solid var(--xl-border);
  }
  .area ul {
    list-style: none;
    margin: 0;
    padding: 2px;
    overflow: auto;
    flex: 1;
  }
  .chip {
    margin: 2px 0;
    border-top: 2px solid transparent;
  }
  .chip.insert {
    border-top-color: var(--xl-accent);
  }
  .chip-btn,
  .sigma {
    display: flex;
    align-items: center;
    width: 100%;
    gap: 4px;
    padding: 2px 4px;
    border: 1px solid var(--xl-border-strong);
    border-radius: 2px;
    background: var(--xl-ribbon);
    font: inherit;
    text-align: left;
    cursor: grab;
  }
  .chip-label {
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .menu-caption {
    padding: 4px 10px 2px;
    color: var(--xl-text-2);
  }
  .items {
    max-height: 220px;
    overflow: auto;
  }
  .settings {
    display: flex;
    flex-direction: column;
    gap: 6px;
    padding: 8px;
    width: 240px;
  }
  .settings .name {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .buttons {
    display: flex;
    justify-content: flex-end;
    gap: 6px;
  }
</style>
