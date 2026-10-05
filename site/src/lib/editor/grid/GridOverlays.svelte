<script lang="ts">
  // HTML widgets that sit on top of the canvas: AutoFilter buttons and their
  // menu, the data-validation drop-down arrow and list, the validation input
  // message, and note pop-ups. Only cells inside the viewport get widgets.
  import { untrack } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { cellAddress, type Range } from '../core/address.ts';
  import { validationAt } from '../core/data.ts';
  import { filterOwners } from '../core/filter.ts';
  import FilterMenu from './FilterMenu.svelte';
  import { autoFill, type FillMode } from '../core/actions.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import { threadAt } from '../core/comments.ts';
  import CommentCard from './CommentCard.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  const geo = $derived(ctl.geometry);

  function visibleRect(range: Range) {
    const r = geo.rectOf(range);
    if (r.w <= 0 || r.h <= 0 || r.x + r.w <= geo.headerW || r.y + r.h <= geo.headerH || r.x >= geo.width || r.y >= geo.height) return undefined;
    return r;
  }

  const filterButtons = $derived.by(() => {
    void doc.version;
    const out: Array<{ row: number; col: number; x: number; y: number; size: number; active: boolean }> = [];
    for (const owner of filterOwners(doc.ws)) {
      const filtered = new Set(owner.autoFilter.filterColumns.map((fc) => owner.range.c1 + fc.colId));
      const row = owner.range.r1;
      for (let col = owner.range.c1; col <= owner.range.c2; col++) {
        const r = visibleRect({ r1: row, c1: col, r2: row, c2: col });
        if (!r) continue;
        const size = Math.min(r.h - 2, Math.round(16 * Math.min(doc.zoom, 1.5)));
        if (size < 6) continue;
        out.push({ row, col, x: r.x + r.w - size - 2, y: r.y + r.h - size - 1, size, active: filtered.has(col) });
      }
    }
    return out;
  });

  const activeValidation = $derived.by(() => {
    void doc.version;
    const { row, col } = doc.selection.active;
    return validationAt(doc.ws, row, col);
  });

  // Excel inverts showDropDown on the wire: true hides the in-cell arrow.
  const listArrow = $derived.by(() => {
    const dv = activeValidation;
    if (!dv || dv.type !== 'list' || dv.showDropDown === true || ctl.edit) return undefined;
    const { row, col } = doc.selection.active;
    const merged = doc.merges.at(row, col) ?? { r1: row, c1: col, r2: row, c2: col };
    const r = visibleRect(merged);
    if (!r) return undefined;
    return { x: r.x + r.w + 1, y: r.y + r.h - 18, h: 18 };
  });

  const inputMessage = $derived.by(() => {
    const dv = activeValidation;
    if (!dv || dv.showInputMessage === false || (!dv.prompt && !dv.promptTitle)) return undefined;
    const { row, col } = doc.selection.active;
    const r = visibleRect({ r1: row, c1: col, r2: row, c2: col });
    if (!r) return undefined;
    return { x: r.x + 16, y: r.y + r.h + 6, title: dv.promptTitle, text: dv.prompt };
  });

  const noteIndex = $derived.by(() => {
    void doc.version;
    const m = new Map<string, { author: string; text: string }>();
    for (const c of doc.ws.legacyComments) m.set(c.ref.replaceAll('$', '').split(':')[0] ?? c.ref, c);
    return m;
  });

  const note = $derived.by(() => {
    const h = ctl.hoverCell;
    if (!h || ctl.edit) return undefined;
    const n = noteIndex.get(cellAddress(h.row, h.col));
    if (!n) return undefined;
    const r = visibleRect({ r1: h.row, c1: h.col, r2: h.row, c2: h.col });
    if (!r) return undefined;
    return { x: r.x + r.w + 8, y: r.y - 4, ...n };
  });

  // A pinned card wins; otherwise hovering a commented cell previews its thread.
  const commentCard = $derived.by(() => {
    if (ctl.edit) return undefined;
    void doc.version;
    const pinned = ctl.commentCard;
    const at = pinned ?? ctl.hoverCell;
    if (!at || (!pinned && !threadAt(doc.ws, at.row, at.col))) return undefined;
    const r = visibleRect({ r1: at.row, c1: at.col, r2: at.row, c2: at.col });
    if (!r) return undefined;
    const CARD_W = 282;
    const right = r.x + r.w + 8;
    const x = right + CARD_W > geo.width && r.x - CARD_W - 8 > geo.headerW ? r.x - CARD_W - 8 : right;
    return { row: at.row, col: at.col, x, y: Math.max(geo.headerH, r.y - 4), preview: !pinned };
  });

  // Excel closes the card once the selection leaves its cell.
  $effect(() => {
    const { row, col } = doc.selection.active;
    void doc.ws;
    untrack(() => {
      const c = ctl.commentCard;
      if (c && (c.row !== row || c.col !== col)) ctl.commentCard = null;
    });
  });

  const pick = $derived.by(() => {
    const p = ctl.pickList;
    if (!p) return undefined;
    const r = geo.rectOf({ r1: p.row, c1: p.col, r2: p.row, c2: p.col });
    return { x: r.x, y: r.y + r.h, w: Math.max(r.w, 120), items: p.items };
  });

  const fillButton = $derived.by(() => {
    const f = ctl.fillOptions;
    if (!f || f.version !== doc.version || ctl.edit) return undefined;
    const r = visibleRect(f.target);
    if (!r) return undefined;
    return { x: r.x + r.w + 2, y: r.y + r.h + 2 };
  });
  let fillMenuOpen = $state(false);
  const FILL_CHOICES: Array<[FillMode, MessageKey]> = [
    ['copy', 'fillCopyCells'],
    ['series', 'fillSeries'],
    ['formats', 'fillFormattingOnly'],
    ['values', 'fillWithoutFormatting'],
  ];
  function refill(mode: FillMode) {
    const f = ctl.fillOptions;
    fillMenuOpen = false;
    if (!f) return;
    // Excel swaps the fill in place: undo the previous fill, redo it the chosen way.
    doc.undo();
    autoFill(ctl, f.source, f.target, mode);
    ctl.fillOptions = { ...f, mode, version: doc.version };
  }

  let pickIndex = $state(0);
  $effect(() => {
    void ctl.pickList;
    pickIndex = 0;
  });

  function choose(item: string) {
    const p = ctl.pickList;
    if (!p) return;
    ctl.pickList = null;
    ctl.selectCell({ row: p.row, col: p.col });
    ctl.startEdit(item);
    ctl.commitEdit();
  }

  function pickKey(ev: KeyboardEvent) {
    const items = ctl.pickList?.items;
    if (!items) return;
    if (ev.key === 'ArrowDown') pickIndex = Math.min(items.length - 1, pickIndex + 1);
    else if (ev.key === 'ArrowUp') pickIndex = Math.max(0, pickIndex - 1);
    else if (ev.key === 'Enter') {
      const item = items[pickIndex];
      if (item !== undefined) choose(item);
    } else if (ev.key === 'Escape') ctl.pickList = null;
    else return;
    ev.preventDefault();
    ev.stopPropagation();
  }

  let pickEl: HTMLDivElement | undefined = $state();
  $effect(() => {
    if (pickEl) pickEl.focus();
  });
</script>

{#each filterButtons as b (`${b.row}:${b.col}`)}
  <button
    class="filter-btn"
    class:active={b.active}
    style:left="{b.x}px"
    style:top="{b.y}px"
    style:width="{b.size}px"
    style:height="{b.size}px"
    aria-label="Filter {cellAddress(b.row, b.col)}"
    onpointerdown={(e) => e.stopPropagation()}
    onclick={() => (ctl.filterMenu = { row: b.row, col: b.col })}
  >
    <svg viewBox="0 0 10 10" width="70%" height="70%" aria-hidden="true">
      {#if b.active}
        <path d="M1 2h8L6 5.5V9L4 8V5.5z" fill="currentColor" />
      {:else}
        <path d="M2 3.5h6L5 7z" fill="currentColor" />
      {/if}
    </svg>
  </button>
{/each}

{#if listArrow}
  <button
    class="list-arrow"
    style:left="{listArrow.x}px"
    style:top="{listArrow.y}px"
    style:height="{listArrow.h}px"
    aria-label="Show list"
    onpointerdown={(e) => e.stopPropagation()}
    onclick={() => ctl.openPickList()}
  >
    <svg viewBox="0 0 10 10" width="10" height="10" aria-hidden="true"><path d="M2 3.5h6L5 7z" fill="currentColor" /></svg>
  </button>
{/if}

{#if inputMessage && !pick}
  <div class="tip" style:left="{inputMessage.x}px" style:top="{inputMessage.y}px">
    {#if inputMessage.title}<strong>{inputMessage.title}</strong>{/if}
    {#if inputMessage.text}<div>{inputMessage.text}</div>{/if}
  </div>
{/if}

{#if commentCard}
  {#key `${commentCard.row}:${commentCard.col}:${commentCard.preview}`}
    <div class="comment-card" style:left="{commentCard.x}px" style:top="{commentCard.y}px">
      <CommentCard row={commentCard.row} col={commentCard.col} preview={commentCard.preview} onclose={() => (ctl.commentCard = null)} />
    </div>
  {/key}
{:else if note}
  <div class="note" style:left="{note.x}px" style:top="{note.y}px">
    {#if note.author}<strong>{note.author}:</strong>{/if}
    <div>{note.text}</div>
  </div>
{/if}

{#if pick}
  <div
    class="pick xl-menu"
    role="listbox"
    tabindex="-1"
    bind:this={pickEl}
    style:left="{pick.x}px"
    style:top="{pick.y}px"
    style:min-width="{pick.w}px"
    onkeydown={pickKey}
    onpointerdown={(e) => e.stopPropagation()}
    onfocusout={(e) => {
      if (!e.currentTarget.contains(e.relatedTarget as Node | null)) ctl.pickList = null;
    }}
  >
    {#each pick.items as item, i (i)}
      <div class="xl-menu-item" class:selected={i === pickIndex} role="option" aria-selected={i === pickIndex} tabindex="-1" onclick={() => choose(item)} onkeydown={() => {}}>{item}</div>
    {/each}
  </div>
{/if}

{#if fillButton}
  <button
    class="fill-options"
    style:left="{fillButton.x}px"
    style:top="{fillButton.y}px"
    aria-label={t('autoFillOptions')}
    title={t('autoFillOptions')}
    onpointerdown={(e) => e.stopPropagation()}
    onclick={() => (fillMenuOpen = !fillMenuOpen)}
  >
    <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><rect x="2" y="2" width="8" height="8" fill="none" stroke="currentColor" /><path d="M7 7h7v7H7z" fill="currentColor" opacity=".35" /></svg>
  </button>
  {#if fillMenuOpen}
    <div class="xl-menu fill-menu" role="menu" style:left="{fillButton.x}px" style:top="{fillButton.y + 22}px">
      {#each FILL_CHOICES as [mode, label] (mode)}
        <button class="xl-menu-item" role="menuitemradio" aria-checked={ctl.fillOptions?.mode === mode} onpointerdown={(e) => e.stopPropagation()} onclick={() => refill(mode)}>
          <span class="radio">{ctl.fillOptions?.mode === mode ? '●' : ''}</span>{t(label)}
        </button>
      {/each}
    </div>
  {/if}
{/if}

{#if ctl.filterMenu}
  {#key `${ctl.filterMenu.row}:${ctl.filterMenu.col}`}
    <FilterMenu at={ctl.filterMenu} />
  {/key}
{/if}

<style>
  .filter-btn,
  .list-arrow {
    position: absolute;
    display: grid;
    place-items: center;
    padding: 0;
    border: 1px solid #b5b5b5;
    border-radius: 2px;
    background: #f7f7f7;
    color: #444;
    cursor: default;
    z-index: 2;
  }
  .filter-btn.active {
    color: var(--xl-accent);
    border-color: var(--xl-accent);
  }
  .list-arrow {
    width: 16px;
  }
  .tip,
  .note {
    position: absolute;
    z-index: 3;
    max-width: 240px;
    padding: 4px 6px;
    font-size: 12px;
    white-space: pre-wrap;
    pointer-events: none;
    box-shadow: 1px 1px 3px rgb(0 0 0 / 0.25);
  }
  .tip {
    background: #ffffe1;
    border: 1px solid #767676;
  }
  .note {
    background: #ffffe1;
    border: 1px solid #8a8a8a;
    min-width: 140px;
    min-height: 60px;
  }
  .comment-card {
    position: absolute;
    z-index: 5;
  }
  .pick {
    position: absolute;
    z-index: 4;
    max-height: 220px;
    overflow: auto;
    outline: none;
  }
  .fill-options {
    position: absolute;
    z-index: 3;
    width: 22px;
    height: 20px;
    display: grid;
    place-items: center;
    padding: 0;
    border: 1px solid #b5b5b5;
    border-radius: 3px;
    background: #fff;
    color: #333;
  }
  .fill-menu {
    position: absolute;
    z-index: 4;
    min-width: 200px;
  }
  .fill-menu .radio {
    display: inline-block;
    width: 14px;
    font-size: 9px;
  }
  .pick .selected {
    background: var(--xl-accent-soft);
  }
</style>
