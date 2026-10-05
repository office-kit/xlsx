<script lang="ts">
  // Right-click menus for cells, row/column headers and sheet tabs, with the
  // items Excel shows in each.
  import { onMount } from 'svelte';
  import * as A from '../core/actions.ts';
  import { deleteCommentsInSelection, newComment, threadAt } from '../core/comments.ts';
  import { getEditor } from '../core/context.ts';
  import { quickSort } from '../core/data.ts';
  import { clearAdvancedFilter } from '../core/advanced-filter.ts';
  import { clearAllFilters, filterBySelectedValue, filterOwners, reapplyFilters } from '../core/filter.ts';
  import type { PasteMode } from '../core/clipboard.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import ColorGrid from './ColorGrid.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  let el: HTMLDivElement | undefined = $state();
  let pos = $state({ x: 0, y: 0 });
  let showTabColors = $state(false);
  // The open submenu (Paste Special ▸, Filter ▸, Sort ▸), shown beside its item on hover as in Excel.
  let sub = $state<'paste' | 'filter' | 'sort' | null>(null);

  const m = $derived(ctl.menu);

  // Each menu opens at its top level, however the last one was dismissed.
  $effect.pre(() => {
    void ctl.menu;
    showTabColors = false;
    sub = null;
  });

  $effect(() => {
    const menu = ctl.menu;
    const node = el;
    if (!menu || !node) return;
    queueMicrotask(() => {
      const r = node.getBoundingClientRect();
      pos = { x: Math.min(menu.x, window.innerWidth - r.width - 4), y: Math.min(menu.y, window.innerHeight - r.height - 4) };
    });
  });

  onMount(() => {
    const down = (e: PointerEvent) => {
      if (el && !el.contains(e.target as Node)) ctl.menu = null;
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && ctl.menu) ctl.menu = null;
    };
    window.addEventListener('pointerdown', down, true);
    window.addEventListener('keydown', key, true);
    return () => {
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('keydown', key, true);
    };
  });

  // Run before closing: the menu's `{@const}` values (the sheet a tab menu
  // is for) read the open menu, and are gone once it is null.
  function run(fn: () => void) {
    fn();
    ctl.menu = null;
    showTabColors = false;
    sub = null;
  }

  const hasNote = $derived.by(() => {
    void doc.version;
    const { row, col } = doc.selection.active;
    return doc.ws.legacyComments.some((c) => c.ref.replaceAll('$', '') === `${A.colName(col)}${row}`);
  });
  const hasThread = $derived.by(() => {
    void doc.version;
    const { row, col } = doc.selection.active;
    return threadAt(doc.ws, row, col) !== undefined;
  });
  const hasFilter = $derived.by(() => {
    void doc.version;
    return filterOwners(doc.ws).length > 0;
  });
  // Excel's Paste Special ▸ items, in its order, for the paste modes the editor has.
  const PASTE_ITEMS: ReadonlyArray<[PasteMode, MessageKey]> = [
    ['all', 'paste'],
    ['formulas', 'pasteFormulas'],
    ['noBorders', 'pasteNoBorders'],
    ['columnWidths', 'pasteColumnWidths'],
    ['transpose', 'pasteTranspose'],
    ['values', 'pasteValues'],
    ['valuesAndFormats', 'pasteValuesNumberFormats'],
    ['formats', 'pasteFormatting'],
  ];
  const hasLink = $derived.by(() => {
    void doc.version;
    const { row, col } = doc.selection.active;
    return doc.ws.hyperlinks.some((h) => h.ref.split(':')[0] === `${A.colName(col)}${row}`);
  });
</script>

{#snippet submenu(id: 'paste' | 'filter' | 'sort', label: string)}
  <div class="sub-host" role="none" onpointerenter={() => (sub = id)}>
    <button class="xl-menu-item" aria-haspopup="menu" aria-expanded={sub === id} onclick={() => (sub = id)}>{label}<span class="shortcut">▸</span></button>
    {#if sub === id}
      <div class="xl-menu sub" role="menu">
        {#if id === 'paste'}
          {#each PASTE_ITEMS as [mode, key] (key)}
            <button class="xl-menu-item" onclick={() => run(() => A.paste(ctl, mode))}>{t(key)}</button>
          {/each}
          <div class="xl-menu-sep"></div>
          <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('pasteSpecial'))}>{t('pasteSpecialEllipsis')}<span class="shortcut">⌃⌘V</span></button>
        {:else if id === 'filter'}
          <button class="xl-menu-item" disabled={!hasFilter} onclick={() => run(() => { clearAllFilters(ctl); clearAdvancedFilter(ctl); })}>{t('ctxClearFilter')}</button>
          <button class="xl-menu-item" disabled={!hasFilter} onclick={() => run(() => reapplyFilters(ctl))}>{t('reapply')}</button>
          <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('advancedFilter'))}>{t('ctxAdvancedFilter')}</button>
          <div class="xl-menu-sep"></div>
          <button class="xl-menu-item" onclick={() => run(() => filterBySelectedValue(ctl))}>{t('ctxFilterByValue')}</button>
        {:else}
          <button class="xl-menu-item" onclick={() => run(() => quickSort(ctl, false))}>{t('sortAZ')}</button>
          <button class="xl-menu-item" onclick={() => run(() => quickSort(ctl, true))}>{t('sortZA')}</button>
          <div class="xl-menu-sep"></div>
          <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('sort'))}>{t('customSortEllipsis')}</button>
        {/if}
      </div>
    {/if}
  </div>
{/snippet}

{#if m}
  <div class="xl-menu" bind:this={el} style:left="{pos.x || m.x}px" style:top="{pos.y || m.y}px" role="menu">
    {#if m.kind === 'sheetTab'}
      {@const i = m.sheetIndex ?? doc.activeSheetIndex}
      {#if showTabColors}
        <ColorGrid palette={doc.styles.palette} noneLabel={t('noColor')} onpick={(rgb) => run(() => A.setTabColor(ctl, i, rgb))} />
      {:else}
        <button class="xl-menu-item" onclick={() => run(() => A.insertSheet(ctl, i))}>{t('insertSheet')}</button>
        <button class="xl-menu-item" onclick={() => run(() => A.deleteSheet(ctl, i))}>{t('delete')}</button>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('renameSheet', { index: i }))}>{t('ctxRename')}</button>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('moveCopySheet', { index: i }))}>{t('ctxMoveCopy')}</button>
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('protectSheet'))}>{t('protectSheetEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => (showTabColors = true)}>{t('tabColor')} ▸</button>
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => run(() => A.setSheetHidden(ctl, i, true))}>{t('hide')}</button>
        <button class="xl-menu-item" disabled={!doc.wb.sheets.some((s) => s.state !== 'visible')} onclick={() => run(() => ctl.openDialog('unhideSheet'))}>{t('unhideEllipsis')}</button>
      {/if}
    {:else}
      <button class="xl-menu-item" onclick={() => run(() => A.cut(ctl))}>{t('cut')}<span class="shortcut">⌘X</span></button>
      <button class="xl-menu-item" onclick={() => run(() => A.copy(ctl))}>{t('copy')}<span class="shortcut">⌘C</span></button>
      <button class="xl-menu-item" onclick={() => run(() => A.paste(ctl))}>{t('paste')}<span class="shortcut">⌘V</span></button>
      {@render submenu('paste', t('pasteSpecial'))}
      <div class="xl-menu-sep"></div>
      {#if m.kind === 'colHeader'}
        <button class="xl-menu-item" onclick={() => run(() => A.insertLines(ctl, 'col'))}>{t('insert')}</button>
        <button class="xl-menu-item" onclick={() => run(() => A.deleteLines(ctl, 'col'))}>{t('delete')}</button>
        <button class="xl-menu-item" onclick={() => run(() => A.clear(ctl, 'contents'))}>{t('clearContents')}</button>
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('formatCells'))}>{t('formatCellsEllipsis')}<span class="shortcut">⌘1</span></button>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('columnWidth'))}>{t('columnWidthEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => run(() => A.hideLines(ctl, 'col', true))}>{t('hide')}<span class="shortcut">⌃0</span></button>
        <button class="xl-menu-item" onclick={() => run(() => A.hideLines(ctl, 'col', false))}>{t('unhide')}<span class="shortcut">⇧⌃0</span></button>
      {:else if m.kind === 'rowHeader'}
        <button class="xl-menu-item" onclick={() => run(() => A.insertLines(ctl, 'row'))}>{t('insert')}</button>
        <button class="xl-menu-item" onclick={() => run(() => A.deleteLines(ctl, 'row'))}>{t('delete')}</button>
        <button class="xl-menu-item" onclick={() => run(() => A.clear(ctl, 'contents'))}>{t('clearContents')}</button>
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('formatCells'))}>{t('formatCellsEllipsis')}<span class="shortcut">⌘1</span></button>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('rowHeight'))}>{t('rowHeightEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => run(() => A.hideLines(ctl, 'row', true))}>{t('hide')}<span class="shortcut">⌃9</span></button>
        <button class="xl-menu-item" onclick={() => run(() => A.hideLines(ctl, 'row', false))}>{t('unhide')}<span class="shortcut">⇧⌃9</span></button>
      {:else}
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('insertCells'))}>{t('insertEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('deleteCells'))}>{t('deleteEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => run(() => A.clear(ctl, 'contents'))}>{t('clearContents')}</button>
        <div class="xl-menu-sep"></div>
        {@render submenu('filter', t('filter'))}
        {@render submenu('sort', t('sort'))}
        <div class="xl-menu-sep"></div>
        {#if hasThread}
          <button class="xl-menu-item" onclick={() => run(() => newComment(ctl))}>{t('cmtReplyToComment')}</button>
          <button class="xl-menu-item" onclick={() => run(() => deleteCommentsInSelection(ctl))}>{t('cmtDeleteComment')}</button>
        {:else if !hasNote}
          <button class="xl-menu-item" onclick={() => run(() => newComment(ctl))}>{t('cmtNewComment')}</button>
        {/if}
        {#if hasNote}
          <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('note'))}>{t('editNote')}</button>
          <button class="xl-menu-item" onclick={() => run(() => A.deleteNotes(ctl))}>{t('deleteNote')}</button>
        {:else}
          <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('note'))}>{t('newNote')}</button>
        {/if}
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('formatCells'))}>{t('formatCellsEllipsis')}<span class="shortcut">⌘1</span></button>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openPickList())}>{t('pickFromList')}</button>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('defineName'))}>{t('defineNameEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => run(() => ctl.openDialog('hyperlink'))}>{hasLink ? t('editHyperlink') : t('ctxHyperlink')}<span class="shortcut">⌘K</span></button>
        {#if hasLink}
          <button class="xl-menu-item" onclick={() => run(() => A.clear(ctl, 'removeHyperlinks'))}>{t('removeHyperlink')}</button>
        {/if}
      {/if}
    {/if}
  </div>
{/if}

<style>
  .sub-host {
    position: relative;
  }
  .sub {
    position: absolute;
    left: 100%;
    top: -4px;
    margin-left: 2px;
  }
</style>
