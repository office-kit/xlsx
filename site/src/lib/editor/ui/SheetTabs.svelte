<script lang="ts">
  // Sheet tab strip: click to activate, double-click to rename in place,
  // drag to reorder (Option/Alt-drag copies), right-click for the sheet menu,
  // "+" to insert a sheet.
  import * as A from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { resolveColor } from '../core/theme.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from './Icon.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  let renaming = $state<number | null>(null);
  let renameText = $state('');
  let dragFrom = $state<number | null>(null);
  let dropAt = $state<number | null>(null);
  let strip: HTMLDivElement;

  const tabs = $derived.by(() => {
    void doc.version;
    return doc.wb.sheets.map((s, index) => ({
      index,
      title: s.sheet.title,
      visible: s.state === 'visible',
      chart: s.kind === 'chartsheet',
      color: s.kind === 'worksheet' ? s.sheet.sheetProperties?.tabColor : undefined,
    }));
  });

  function activate(i: number) {
    // Pointing at a reference while editing a formula stays on worksheets.
    if (ctl.edit && ctl.canPoint(true)) {
      if (doc.wb.sheets[i]?.kind === 'worksheet') doc.activateSheet(i);
      ctl.gridFocusRequest++;
      return;
    }
    if (ctl.edit && !ctl.commitEdit()) return;
    ctl.showSheet(i);
    // Keys go to the sheet just opened, not to the tab button, as in Excel.
    ctl.gridFocusRequest++;
  }

  function startRename(i: number) {
    renaming = i;
    renameText = doc.wb.sheets[i]?.sheet.title ?? '';
  }

  function finishRename() {
    if (renaming === null) return;
    const i = renaming;
    renaming = null;
    const err = A.renameSheetAt(ctl, i, renameText);
    if (err) ctl.dialog = { kind: 'alert', props: { message: err } };
  }

  function focusSelect(node: HTMLInputElement) {
    node.focus();
    node.select();
  }

  function onDragStart(e: DragEvent, i: number) {
    dragFrom = i;
    e.dataTransfer?.setData('text/plain', String(i));
  }

  function onDragOver(e: DragEvent, i: number) {
    if (dragFrom === null) return;
    e.preventDefault();
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    dropAt = e.clientX < rect.left + rect.width / 2 ? i : i + 1;
  }

  function onDrop(e: DragEvent) {
    e.preventDefault();
    if (dragFrom === null || dropAt === null) return;
    const to = dropAt > dragFrom ? dropAt - 1 : dropAt;
    if (e.altKey || e.ctrlKey) A.duplicateSheet(ctl, dragFrom, dropAt);
    else A.moveSheetTo(ctl, dragFrom, to);
    dragFrom = null;
    dropAt = null;
  }
</script>

<div class="sheetbar">
  <div class="nav">
    <button class="xl-btn" title={t('previousSheet')} onclick={() => ctl.switchSheet(-1)}><Icon name="chevron-left" size={14} /></button>
    <button class="xl-btn" title={t('nextSheet')} onclick={() => ctl.switchSheet(1)}><Icon name="chevron-right" size={14} /></button>
  </div>
  <div class="tabs" bind:this={strip} role="tablist" tabindex="-1" ondragover={(e) => e.preventDefault()} ondrop={onDrop}>
    {#each tabs.filter((x) => x.visible) as tab (tab.index)}
      <div
        class="tab"
        class:active={tab.index === (ctl.shownChartsheet ?? doc.activeSheetIndex)}
        class:drop-before={dropAt === tab.index}
        role="tab"
        tabindex="-1"
        aria-selected={tab.index === (ctl.shownChartsheet ?? doc.activeSheetIndex)}
        draggable={renaming !== tab.index}
        ondragstart={(e) => onDragStart(e, tab.index)}
        ondragover={(e) => onDragOver(e, tab.index)}
        ondragend={() => { dragFrom = null; dropAt = null; }}
        onclick={() => activate(tab.index)}
        ondblclick={() => startRename(tab.index)}
        oncontextmenu={(e) => { e.preventDefault(); activate(tab.index); ctl.menu = { x: e.clientX, y: e.clientY, kind: 'sheetTab', sheetIndex: tab.index }; }}
        onkeydown={() => {}}
        style:--tab-color={tab.color ? resolveColor(tab.color, doc.styles.palette, 'transparent') : 'transparent'}
      >
        {#if renaming === tab.index}
          <input class="rename" bind:value={renameText} use:focusSelect onblur={finishRename} onkeydown={(e) => { if (e.key === 'Enter') finishRename(); if (e.key === 'Escape') renaming = null; e.stopPropagation(); }} />
        {:else}
          {tab.title}
        {/if}
      </div>
    {/each}
    <button class="xl-btn add" title={t('newSheet')} onclick={() => A.insertSheet(ctl, doc.wb.sheets.length)}><Icon name="plus" size={14} /></button>
  </div>
</div>

<style>
  .sheetbar {
    display: flex;
    align-items: stretch;
    height: 28px;
    background: var(--xl-bg);
    border-top: 1px solid var(--xl-border);
    flex-shrink: 0;
    min-width: 0;
  }
  .nav {
    display: flex;
    align-items: center;
    padding: 0 4px;
  }
  .tabs {
    display: flex;
    align-items: stretch;
    overflow-x: auto;
    min-width: 0;
    flex: 1;
    scrollbar-width: none;
  }
  .tab {
    display: flex;
    align-items: center;
    min-width: 64px;
    justify-content: center;
    padding: 0 14px;
    border-right: 1px solid var(--xl-border);
    cursor: pointer;
    white-space: nowrap;
    user-select: none;
    color: var(--xl-text-2);
    box-shadow: inset 0 -3px 0 var(--tab-color);
  }
  .tab:hover {
    background: var(--xl-hover);
  }
  /* The active sheet is a raised white tab with accent-coloured, non-bold text. */
  .tab.active {
    background: #fff;
    color: var(--xl-accent);
    margin: 2px 0 3px;
    border-right-color: transparent;
    border-radius: 3px;
    box-shadow:
      0 0 0 1px var(--xl-border-strong),
      inset 0 -3px 0 var(--tab-color);
  }
  .tab.drop-before {
    box-shadow: inset 3px 0 0 var(--xl-accent);
  }
  .rename {
    font: inherit;
    width: 110px;
    border: 1px solid var(--xl-accent);
    padding: 1px 4px;
  }
  .add {
    margin: 2px 6px;
  }
</style>
