<script lang="ts">
  // The spreadsheet application shell: title bar with the file commands,
  // ribbon, formula bar, grid, sheet tabs and status bar, plus the overlays
  // (dialogs, context menu). File I/O lives here because it is the only part
  // that touches browser file pickers.
  import { onMount } from 'svelte';
  import './ui/tokens.css';
  import { EditorController, type EditorHost } from './core/controller.svelte.ts';
  import { setEditor } from './core/context.ts';
  import { blankWorkbook } from './core/editor.svelte.ts';
  import { printWorkbook } from './core/print.ts';
  import DialogHost from './dialogs/DialogHost.svelte';
  import GridView from './grid/GridView.svelte';
  import ChartsheetView from './ui/ChartsheetView.svelte';
  import { i18n, t } from './i18n/i18n.svelte.ts';
  import Ribbon from './ribbon/Ribbon.svelte';
  import CommentsPane from './ui/CommentsPane.svelte';
  import ContextMenu from './ui/ContextMenu.svelte';
  import ValidationAlert from './ui/ValidationAlert.svelte';
  import WatchWindow from './ui/WatchWindow.svelte';
  import NavigationPane from './ui/NavigationPane.svelte';
  import FormulaBar from './ui/FormulaBar.svelte';
  import PivotFieldsPane from './ui/PivotFieldsPane.svelte';
  import Icon from './ui/Icon.svelte';
  import SheetTabs from './ui/SheetTabs.svelte';
  import StatusBar from './ui/StatusBar.svelte';

  const ctl = new EditorController();
  setEditor(ctl);
  const doc = ctl.doc;

  let fileInput: HTMLInputElement;
  let dragging = $state(false);
  let fileMenu = $state(false);
  // File System Access handle of the opened file, so Save writes back in place where supported.
  let handle: FileSystemFileHandle | null = null;

  const XLSX_TYPES = [
    {
      description: 'Excel Workbook',
      accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'] },
    },
  ];

  type PickerTypes = typeof XLSX_TYPES;
  interface FilePickers {
    open?: (opts: { types: PickerTypes }) => Promise<FileSystemFileHandle[]>;
    save?: (opts: { suggestedName: string; types: PickerTypes }) => Promise<FileSystemFileHandle>;
  }

  // The File System Access API is Chromium-only and absent from TypeScript's DOM lib.
  function filePickers(): FilePickers {
    const w = window as Window & { showOpenFilePicker?: FilePickers['open']; showSaveFilePicker?: FilePickers['save'] };
    return {
      ...(w.showOpenFilePicker ? { open: w.showOpenFilePicker.bind(window) } : {}),
      ...(w.showSaveFilePicker ? { save: w.showSaveFilePicker.bind(window) } : {}),
    };
  }

  async function openFile(file: File): Promise<void> {
    if (!confirmDiscard()) return;
    try {
      await doc.open(file);
    } catch (err) {
      ctl.dialog = { kind: 'alert', props: { message: 'openFailed', detail: err instanceof Error ? err.message : String(err) } };
    }
  }

  function confirmDiscard(): boolean {
    return !doc.dirty || window.confirm(t('discardChanges'));
  }

  async function pickAndOpen(): Promise<void> {
    const picker = filePickers();
    if (picker.open) {
      let handles: FileSystemFileHandle[];
      try {
        handles = await picker.open({ types: XLSX_TYPES });
      } catch {
        // The user dismissed the picker.
        return;
      }
      const h = handles[0];
      if (!h) return;
      await openFile(await h.getFile());
      handle = h;
      return;
    }
    fileInput.click();
  }

  async function writeTo(h: FileSystemFileHandle): Promise<void> {
    const bytes = await doc.toBytes();
    const w = await h.createWritable();
    await w.write(new Blob([bytes.slice()]));
    await w.close();
    doc.fileName = h.name;
    doc.dirty = false;
  }

  async function download(): Promise<void> {
    const bytes = await doc.toBytes();
    const blob = new Blob([bytes.slice()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = doc.fileName.endsWith('.xlsx') || doc.fileName.endsWith('.xlsm') ? doc.fileName : `${doc.fileName}.xlsx`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    doc.dirty = false;
  }

  async function saveAs(): Promise<void> {
    const picker = filePickers().save;
    if (!picker) return download();
    let h: FileSystemFileHandle;
    try {
      h = await picker({ suggestedName: doc.fileName, types: XLSX_TYPES });
    } catch {
      return;
    }
    await writeTo(h);
    handle = h;
  }

  async function save(): Promise<void> {
    if (ctl.edit && !ctl.commitEdit()) return;
    try {
      if (handle) await writeTo(handle);
      else await saveAs();
    } catch (err) {
      ctl.dialog = { kind: 'alert', props: { message: 'saveFailed', detail: err instanceof Error ? err.message : String(err) } };
    }
  }

  const host: EditorHost = {
    open: () => void pickAndOpen(),
    save: () => void save(),
    saveAs: () => void saveAs(),
    newWorkbook: () => {
      if (!confirmDiscard()) return;
      handle = null;
      doc.replaceWorkbook(blankWorkbook(), nextBookName());
    },
    print: () => printWorkbook(ctl),
  };
  ctl.host = host;

  let bookCounter = 1;
  function nextBookName(): string {
    bookCounter++;
    return `Book${bookCounter}.xlsx`;
  }

  function fileMenuRun(fn: () => void) {
    fileMenu = false;
    fn();
  }

  onMount(() => {
    // Lets browser automation drive and inspect the editor during development.
    if (import.meta.env.DEV) Object.assign(window, { __xlEditor: ctl });
    const beforeUnload = (e: BeforeUnloadEvent) => {
      if (doc.dirty) e.preventDefault();
    };
    window.addEventListener('beforeunload', beforeUnload);
    return () => window.removeEventListener('beforeunload', beforeUnload);
  });

  $effect(() => {
    document.title = `${doc.dirty ? '• ' : ''}${doc.fileName} — xlsx editor`;
  });

  function onDrop(e: DragEvent) {
    e.preventDefault();
    dragging = false;
    const file = e.dataTransfer?.files[0];
    if (file) void openFile(file);
  }

  // Excel's toolbars never take the keyboard from the sheet: pressing a button
  // must leave arrows and typing going to the grid. Dialogs keep their own focus.
  function keepGridFocus(e: MouseEvent): void {
    if (!(e.target instanceof Element)) return;
    if (e.target.closest('input, textarea, select, [contenteditable], [role="dialog"]')) return;
    if (e.target.closest('button, [role="menuitem"], [role="option"]')) e.preventDefault();
  }

  // A menu that closes takes its focused item with it; hand the keyboard back to the grid.
  function regainFocus(e: FocusEvent): void {
    if (e.relatedTarget !== null) return;
    queueMicrotask(() => {
      if (document.activeElement === document.body && !ctl.dialog) ctl.gridFocusRequest++;
    });
  }

  // After a command runs, the grid takes the keyboard again (unless it opened a dialog or a drop-down).
  function refocusAfterCommand(e: MouseEvent): void {
    if (!(e.target instanceof Element)) return;
    const button = e.target.closest('button, [role="menuitem"]');
    if (!button || button.closest('[role="dialog"], input, textarea, select')) return;
    if (button.hasAttribute('aria-expanded') || button.hasAttribute('aria-haspopup')) return;
    queueMicrotask(() => {
      if (!ctl.dialog && !ctl.menu) ctl.gridFocusRequest++;
    });
  }

  let root = $state<HTMLDivElement>();
  $effect(() => {
    const node = root;
    if (!node) return;
    node.addEventListener('mousedown', keepGridFocus);
    node.addEventListener('focusout', regainFocus);
    node.addEventListener('click', refocusAfterCommand);
    return () => {
      node.removeEventListener('click', refocusAfterCommand);
      node.removeEventListener('mousedown', keepGridFocus);
      node.removeEventListener('focusout', regainFocus);
    };
  });
</script>

<div
  class="xl-editor app"
  role="application"
  bind:this={root}
  ondragover={(e) => {
    if (e.dataTransfer?.types.includes('Files')) {
      e.preventDefault();
      dragging = true;
    }
  }}
  ondragleave={(e) => {
    if (e.currentTarget === e.target) dragging = false;
  }}
  ondrop={onDrop}
>
  <header class="titlebar">
    <div class="file">
      <button class="xl-btn file-btn" onclick={() => (fileMenu = !fileMenu)} aria-expanded={fileMenu}>{t('file')}</button>
      {#if fileMenu}
        <div class="xl-menu file-menu" role="menu">
          <button class="xl-menu-item" onclick={() => fileMenuRun(host.newWorkbook)}>{t('newWorkbook')}<span class="shortcut">⌘N</span></button>
          <button class="xl-menu-item" onclick={() => fileMenuRun(host.open)}>{t('openEllipsis')}<span class="shortcut">⌘O</span></button>
          <div class="xl-menu-sep"></div>
          <button class="xl-menu-item" onclick={() => fileMenuRun(host.save)}>{t('save')}<span class="shortcut">⌘S</span></button>
          <button class="xl-menu-item" onclick={() => fileMenuRun(host.saveAs)}>{t('saveAsEllipsis')}<span class="shortcut">⇧⌘S</span></button>
          <button class="xl-menu-item" onclick={() => fileMenuRun(() => void download())}>{t('downloadCopy')}</button>
          <div class="xl-menu-sep"></div>
          <button class="xl-menu-item" onclick={() => fileMenuRun(host.print)}>{t('printEllipsis')}<span class="shortcut">⌘P</span></button>
          <div class="xl-menu-sep"></div>
          <p class="disclaimer">{t('notAffiliated')}</p>
        </div>
      {/if}
    </div>
    <div class="quick">
      <button class="xl-btn icon" title={t('save')} onclick={host.save}><Icon name="save" /></button>
      <button class="xl-btn icon" title={t('undo')} disabled={!doc.canUndo} onclick={() => doc.undo()}><Icon name="undo" /></button>
      <button class="xl-btn icon" title={t('redo')} disabled={!doc.canRedo} onclick={() => doc.redo()}><Icon name="redo" /></button>
    </div>
    <div class="name">{doc.fileName}{doc.dirty ? ` — ${t('edited')}` : ''}</div>
    <select class="xl-input lang" bind:value={i18n.locale} aria-label={t('language')}>
      <option value="en">English</option>
      <option value="ja">日本語</option>
    </select>
  </header>

  <Ribbon />
  {#if ctl.showFormulaBar && ctl.shownChartsheet === null}
    <FormulaBar />
  {/if}
  <div class="workspace">
    <div class="grid">
      {#if ctl.shownChartsheet !== null}
        <ChartsheetView index={ctl.shownChartsheet} />
      {:else}
        <GridView />
      {/if}
    </div>
    {#if ctl.pivotFieldList && ctl.activePivot}
      <PivotFieldsPane />
    {/if}
    {#if ctl.commentsPane}
      <CommentsPane />
    {/if}
    {#if ctl.navigationPane}
      <NavigationPane />
    {/if}
  </div>
  <SheetTabs />
  <StatusBar />

  <DialogHost />
  <ContextMenu />
  <ValidationAlert />
  {#if ctl.watchWindow}<WatchWindow />{/if}

  {#if dragging}
    <div class="drop">{t('dropToOpen')}</div>
  {/if}

  <input
    bind:this={fileInput}
    type="file"
    accept=".xlsx,.xlsm"
    hidden
    onchange={(e) => {
      const file = e.currentTarget.files?.[0];
      e.currentTarget.value = '';
      if (file) {
        handle = null;
        void openFile(file);
      }
    }}
  />
</div>

<style>
  .app {
    position: relative;
    display: flex;
    flex-direction: column;
    height: 100%;
    min-height: 0;
    overflow: hidden;
    background: var(--xl-bg);
  }

  .titlebar {
    display: flex;
    align-items: center;
    gap: 8px;
    height: 34px;
    padding: 0 8px;
    background: var(--xl-accent);
    color: #fff;
    flex: none;
  }

  .file {
    position: relative;
  }

  .file-btn {
    color: #fff;
    font-weight: 600;
  }

  .file-menu {
    position: absolute;
    top: 100%;
    left: 0;
    min-width: 220px;
    z-index: 50;
    color: var(--xl-text);
  }
  .disclaimer {
    max-width: 220px;
    margin: 0;
    padding: 4px 12px 6px;
    font-size: 11px;
    line-height: 1.4;
    color: var(--xl-text-3);
    white-space: normal;
  }

  .quick {
    display: flex;
    gap: 2px;
  }

  .quick .xl-btn {
    color: #fff;
  }

  .name {
    flex: 1;
    text-align: center;
    font-size: 13px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .lang {
    height: 22px;
    font-size: 12px;
  }

  .workspace {
    display: flex;
    flex: 1;
    min-height: 0;
  }

  .grid {
    position: relative;
    min-width: 0;
    display: flex;
    flex-direction: column;
    flex: 1;
    min-height: 0;
  }

  .drop {
    position: absolute;
    inset: 0;
    display: grid;
    place-items: center;
    background: color-mix(in srgb, var(--xl-accent) 18%, transparent);
    border: 3px dashed var(--xl-accent);
    font-size: 20px;
    pointer-events: none;
    z-index: 100;
  }
</style>
