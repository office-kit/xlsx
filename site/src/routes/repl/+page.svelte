<script lang="ts">
  import { onMount } from 'svelte';
  import SheetGrid from '$lib/components/SheetGrid.svelte';
  import starterSource from '$lib/examples/repl-default.ts?raw';
  import type { ReplEditor } from '$lib/repl/editor';
  import type { ReplResult } from '$lib/repl/runtime';

  // Above the marker are the imports that let svelte-check compile the starter
  // code; the REPL shows and runs what comes after it.
  const START_MARKER = '// repl:start\n';
  const DEFAULT_CODE = starterSource.slice(starterSource.indexOf(START_MARKER) + START_MARKER.length);

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const DOWNLOAD_NAME = 'office-kit-xlsx-repl.xlsx';
  // Long enough that a burst of keystrokes is one run, short enough to feel live.
  const RUN_DEBOUNCE_MS = 250;
  const COPY_NOTE_MS = 1500;

  type Runtime = typeof import('$lib/repl/runtime');

  let code = $state(DEFAULT_CODE);
  let error = $state('');
  let result = $state.raw<ReplResult | undefined>();
  let busy = $state(true);
  let copyStatus = $state<'idle' | 'copied' | 'failed'>('idle');
  // eslint-disable-next-line prefer-const -- reassigned by `bind:this` in template
  let editorContainer = $state<HTMLDivElement | undefined>();

  // Both load on mount: prerendering never evaluates user code, and the
  // library and CodeMirror stay out of the route's initial JavaScript.
  let runtime = $state.raw<Runtime | undefined>();
  let editor: ReplEditor | undefined;

  const clipped = $derived(result?.sheets.filter((sheet) => sheet.clipped) ?? []);

  onMount(() => {
    let disposed = false;
    void (async () => {
      try {
        const [editorModule, runtimeModule] = await Promise.all([
          import('$lib/repl/editor'),
          import('$lib/repl/runtime'),
        ]);
        if (disposed || !editorContainer) return;
        editor = editorModule.createEditor(editorContainer, code, (text) => (code = text));
        runtime = runtimeModule;
      } catch (err) {
        console.error(err);
        error = 'The editor or the library failed to load. Reload the page to try again.';
        busy = false;
      }
    })();
    return () => {
      disposed = true;
      editor?.destroy();
    };
  });

  // Runs overlap when saving takes longer than the debounce; only the newest
  // one may write to the page.
  let latestRun = 0;

  async function run(using: Runtime, source: string): Promise<void> {
    const id = ++latestRun;
    busy = true;
    try {
      const next = await using.runCode(source);
      if (id !== latestRun) return;
      result = next;
      error = '';
    } catch (err) {
      if (id !== latestRun) return;
      // Code is broken for most keystrokes while someone types. Keep the last
      // workbook that built on screen (dimmed) instead of blanking the preview.
      error = describe(using, err);
    } finally {
      if (id === latestRun) busy = false;
    }
  }

  function describe(using: Runtime, err: unknown): string {
    if (!(err instanceof using.ReplError)) return String(err);
    const line = err.line ?? (err.syntax ? editor?.firstSyntaxErrorLine() : undefined);
    return line === undefined ? err.message : `Line ${line}: ${err.message}`;
  }

  $effect(() => {
    const source = code;
    const using = runtime;
    if (!using) return;
    const timer = setTimeout(() => void run(using, source), RUN_DEBOUNCE_MS);
    return () => clearTimeout(timer);
  });

  function download(): void {
    if (!result) return;
    const url = URL.createObjectURL(new Blob([result.bytes.slice()], { type: XLSX_MIME }));
    const a = document.createElement('a');
    a.href = url;
    a.download = DOWNLOAD_NAME;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function copyCode(): Promise<void> {
    try {
      await navigator.clipboard.writeText(code);
      copyStatus = 'copied';
    } catch {
      // The browser refused clipboard access; say so instead of pretending.
      copyStatus = 'failed';
    }
    setTimeout(() => (copyStatus = 'idle'), COPY_NOTE_MS);
  }

  function resetCode(): void {
    code = DEFAULT_CODE;
    // The editor owns its document, so the state change alone would leave the
    // old text on screen.
    editor?.setText(DEFAULT_CODE);
  }
</script>

<svelte:head>
  <title>REPL · @office-kit/xlsx</title>
</svelte:head>

<section class="content">
  <header class="intro">
    <h1>REPL</h1>
    <p class="lede">
      Write code and the workbook redraws as you type. The library’s functions are in scope and
      <code>wb</code> is an empty workbook from <code>createWorkbook()</code>. Nothing leaves this tab:
      the code runs here, <code>workbookToBytes</code> saves the result, <code>loadWorkbook</code> reads
      those bytes back, and the grid is drawn from what it read. The download is the same bytes.
    </p>
  </header>

  <div class="repl-grid">
    <div class="pane editor-pane">
      <div class="pane-head">
        <h2>Code</h2>
        <div class="pane-actions">
          <button type="button" onclick={resetCode}>Reset</button>
          <button type="button" onclick={copyCode}>{copyStatus === 'copied' ? 'Copied' : copyStatus === 'failed' ? 'Copy failed' : 'Copy'}</button>
        </div>
      </div>
      <div class="editor" bind:this={editorContainer}></div>
      {#if error}
        <pre class="error" role="alert">{error}</pre>
      {/if}
    </div>

    <div class="pane preview-pane">
      <div class="pane-head">
        <h2>Preview</h2>
        <div class="pane-actions">
          <span class="busy" aria-live="polite">{busy ? 'Building…' : ''}</span>
          <button type="button" class="strong" onclick={download} disabled={!result}>
            Download .xlsx
          </button>
        </div>
      </div>
      <div class="preview" class:stale={error !== ''}>
        {#if result && result.sheets.length > 0}
          <SheetGrid sheets={result.sheets} />
        {:else if result}
          <p class="empty">
            The workbook has no worksheets yet. Add one with <code>addWorksheet(wb, 'Sheet1')</code>.
          </p>
        {/if}
      </div>
      <p class="note">
        {#if clipped.length > 0}
          Showing the first {runtime?.PREVIEW_LIMIT.maxRows} rows and {runtime?.PREVIEW_LIMIT.maxCols}
          columns of {clipped.map((sheet) => sheet.name).join(', ')}; the download has everything.
        {:else}
          The grid draws values, number formats, fonts, fills, borders, merges, freeze panes, and
          column widths. Conditional formats, data validation, and comments are in the download, not
          the grid.
        {/if}
      </p>
    </div>
  </div>
</section>

<style>
  .content {
    max-width: 1440px;
    margin: 0 auto;
    padding: 2.25rem var(--gutter) 3rem;
  }

  .intro h1 {
    margin-bottom: 0.6rem;
  }

  .lede {
    max-width: 78ch;
    margin: 0 0 1.5rem;
    color: var(--ink-2);
  }

  .repl-grid {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    gap: 1rem;
    align-items: start;
  }

  .pane {
    display: flex;
    flex-direction: column;
    min-width: 0;
    height: calc(100vh - var(--header-h) - 2rem);
    min-height: 480px;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius);
    overflow: hidden;
  }

  .editor-pane {
    position: sticky;
    top: calc(var(--header-h) + 1rem);
    background: var(--night);
    border-color: var(--night-line);
  }

  .pane-head {
    flex: none;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.75rem;
    min-height: 48px;
    padding: 0 0.6rem 0 1rem;
    border-bottom: 1px solid var(--line);
    background: var(--wash);
  }

  .editor-pane .pane-head {
    background: var(--night);
    border-bottom-color: var(--night-line);
    color: var(--night-ink);
  }

  .pane-head h2 {
    margin: 0;
    font-family: var(--sans);
    font-size: 0.9rem;
    font-weight: 600;
    letter-spacing: 0;
    color: inherit;
  }

  .pane-actions {
    display: flex;
    align-items: center;
    gap: 0.4rem;
  }

  .pane-actions button {
    height: 32px;
    padding: 0 0.75rem;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius-sm);
    background: var(--paper);
    color: var(--ink);
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 550;
    cursor: pointer;
  }

  .pane-actions button:hover:not(:disabled) {
    border-color: var(--ink-3);
  }

  .pane-actions button:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }

  .editor-pane .pane-actions button {
    background: var(--night-2);
    border-color: var(--night-line);
    color: var(--night-ink);
  }

  .editor-pane .pane-actions button:hover {
    border-color: var(--night-ink-2);
  }

  .pane-actions button.strong {
    background: var(--ink);
    border-color: var(--ink);
    color: var(--paper);
  }

  .pane-actions button.strong:hover:not(:disabled) {
    background: var(--accent);
    border-color: var(--accent);
    color: var(--on-accent);
  }

  .busy {
    color: var(--ink-3);
    font-size: 0.82rem;
  }

  .editor {
    flex: 1;
    min-height: 0;
    overflow: hidden;
  }

  .editor :global(.cm-editor) {
    height: 100%;
  }

  /* The error sits on the editor's night surface in both themes, so its
   * colours are fixed like the syntax colours above it. */
  .error {
    --error-rule: #7a2a1a;
    --error-surface: #2a1410;
    --error-ink: #ffb4a1;

    flex: none;
    max-height: 35%;
    margin: 0;
    border: none;
    border-top: 1px solid var(--error-rule);
    border-radius: 0;
    background: var(--error-surface);
    color: var(--error-ink);
    font-size: 0.8rem;
    white-space: pre-wrap;
    overflow: auto;
  }

  .preview {
    flex: 1;
    min-height: 0;
    background: var(--wash);
    transition: opacity 120ms ease;
  }

  .preview.stale {
    opacity: 0.45;
  }

  .empty {
    margin: 3rem 1.5rem;
    color: var(--ink-2);
    text-align: center;
  }

  .note {
    flex: none;
    margin: 0;
    padding: 0.6rem 1rem;
    border-top: 1px solid var(--line);
    background: var(--wash);
    color: var(--ink-3);
    font-size: 0.82rem;
    line-height: 1.45;
  }

  @media (max-width: 900px) {
    .repl-grid {
      grid-template-columns: minmax(0, 1fr);
    }

    .pane {
      height: auto;
      min-height: 0;
    }

    .editor-pane {
      position: static;
    }

    .editor {
      height: 46vh;
      min-height: 280px;
      flex: none;
    }

    /* The page scrolls here, so the sheet scrolls inside a box of its own
     * instead of growing to two hundred rows. */
    .preview {
      --sheet-max-height: 70vh;
    }
  }
</style>
