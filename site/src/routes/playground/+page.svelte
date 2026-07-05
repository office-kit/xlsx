<script lang="ts">
  import { base } from '$app/paths';
  import { onMount } from 'svelte';

  // @office-kit/xlsx is consumed via the source-tree path aliases set in
  // svelte.config.js (`@office-kit/xlsx/io` → `../src/io/index.ts`, etc). The
  // playground exercises the real surface the same way the rest of the
  // docs site does.
  import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
  import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
  import { appendRow, getMaxCol, getMaxRow, getCell } from '@office-kit/xlsx/worksheet';
  import type { CellValue } from '@office-kit/xlsx/cell';
  import type { Workbook } from '@office-kit/xlsx/workbook';
  import type { Worksheet } from '@office-kit/xlsx/worksheet';

  type GridCell = { value: string; kind: string };

  let fileName = $state<string>('built-in sample');
  let status = $state<string>('Ready.');
  let busy = $state<boolean>(false);
  let workbook = $state<Workbook | null>(null);
  let sheetTitles = $state<string[]>([]);
  let activeSheetIdx = $state<number>(0);
  let grid = $state<GridCell[][]>([]);
  let dropping = $state<boolean>(false);
  let lastBytes = $state<Uint8Array | null>(null);

  function formatValue(v: CellValue): { value: string; kind: string } {
    if (v == null) return { value: '', kind: 'empty' };
    if (typeof v === 'string') return { value: v, kind: 'string' };
    if (typeof v === 'number') return { value: String(v), kind: 'number' };
    if (typeof v === 'boolean') return { value: v ? 'TRUE' : 'FALSE', kind: 'boolean' };
    if (v instanceof Date) return { value: v.toISOString().slice(0, 10), kind: 'date' };
    // Discriminated complex values: rich text, formulas, errors, durations.
    if (typeof v === 'object' && 'kind' in v) {
      switch (v.kind) {
        case 'rich-text':
          return { value: v.runs.map((r) => r.text).join(''), kind: 'string' };
        case 'error':
          return { value: v.code, kind: 'error' };
        case 'duration':
          return { value: `${v.ms} ms`, kind: 'number' };
        default:
          // Formula values (with .formula / .result) and anything else: show
          // the cached result if available, else stringify.
          if ('result' in v && v.result != null) return formatValue(v.result as CellValue);
          return { value: JSON.stringify(v), kind: 'string' };
      }
    }
    return { value: String(v), kind: 'string' };
  }

  function snapshotSheet(ws: Worksheet): GridCell[][] {
    const maxRow = getMaxRow(ws);
    const maxCol = getMaxCol(ws);
    const out: GridCell[][] = [];
    for (let r = 1; r <= maxRow; r++) {
      const row: GridCell[] = [];
      for (let c = 1; c <= maxCol; c++) {
        const cell = getCell(ws, r, c);
        row.push(cell ? formatValue(cell.value) : { value: '', kind: 'empty' });
      }
      out.push(row);
    }
    return out;
  }

  function refresh(): void {
    if (!workbook) return;
    const sheet = workbook.sheets[activeSheetIdx];
    if (!sheet || sheet.kind !== 'worksheet') {
      grid = [];
      return;
    }
    grid = snapshotSheet(sheet.sheet);
  }

  function applyWorkbook(wb: Workbook): void {
    workbook = wb;
    sheetTitles = wb.sheets
      .filter((s) => s.kind === 'worksheet')
      .map((s) => s.sheet.title);
    activeSheetIdx = 0;
    refresh();
  }

  function buildSampleWorkbook(): Workbook {
    const wb = createWorkbook();

    const sales = addWorksheet(wb, 'Sales by quarter');
    appendRow(sales, ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total']);
    const data: Array<[string, number, number, number, number]> = [
      ['North America', 120000, 132000, 148000, 165000],
      ['Europe', 88000, 91000, 97000, 104000],
      ['Asia / Pacific', 64000, 72000, 81000, 95000],
      ['Latin America', 22000, 24000, 27000, 31000],
    ];
    for (const [region, q1, q2, q3, q4] of data) {
      appendRow(sales, [region, q1, q2, q3, q4, q1 + q2 + q3 + q4]);
    }
    appendRow(sales, ['']);
    appendRow(sales, ['Note', 'All figures in USD. Built in the browser by @office-kit/xlsx.']);

    const inventory = addWorksheet(wb, 'Inventory');
    appendRow(inventory, ['SKU', 'Item', 'On hand', 'Reorder at']);
    appendRow(inventory, ['A-0001', 'Spec sheet, A4 ruled', 480, 120]);
    appendRow(inventory, ['A-0002', 'Spec sheet, A4 grid', 215, 120]);
    appendRow(inventory, ['B-0010', 'Editor pencil, soft', 1320, 400]);
    appendRow(inventory, ['B-0011', 'Editor pencil, hard', 905, 400]);

    return wb;
  }

  async function loadSample(): Promise<void> {
    busy = true;
    try {
      fileName = 'built-in sample';
      status = 'Building sample workbook…';
      const wb = buildSampleWorkbook();
      applyWorkbook(wb);
      lastBytes = await workbookToBytes(wb);
      status = `Rendered ${sheetTitles.length} sheets · ${lastBytes.byteLength.toLocaleString()} bytes.`;
    } catch (err) {
      status = `Failed: ${(err as Error).message}`;
    } finally {
      busy = false;
    }
  }

  async function loadFromFile(file: File): Promise<void> {
    busy = true;
    try {
      fileName = file.name;
      status = `Reading ${file.name}…`;
      const buf = await file.arrayBuffer();
      const wb = await loadWorkbook(fromArrayBuffer(new Uint8Array(buf)));
      applyWorkbook(wb);
      lastBytes = new Uint8Array(buf);
      status = `Loaded ${sheetTitles.length} sheets · ${buf.byteLength.toLocaleString()} bytes.`;
    } catch (err) {
      status = `Failed: ${(err as Error).message}`;
    } finally {
      busy = false;
    }
  }

  function handleFileInput(event: Event): void {
    const input = event.currentTarget as HTMLInputElement;
    const file = input.files?.[0];
    if (file) void loadFromFile(file);
  }

  function handleDrop(event: DragEvent): void {
    event.preventDefault();
    dropping = false;
    const file = event.dataTransfer?.files[0];
    if (file) void loadFromFile(file);
  }

  function handleDragOver(event: DragEvent): void {
    event.preventDefault();
    dropping = true;
  }

  function handleDragLeave(): void {
    dropping = false;
  }

  function downloadCurrent(): void {
    if (!lastBytes) return;
    // Force a fresh ArrayBuffer-backed view so the Blob constructor's
    // BlobPart type is happy under strict TS lib settings.
    const ab = lastBytes.slice().buffer as ArrayBuffer;
    const blob = new Blob([ab], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function colLetter(idx: number): string {
    // 1-indexed column number → A..Z, AA..AZ, … (Excel-style).
    let n = idx;
    let s = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function selectSheet(idx: number): void {
    activeSheetIdx = idx;
    refresh();
  }

  onMount(() => {
    void loadSample();
  });
</script>

<svelte:head>
  <title>Playground · @office-kit/xlsx</title>
</svelte:head>

<section class="head">
  <div class="head-inner">
    <p class="eyebrow">Playground</p>
    <h1>Open an <em>.xlsx</em> in your browser.</h1>
    <p class="lede">
      Drop a workbook onto the canvas, or generate a built-in sample with
      <code>createWorkbook</code> + <code>addWorksheet</code> + <code>appendRow</code>.
      The cells are read back with <code>iterRows</code> / <code>getCell</code> and rendered
      live below. Nothing leaves the page — every step runs in your browser.
    </p>
  </div>
</section>

<section class="workbench">
  <aside class="controls" aria-label="Playground controls">
    <header class="controls-head">
      <span class="bracket">[</span>
      <span class="title">controls</span>
      <span class="bracket">]</span>
    </header>

    <button class="btn primary" onclick={loadSample} disabled={busy}>
      <span>Generate sample</span>
      <span class="arrow">↻</span>
    </button>

    <label class="file-btn">
      <input type="file" accept=".xlsx" onchange={handleFileInput} />
      <span>Open .xlsx…</span>
    </label>

    <button class="btn ghost" onclick={downloadCurrent} disabled={!lastBytes}>
      Download workbook bytes
    </button>

    <dl class="status">
      <div>
        <dt>file</dt>
        <dd>{fileName}</dd>
      </div>
      <div>
        <dt>size</dt>
        <dd>{lastBytes ? `${lastBytes.byteLength.toLocaleString()} B` : '—'}</dd>
      </div>
      <div>
        <dt>sheets</dt>
        <dd>{sheetTitles.length}</dd>
      </div>
      <div>
        <dt>state</dt>
        <dd class:busy>{status}</dd>
      </div>
    </dl>

    <p class="note">
      <strong>What is this?</strong>
      <br />
      A read-only tabular preview of the first <em>n</em> populated rows / columns. The
      workbook stays in memory — switch sheets in the tab strip below.
    </p>

    <p class="note muted">
      Streaming reader (<code>loadWorkbookStream</code>) and write-only writer
      (<code>createWriteOnlyWorkbook</code>) ship in the same package — see
      <a href="{base}/docs/streaming">Streaming</a>.
    </p>
  </aside>

  <div class="stage" class:dropping ondragover={handleDragOver} ondragleave={handleDragLeave} ondrop={handleDrop} role="region" aria-label="Workbook preview">
    <div class="stage-shell">
      <div class="stage-head">
        <span class="dot dot-a"></span>
        <span class="dot dot-b"></span>
        <span class="dot dot-c"></span>
        <span class="stage-name">{fileName}</span>
      </div>

      {#if sheetTitles.length > 0}
        <div class="tabs" role="tablist">
          {#each sheetTitles as title, idx (title)}
            <button
              role="tab"
              class="tab"
              class:active={idx === activeSheetIdx}
              onclick={() => selectSheet(idx)}
              aria-selected={idx === activeSheetIdx}
            >
              <span class="tab-idx">{(idx + 1).toString().padStart(2, '0')}</span>
              <span class="tab-label">{title}</span>
            </button>
          {/each}
        </div>
      {/if}

      <div class="stage-scroll">
        {#if grid.length === 0}
          <p class="stage-placeholder">No populated cells in this sheet.</p>
        {:else}
          <table class="grid">
            <thead>
              <tr>
                <th class="corner" aria-hidden="true"></th>
                {#each Array(grid[0]?.length ?? 0) as _, c (c)}
                  <th class="col-h">{colLetter(c + 1)}</th>
                {/each}
              </tr>
            </thead>
            <tbody>
              {#each grid as row, r (r)}
                <tr>
                  <th class="row-h">{r + 1}</th>
                  {#each row as cell, c (c)}
                    <td class="cell kind-{cell.kind}">{cell.value}</td>
                  {/each}
                </tr>
              {/each}
            </tbody>
          </table>
        {/if}
      </div>

      {#if dropping}
        <div class="drop-overlay">
          <span>Drop the .xlsx to render</span>
        </div>
      {/if}
    </div>
  </div>
</section>

<style>
  .head {
    padding: 2.6rem 1.5rem 1.5rem;
    border-bottom: 1px solid var(--border);
  }

  .head-inner {
    max-width: var(--max-wide);
    margin: 0 auto;
  }

  h1 {
    font-family: var(--display);
    font-weight: 460;
    font-size: clamp(1.9rem, 4.5vw, 2.85rem);
    line-height: 1.05;
    margin: 0 0 0.85rem;
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
    letter-spacing: -0.025em;
    max-width: 22ch;
  }

  h1 em {
    color: var(--accent);
    font-style: italic;
    font-variation-settings: 'opsz' 144, 'SOFT' 70;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.04rem;
    line-height: 1.55;
    max-width: 66ch;
    margin: 0;
  }

  .workbench {
    display: grid;
    grid-template-columns: 280px 1fr;
    gap: 0;
    max-width: var(--max-wide);
    margin: 0 auto;
    padding: 1.75rem 1.5rem 4rem;
  }

  @media (max-width: 880px) {
    .workbench {
      grid-template-columns: 1fr;
      gap: 1.25rem;
    }
  }

  .controls {
    border: 1px solid var(--border);
    border-right: none;
    border-radius: var(--radius) 0 0 var(--radius);
    background: var(--bg-elev);
    padding: 1.1rem 1.1rem 1.25rem;
    display: flex;
    flex-direction: column;
    gap: 0.75rem;
  }

  @media (max-width: 880px) {
    .controls {
      border-right: 1px solid var(--border);
      border-radius: var(--radius);
    }
  }

  .controls-head {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--fg-muted);
    letter-spacing: 0.18em;
    text-transform: uppercase;
    margin-bottom: 0.25rem;
    display: flex;
    gap: 0.4rem;
    align-items: center;
  }

  .controls-head .bracket {
    color: var(--accent);
  }

  .btn {
    display: inline-flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.5rem;
    padding: 0.65rem 0.85rem;
    border-radius: var(--radius-sm);
    border: 1px solid var(--border);
    background: var(--bg-soft);
    color: var(--fg);
    font-family: var(--sans);
    font-size: 0.9rem;
    font-weight: 540;
    cursor: pointer;
    transition:
      background 120ms ease,
      border-color 120ms ease,
      color 120ms ease,
      transform 120ms ease;
  }

  .btn:hover:not([disabled]) {
    background: var(--bg-paper);
    border-color: var(--border-strong);
    transform: translateY(-1px);
  }

  .btn[disabled] {
    opacity: 0.55;
    cursor: not-allowed;
  }

  .btn.primary {
    background: var(--accent);
    color: var(--bg);
    border-color: var(--accent);
    box-shadow: 0 8px 24px -14px var(--accent-glow);
  }

  .btn.primary:hover:not([disabled]) {
    background: var(--accent-hot);
    border-color: var(--accent-hot);
  }

  .btn.ghost {
    background: transparent;
  }

  .btn .arrow {
    color: var(--bg);
    opacity: 0.7;
  }

  .file-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    padding: 0.65rem 0.85rem;
    border-radius: var(--radius-sm);
    border: 1px dashed var(--border-strong);
    background: var(--bg-soft);
    color: var(--fg);
    font-family: var(--sans);
    font-size: 0.9rem;
    font-weight: 540;
    cursor: pointer;
    text-align: center;
    transition: background 120ms ease, border-color 120ms ease;
  }

  .file-btn:hover {
    background: var(--bg-paper);
    border-color: var(--accent);
    color: var(--fg);
  }

  .file-btn input {
    display: none;
  }

  .status {
    margin: 0.5rem 0 0;
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    background: var(--code-bg);
    padding: 0.7rem 0.75rem;
    display: flex;
    flex-direction: column;
    gap: 0.45rem;
    font-family: var(--mono);
    font-size: 11.5px;
  }

  .status > div {
    display: grid;
    grid-template-columns: 5ch 1fr;
    align-items: baseline;
    gap: 0.6rem;
  }

  .status dt {
    color: var(--fg-muted);
    text-transform: uppercase;
    letter-spacing: 0.1em;
    font-size: 10px;
  }

  .status dd {
    color: var(--fg);
    margin: 0;
    word-break: break-word;
  }

  .status dd.busy {
    color: var(--accent);
  }

  .note {
    color: var(--fg-soft);
    font-size: 0.86rem;
    line-height: 1.5;
    margin: 0;
  }

  .note.muted {
    color: var(--fg-muted);
    font-size: 0.78rem;
  }

  .note strong {
    color: var(--fg);
  }

  .stage {
    border: 1px solid var(--border);
    border-radius: 0 var(--radius) var(--radius) 0;
    background: var(--bg-paper);
    overflow: hidden;
    position: relative;
    min-height: 70vh;
  }

  @media (max-width: 880px) {
    .stage {
      border-radius: var(--radius);
    }
  }

  .stage-shell {
    display: flex;
    flex-direction: column;
    height: 100%;
  }

  .stage-head {
    display: flex;
    align-items: center;
    gap: 0.45rem;
    padding: 0.6rem 0.85rem;
    background: var(--bg-elev);
    border-bottom: 1px solid var(--border);
    font-family: var(--mono);
    font-size: 12px;
    color: var(--fg-muted);
  }

  .dot {
    width: 9px;
    height: 9px;
    border-radius: 50%;
    background: var(--bg-soft);
    border: 1px solid var(--border-strong);
  }

  .dot-a {
    background: color-mix(in oklab, var(--accent) 60%, var(--bg-soft));
    border-color: var(--accent-soft);
  }

  .dot-b {
    background: color-mix(in oklab, var(--brass) 50%, var(--bg-soft));
    border-color: var(--brass-soft);
  }

  .stage-name {
    margin-left: 0.4rem;
    color: var(--fg-soft);
  }

  /* Sheet tab strip — visual nod to Excel's sheet selector. */
  .tabs {
    display: flex;
    gap: 0.25rem;
    padding: 0.45rem 0.6rem;
    background: var(--bg-elev);
    border-bottom: 1px solid var(--border);
    overflow-x: auto;
  }

  .tab {
    display: inline-flex;
    align-items: baseline;
    gap: 0.45rem;
    padding: 0.35rem 0.7rem;
    background: transparent;
    border: 1px solid transparent;
    color: var(--fg-soft);
    font-family: var(--mono);
    font-size: 11.5px;
    cursor: pointer;
    border-radius: var(--radius-sm);
    white-space: nowrap;
    transition:
      background 120ms ease,
      color 120ms ease,
      border-color 120ms ease;
  }

  .tab:hover {
    color: var(--fg);
    background: var(--bg-soft);
  }

  .tab.active {
    color: var(--fg);
    background: var(--bg-paper);
    border-color: var(--border);
  }

  .tab.active .tab-idx {
    color: var(--accent);
  }

  .tab-idx {
    font-size: 10px;
    color: var(--fg-muted);
    letter-spacing: 0.06em;
  }

  .tab-label {
    font-family: var(--sans);
    font-size: 12.5px;
    font-weight: 540;
    letter-spacing: -0.005em;
  }

  /* Scroll slot — zero edge padding so the spreadsheet table aligns with
   * the row/column header gutters at the slot edge. */
  .stage-scroll {
    flex: 1;
    overflow: auto;
    padding: 0;
    background: var(--bg);
  }

  .stage-placeholder {
    color: var(--fg-muted);
    font-family: var(--mono);
    font-size: 12px;
    padding: 1.2rem 1.4rem;
    margin: 0;
  }

  /* The grid is the rendered worksheet. Borders are 1px hairlines that
   * collapse so column / row dividers stay 1 device pixel even when the
   * table scrolls. */
  .grid {
    border-collapse: collapse;
    width: max-content;
    min-width: 100%;
    margin: 0;
    font-family: var(--mono);
    font-size: 12.5px;
    background: var(--bg);
  }

  .grid th,
  .grid td {
    border-right: 1px solid var(--border);
    border-bottom: 1px solid var(--border);
    padding: 0.32rem 0.55rem;
    white-space: nowrap;
    vertical-align: top;
  }

  .grid thead th {
    position: sticky;
    top: 0;
    z-index: 2;
    background: var(--bg-paper);
    color: var(--fg-soft);
    font-weight: 500;
    text-align: center;
    font-size: 11px;
    letter-spacing: 0.06em;
  }

  .grid thead th.corner {
    left: 0;
    z-index: 3;
    border-right: 1px solid var(--border-strong);
  }

  .grid tbody th.row-h {
    position: sticky;
    left: 0;
    z-index: 1;
    background: var(--bg-paper);
    color: var(--fg-muted);
    font-weight: 500;
    text-align: right;
    font-size: 11px;
    min-width: 4ch;
    border-right: 1px solid var(--border-strong);
  }

  .grid td.cell {
    color: var(--fg);
    min-width: 7ch;
  }

  /* Right-align numeric and date cells, like a real spreadsheet. */
  .grid td.kind-number,
  .grid td.kind-date {
    text-align: right;
    color: var(--fg);
    font-variant-numeric: tabular-nums;
  }

  .grid td.kind-boolean {
    color: var(--brass);
    text-align: center;
    letter-spacing: 0.04em;
  }

  .grid td.kind-empty {
    background: color-mix(in oklab, var(--bg-paper) 60%, transparent);
  }

  /* Stripe alternating rows for readability. */
  .grid tbody tr:nth-child(even) td.cell {
    background: color-mix(in oklab, var(--bg-elev) 80%, transparent);
  }

  .drop-overlay {
    position: absolute;
    inset: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    background: color-mix(in oklab, var(--accent) 18%, transparent);
    border: 2px dashed var(--accent);
    color: var(--fg);
    font-family: var(--mono);
    font-size: 14px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    backdrop-filter: blur(2px);
    pointer-events: none;
  }

  .stage.dropping {
    border-color: var(--accent);
  }
</style>
