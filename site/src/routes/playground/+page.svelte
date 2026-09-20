<script lang="ts">
  import { base } from '$app/paths';
  import { onMount } from 'svelte';

  // @office-kit/xlsx is consumed via the source-tree path aliases set in
  // svelte.config.js (`@office-kit/xlsx/io` → `../src/io/index.ts`, etc). The
  // playground exercises the real surface the same way the rest of the
  // docs site does.
  import { isFormulaValue } from '@office-kit/xlsx/cell';
  import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
  import { getCellDisplayText, registerCellStyle } from '@office-kit/xlsx/styles';
  import { columnLetterFromIndex } from '@office-kit/xlsx/utils';
  import { addWorksheet, createWorkbook, iterWorksheets } from '@office-kit/xlsx/workbook';
  import { appendRow, appendRows, getCell, getCellExtent, getCellExtentRef } from '@office-kit/xlsx/worksheet';
  import type { Workbook } from '@office-kit/xlsx/workbook';
  import type { Worksheet } from '@office-kit/xlsx/worksheet';

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const SAMPLE_NAME = 'sample.xlsx';

  // A DOM table of a whole sheet would freeze the tab on a large workbook, so
  // the preview stops here and says so.
  const MAX_PREVIEW_ROWS = 200;
  const MAX_PREVIEW_COLS = 40;

  type GridCell = { text: string; numeric: boolean };
  type Preview = {
    columns: string[];
    rows: Array<{ number: number; cells: GridCell[] }>;
    usedRange: string;
    clipped: boolean;
  };

  let fileName = $state('');
  let status = $state('');
  let busy = $state(false);
  let dropping = $state(false);
  let workbook = $state.raw<Workbook | undefined>();
  let loadedBytes = $state(0);
  let activeSheet = $state(0);

  const worksheets = $derived(workbook ? [...iterWorksheets(workbook)] : []);
  const preview = $derived.by(() => {
    const ws = worksheets[activeSheet];
    return workbook && ws ? previewSheet(workbook, ws) : undefined;
  });

  function previewSheet(wb: Workbook, ws: Worksheet): Preview | undefined {
    const extent = getCellExtent(ws);
    if (!extent) return undefined;
    const lastRow = Math.min(extent.maxRow, MAX_PREVIEW_ROWS);
    const lastCol = Math.min(extent.maxCol, MAX_PREVIEW_COLS);
    const columnIndices = Array.from({ length: lastCol }, (_, i) => i + 1);
    return {
      columns: columnIndices.map(columnLetterFromIndex),
      rows: Array.from({ length: lastRow }, (_, i) => ({
        number: i + 1,
        cells: columnIndices.map((col): GridCell => {
          const cell = getCell(ws, i + 1, col);
          if (!cell) return { text: '', numeric: false };
          const value = isFormulaValue(cell.value) ? cell.value.cachedValue : cell.value;
          return { text: getCellDisplayText(wb, cell), numeric: typeof value === 'number' };
        }),
      })),
      usedRange: getCellExtentRef(ws) ?? '',
      clipped: extent.maxRow > lastRow || extent.maxCol > lastCol,
    };
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
    // Formatted, so the preview has a number format to put the values through.
    const usd = registerCellStyle(wb, { numberFormat: '"$"#,##0' });
    for (const [region, q1, q2, q3, q4] of data) {
      appendRow(sales, [region, q1, q2, q3, q4, q1 + q2 + q3 + q4], {
        styleIds: [undefined, usd, usd, usd, usd, usd],
      });
    }

    const inventory = addWorksheet(wb, 'Inventory');
    appendRow(inventory, ['SKU', 'Item', 'On hand', 'Reorder at']);
    const count = registerCellStyle(wb, { numberFormat: '#,##0' });
    appendRows(
      inventory,
      [
        ['A-0001', 'Spec sheet, A4 ruled', 480, 120],
        ['A-0002', 'Spec sheet, A4 grid', 215, 120],
        ['B-0010', 'Editor pencil, soft', 1320, 400],
        ['B-0011', 'Editor pencil, hard', 905, 400],
      ],
      { styleIds: [undefined, undefined, count, count] },
    );

    return wb;
  }

  const sheetCount = (wb: Workbook): string => {
    const n = [...iterWorksheets(wb)].length;
    return `${n} worksheet${n === 1 ? '' : 's'}`;
  };

  function show(wb: Workbook, name: string, byteLength: number): void {
    workbook = wb;
    fileName = name;
    loadedBytes = byteLength;
    activeSheet = 0;
  }

  async function loadSample(): Promise<void> {
    busy = true;
    status = 'Building the sample workbook…';
    try {
      // Saved and loaded again, so the sample takes the same path a dropped file does.
      const bytes = await workbookToBytes(buildSampleWorkbook());
      const wb = await loadWorkbook(fromArrayBuffer(bytes));
      show(wb, SAMPLE_NAME, bytes.byteLength);
      status = `Built and re-read the sample: ${sheetCount(wb)}.`;
    } catch (err) {
      console.error(err);
      status = `The sample could not be built: ${(err as Error).message}`;
    } finally {
      busy = false;
    }
  }

  async function loadFromFile(file: File): Promise<void> {
    busy = true;
    status = `Reading ${file.name}…`;
    try {
      const wb = await loadWorkbook(fromArrayBuffer(await file.arrayBuffer()));
      show(wb, file.name, file.size);
      status = `Loaded ${file.name}: ${sheetCount(wb)}.`;
    } catch (err) {
      // A bad file is the expected failure here, and the library's message
      // names what is wrong with it (encrypted, ISO strict, not a zip).
      status = `${file.name} could not be read: ${(err as Error).message}`;
    } finally {
      busy = false;
    }
  }

  function onDrop(event: DragEvent): void {
    event.preventDefault();
    dropping = false;
    const file = event.dataTransfer?.files[0];
    if (file) void loadFromFile(file);
  }

  async function downloadResaved(): Promise<void> {
    if (!workbook) return;
    busy = true;
    try {
      const bytes = await workbookToBytes(workbook);
      const url = URL.createObjectURL(new Blob([bytes.slice()], { type: XLSX_MIME }));
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName.replace(/(\.\w+)?$/, (ext) => `.resaved${ext || '.xlsx'}`);
      a.click();
      URL.revokeObjectURL(url);
      status = `Saved ${a.download}: ${bytes.byteLength.toLocaleString()} bytes written by the library.`;
    } catch (err) {
      console.error(err);
      status = `The workbook could not be saved: ${(err as Error).message}`;
    } finally {
      busy = false;
    }
  }

  onMount(() => {
    void loadSample();
  });
</script>

<svelte:head>
  <title>Playground · @office-kit/xlsx</title>
</svelte:head>

<section class="content">
  <h1>Open an .xlsx in your browser</h1>
  <p class="lede">
    Drop a workbook and this page parses it with the real <code>@office-kit/xlsx</code> source, then
    shows each cell as <code>getCellDisplayText</code> reads it: the value put through the cell’s
    number format. Nothing is uploaded: the whole pipeline runs in this tab.
  </p>

  <div
    class="drop"
    class:dropping
    role="group"
    aria-label="Choose an .xlsx file"
    ondragover={(e) => {
      e.preventDefault();
      dropping = true;
    }}
    ondragleave={() => (dropping = false)}
    ondrop={onDrop}
  >
    <p class="drop-text">{fileName || 'Drop an .xlsx file here'}</p>
    <div class="drop-actions">
      <label class="btn primary drop-pick">
        <input
          type="file"
          accept=".xlsx,.xlsm,{XLSX_MIME}"
          onchange={(e) => {
            const file = e.currentTarget.files?.[0];
            if (file) void loadFromFile(file);
          }}
        />
        Choose a file
      </label>
      <button type="button" class="btn" onclick={loadSample} disabled={busy}>
        Load the sample workbook
      </button>
      {#if workbook}
        <button type="button" class="btn" onclick={downloadResaved} disabled={busy}>
          Download the re-saved file
        </button>
      {/if}
    </div>
  </div>

  <p class="caveat">
    The preview is values only: no fills, fonts, merges, or charts, and formulas show the result
    Excel cached, because the library never calculates. The re-saved file is the loaded workbook
    written back out by the library, with everything it does not model carried through. For files
    too large to load whole, see <a href="{base}/docs/streaming">streaming</a>.
  </p>

  <p class="status" class:busy aria-live="polite">{status}</p>

  {#if workbook}
    <div class="meta-clip">
      <div class="meta">
        <div class="cell">
          <span class="label">File size</span>
          <span class="value">{loadedBytes.toLocaleString()} bytes</span>
        </div>
        <div class="cell">
          <span class="label">Worksheets</span>
          <span class="value">{worksheets.length}</span>
        </div>
        <div class="cell">
          <span class="label">Used range of this sheet</span>
          <span class="value">{preview?.usedRange || 'Empty'}</span>
        </div>
      </div>
    </div>

    <div class="tabs" role="tablist" aria-label="Worksheets">
      {#each worksheets as ws, index (index)}
        <button
          type="button"
          role="tab"
          aria-selected={index === activeSheet}
          onclick={() => (activeSheet = index)}
        >
          {ws.title}
        </button>
      {/each}
    </div>

    {#if preview}
      <!-- A scrollable region must be focusable, or keyboard users cannot scroll it. -->
      <!-- svelte-ignore a11y_no_noninteractive_tabindex -->
      <div class="grid-scroll" tabindex="0" role="region" aria-label="Cells of the selected worksheet">
        <table class="grid">
          <thead>
            <tr>
              <td></td>
              {#each preview.columns as letter (letter)}
                <th scope="col">{letter}</th>
              {/each}
            </tr>
          </thead>
          <tbody>
            {#each preview.rows as row (row.number)}
              <tr>
                <th scope="row">{row.number}</th>
                {#each row.cells as cell, col (col)}
                  <td class:numeric={cell.numeric}>{cell.text}</td>
                {/each}
              </tr>
            {/each}
          </tbody>
        </table>
      </div>
      {#if preview.clipped}
        <p class="caveat">
          Showing the first {MAX_PREVIEW_ROWS} rows and {MAX_PREVIEW_COLS} columns of {preview.usedRange}.
        </p>
      {/if}
    {:else}
      <p class="empty">This sheet has no cells.</p>
    {/if}
  {/if}
</section>

<style>
  .content {
    max-width: 1000px;
    margin: 0 auto;
    padding: 2.75rem var(--gutter) 5rem;
  }

  .lede {
    max-width: 66ch;
    color: var(--ink-2);
    font-size: 1.08rem;
  }

  .drop {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 1.1rem;
    margin: 2rem 0 0;
    padding: 2.25rem 1.25rem;
    border: 1.5px dashed var(--line-strong);
    border-radius: 12px;
    background: var(--wash);
    text-align: center;
    transition:
      border-color 120ms ease,
      background 120ms ease;
  }

  .drop.dropping {
    border-color: var(--accent);
    background: var(--accent-wash);
  }

  .drop-text {
    margin: 0;
    font-family: var(--display);
    font-size: 1.25rem;
    font-weight: 600;
    letter-spacing: -0.015em;
    overflow-wrap: anywhere;
  }

  .drop-actions {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 0.6rem;
  }

  .drop-pick input {
    position: absolute;
    width: 1px;
    height: 1px;
    opacity: 0;
  }

  .drop-pick:focus-within {
    outline: 2px solid var(--accent);
    outline-offset: 2px;
  }

  .btn:disabled {
    opacity: 0.6;
    cursor: progress;
  }

  .caveat {
    max-width: 72ch;
    margin: 1rem 0 0;
    color: var(--ink-3);
    font-size: 0.88rem;
  }

  .status {
    min-height: 1.6em;
    margin: 1.5rem 0 0;
    color: var(--ink-2);
    font-size: 0.95rem;
    overflow-wrap: anywhere;
  }

  .status.busy {
    color: var(--accent-ink);
  }

  /* Each cell draws its own right and bottom rule and the grid is pulled 1px
   * past the clipping box, so the outer edge never doubles up and a short last
   * row leaves plain paper rather than a filled gap. */
  .meta-clip {
    margin-top: 1rem;
    border: 1px solid var(--line);
    border-radius: var(--radius);
    overflow: hidden;
  }

  .meta {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    margin: 0 -1px -1px 0;
  }

  .cell {
    display: flex;
    flex-direction: column;
    gap: 0.15rem;
    padding: 0.8rem 1rem;
    border-right: 1px solid var(--line);
    border-bottom: 1px solid var(--line);
  }

  .label {
    color: var(--ink-3);
    font-size: 0.8rem;
  }

  .value {
    font-weight: 600;
    font-size: 0.97rem;
    overflow-wrap: anywhere;
  }

  .tabs {
    display: flex;
    gap: 0.25rem;
    margin-top: 2rem;
    overflow-x: auto;
    border-bottom: 1px solid var(--line-strong);
  }

  .tabs button {
    flex: none;
    min-height: 40px;
    margin-bottom: -1px;
    padding: 0 0.9rem;
    border: none;
    border-bottom: 2px solid transparent;
    background: none;
    color: var(--ink-2);
    font-family: var(--sans);
    font-size: 0.93rem;
    font-weight: 550;
    cursor: pointer;
  }

  .tabs button:hover {
    color: var(--ink);
  }

  .tabs button[aria-selected='true'] {
    border-bottom-color: var(--accent);
    color: var(--accent-ink);
  }

  .tabs button:focus-visible {
    outline-offset: -2px;
  }

  .grid-scroll {
    max-height: 70vh;
    overflow: auto;
    border: 1px solid var(--line-strong);
    border-top: none;
  }

  .grid {
    width: max-content;
    min-width: 100%;
    margin: 0;
    border-collapse: separate;
    border-spacing: 0;
    font-size: 0.86rem;
    font-variant-numeric: tabular-nums;
  }

  .grid th,
  .grid td {
    max-width: 22rem;
    padding: 0.35rem 0.6rem;
    border: none;
    border-right: 1px solid var(--line);
    border-bottom: 1px solid var(--line);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .grid th,
  .grid thead td {
    background: var(--wash);
    color: var(--ink-3);
    font-size: 0.78rem;
    font-weight: 500;
    text-align: center;
  }

  /* Headers stay in view while the cells scroll under them, as in Excel. */
  .grid thead th,
  .grid thead td {
    position: sticky;
    top: 0;
    z-index: 2;
    border-bottom-color: var(--line-strong);
  }

  .grid tbody th {
    position: sticky;
    left: 0;
    z-index: 1;
    border-right-color: var(--line-strong);
  }

  .grid thead td {
    left: 0;
    z-index: 3;
    border-right-color: var(--line-strong);
  }

  .grid td.numeric {
    text-align: right;
  }

  .empty {
    margin: 0;
    padding: 2rem 1rem;
    border: 1px solid var(--line-strong);
    border-top: none;
    color: var(--ink-3);
    text-align: center;
  }
</style>
