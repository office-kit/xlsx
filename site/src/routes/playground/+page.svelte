<script lang="ts">
  import { base } from '$app/paths';
  import { onMount } from 'svelte';

  // @office-kit/xlsx is consumed via the source-tree path aliases set in
  // svelte.config.js (`@office-kit/xlsx/io` → `../src/io/index.ts`, etc). The
  // playground exercises the real surface the same way the rest of the
  // docs site does.
  import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
  import { registerCellStyle } from '@office-kit/xlsx/styles';
  import { addWorksheet, createWorkbook, iterWorksheets } from '@office-kit/xlsx/workbook';
  import { appendRow, appendRows } from '@office-kit/xlsx/worksheet';
  import type { Workbook } from '@office-kit/xlsx/workbook';
  import SheetGrid from '$lib/components/SheetGrid.svelte';
  import { readWorkbookGrids } from '$lib/sheet-grid';

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const SAMPLE_NAME = 'sample.xlsx';

  // A DOM table of a whole sheet would freeze the tab on a large workbook, so
  // the preview stops here and says so.
  const PREVIEW_LIMIT = { maxRows: 200, maxCols: 40 };

  let fileName = $state('');
  let status = $state('');
  let busy = $state(false);
  let dropping = $state(false);
  let workbook = $state.raw<Workbook | undefined>();
  let loadedBytes = $state(0);

  const sheets = $derived(workbook ? readWorkbookGrids(workbook, PREVIEW_LIMIT) : []);
  const clipped = $derived(sheets.filter((sheet) => sheet.clipped));

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
    The preview draws values, number formats, fonts, fills, borders, merges, and column widths. It
    does not draw charts, images, or conditional formats, and formulas show the result Excel cached,
    because the library never calculates. The re-saved file is the loaded workbook
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
          <span class="value">{sheets.length}</span>
        </div>
        <div class="cell">
          <span class="label">Used ranges</span>
          <span class="value">
            {sheets.map((sheet) => `${sheet.name} ${sheet.usedRange || 'empty'}`).join(', ')}
          </span>
        </div>
      </div>
    </div>

    <div class="preview">
      <SheetGrid {sheets} />
    </div>
    {#if clipped.length > 0}
      <p class="caveat">
        Showing the first {PREVIEW_LIMIT.maxRows} rows and {PREVIEW_LIMIT.maxCols} columns of
        {clipped.map((sheet) => sheet.name).join(', ')}.
      </p>
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

  .preview {
    --sheet-max-height: 70vh;
    margin-top: 2rem;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius);
    overflow: hidden;
  }
</style>
