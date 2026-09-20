<script lang="ts">
  import { base } from '$app/paths';
  import { FamilyGrid, InstallCommand, getProduct } from '@office-kit/site-kit';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  const currentProduct = getProduct('xlsx');

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const DOWNLOAD_NAME = 'office-kit-demo.xlsx';
  const BYTES_PER_KB = 1024;

  // The formula bar describes the first formula cell, the way Excel's does for
  // the selected cell.
  const selected = $derived.by(() => {
    for (const row of data.grid.rows) {
      const index = row.cells.findIndex((cell) => cell.formula !== undefined);
      if (index >= 0) {
        return {
          row: row.number,
          col: index,
          ref: `${data.grid.columns[index]?.letter}${row.number}`,
          formula: row.cells[index]?.formula,
        };
      }
    }
    return undefined;
  });

  let download = $state<{ state: 'idle' | 'working' | 'done' | 'failed'; note: string }>({
    state: 'idle',
    note: '',
  });

  // Runs the exact function shown in the code panel, in the visitor's browser.
  // The library loads on click so the landing page itself ships none of it.
  async function downloadWorkbook(): Promise<void> {
    download = { state: 'working', note: '' };
    try {
      const [{ workbookToBytes }, { buildHeroWorkbook }] = await Promise.all([
        import('@office-kit/xlsx/io'),
        import('$lib/examples/hero-workbook'),
      ]);
      const bytes = await workbookToBytes(buildHeroWorkbook());
      const url = URL.createObjectURL(new Blob([bytes.slice()], { type: XLSX_MIME }));
      const a = document.createElement('a');
      a.href = url;
      a.download = DOWNLOAD_NAME;
      a.click();
      URL.revokeObjectURL(url);
      download = {
        state: 'done',
        note: `Saved ${DOWNLOAD_NAME} (${(bytes.byteLength / BYTES_PER_KB).toFixed(1)} KB), built in this tab.`,
      };
    } catch (err) {
      console.error(err);
      download = {
        state: 'failed',
        note: 'The workbook could not be built in this browser. The console has the error.',
      };
    }
  }

  const paths = [
    {
      title: 'Edit a workbook you already have',
      body: 'Load an .xlsx or .xlsm from bytes, a file, a Blob, or a fetch response. Change the cells that change and save. Pivot tables, VBA projects, threaded comments, external links, and custom XML are carried through byte for byte.',
      code: 'loadWorkbook(source)',
      href: '/docs/getting-started',
      link: 'Read the getting started guide',
    },
    {
      title: 'Build one from nothing',
      body: 'createWorkbook() returns an empty workbook with no template file to ship. Add sheets and rows, register a style once and reuse it by id, then add number formats, formulas, tables, data validation, conditional formatting, images, and charts.',
      code: 'createWorkbook()',
      href: '/docs/recipes',
      link: 'Browse the recipes',
    },
    {
      title: 'Stream the big ones',
      body: 'The write-only workbook deflates each row as you append it and pushes the bytes straight to the sink. The streaming reader walks a sheet with a SAX parser and yields one row at a time, so neither side holds the sheet in memory.',
      code: 'createWriteOnlyWorkbook(sink)',
      href: '/docs/streaming',
      link: 'Read the streaming guide',
    },
  ];

  const entries = [
    { path: 'io', what: 'Load and save a whole workbook', size: '107 KB', cap: '120 KB' },
    { path: 'streaming', what: 'Row-by-row reader and append-only writer', size: '61 KB', cap: '80 KB' },
    { path: 'styles', what: 'Fonts, fills, borders, number formats', size: '14 KB', cap: '60 KB' },
    { path: 'worksheet', what: 'Cells, ranges, merges, tables, validation', size: '11 KB', cap: '100 KB' },
    { path: 'workbook', what: 'Sheets, defined names, properties', size: '9 KB', cap: '100 KB' },
  ];

  const proof = [
    {
      claim: 'Every kind of write is checked against the ECMA-376 schemas.',
      how: 'A three-tier validator runs in CI: the OPC package structure, the ECMA-376 Transitional XSDs through xmllint, and the rules a schema cannot express, such as overlapping merges or a style index past the end of the table. A fast-check property test feeds it generated workbooks.',
    },
    {
      claim: 'What the library does not model, it does not touch.',
      how: 'Pivot tables, VBA, OLE objects, external links, and custom XML pass through load and save byte for byte. Round-trip tests pin that against real workbooks, including the openpyxl fixture corpus.',
    },
    {
      claim: 'A bundle that grows too much fails the build.',
      how: 'size-limit runs in CI on every push and pull request. The full load and save path has to stay under 120 KB minified and brotli-compressed with its dependencies, and the streaming entry under 80 KB.',
    },
    {
      claim: 'Three runtime dependencies.',
      how: 'fflate for ZIP, saxes for streaming XML, and fast-xml-parser. Only the @office-kit/xlsx/node subpath touches fs; the rest is Web Streams, Blob, and Uint8Array.',
    },
    {
      claim: 'Tested where you run it, and as you install it.',
      how: 'The suite runs on Linux, macOS, and Windows against Node 22, 24, and 26. A packaging job installs the packed tarball and compiles a consumer under node16, nodenext, and bundler resolution.',
    },
  ];

  const capabilities = [
    { area: 'Cells', items: 'Numbers, strings, booleans, errors, dates, and inline rich text. Normal, array, shared, and data-table formulas with cached values.' },
    { area: 'Styles', items: 'Fonts, fills, borders, alignment, protection, number formats, named styles, and differential styles, deduplicated into one pool.' },
    { area: 'Sheets', items: 'Merged cells, freeze panes, column widths and row heights, grouping, hidden rows and columns, page setup, sheet protection.' },
    { area: 'Data tools', items: 'Excel tables, autofilter, data validation, conditional formatting, and workbook- or sheet-scoped defined names.' },
    { area: 'Reading', items: 'A typed cell value union, the text Excel shows under a number format, dates read through the workbook epoch, and value extents.' },
    { area: 'Charts', items: '16 classic chart kinds and 8 modern ones such as sunburst, treemap, waterfall, and funnel, with trendlines, error bars, and chartsheets.' },
    { area: 'Drawings', items: 'PNG, JPEG, GIF, BMP, WebP, TIFF, SVG, EMF, and WMF images, with format and size detected from the bytes.' },
    { area: 'Package', items: 'Hyperlinks and comments in bulk, ZIP64 past 65,535 entries, deterministic output bytes, and macro-enabled .xlsm.' },
  ];

  const notYet = [
    'Formula evaluation. Cache the values you know, or ask Excel to recalculate on open',
    'Other formats: .xls, .xlsb, .ods, and .csv are out of scope',
    'ISO 29500 Strict workbooks are detected and named, not read',
    'Encrypted workbooks are detected, not decrypted',
    'Pivot table authoring (existing pivots are preserved)',
    'Charts, images, and tables in the streaming writer',
  ];
</script>

<svelte:head>
  <title>@office-kit/xlsx: read, edit, and write Excel files in TypeScript</title>
</svelte:head>

<section class="band hero">
  <div class="frame hero-inner">
    <h1>Read, edit, and write Excel files in TypeScript</h1>
    <p class="lede">
      Open any .xlsx or start from an empty workbook. Change cells, styles, formulas, and charts
      through typed functions, then save a file that validates against the ECMA-376 schemas. It runs
      in Node 22 and later, and in the browser.
    </p>
    <div class="cta">
      <a href="{base}/docs/getting-started" class="btn primary">Get started</a>
      <a href="{base}/playground" class="btn">Open the playground</a>
      <InstallCommand pkg={currentProduct.pkg} />
    </div>
  </div>
</section>

<section class="band stage" aria-labelledby="stage-title">
  <div class="frame stage-inner">
    <h2 id="stage-title" class="visually-hidden">A worksheet and the code that built it</h2>
    <div class="stage-grid">
      <figure class="code-pane">
        <figcaption>hero-workbook.ts</figcaption>
        <!-- A scrollable region must be focusable, or keyboard users cannot scroll it. -->
        <!-- svelte-ignore a11y_no_noninteractive_tabindex -->
        <div class="code-scroll" tabindex="0" role="region" aria-label="Source of hero-workbook.ts">
          {@html data.heroCode}
        </div>
      </figure>
      <figure class="sheet-pane">
        <div class="sheet">
          {#if selected}
            <div class="formula-bar">
              <span class="name-box">{selected.ref}</span>
              <span class="fx" aria-hidden="true">fx</span>
              <span class="formula">{selected.formula}</span>
            </div>
          {/if}
          <!-- svelte-ignore a11y_no_noninteractive_tabindex -->
          <div class="sheet-scroll" tabindex="0" role="region" aria-label="The worksheet the code produced">
            <table>
              <colgroup>
                <col class="gutter-col" />
                {#each data.grid.columns as column (column.letter)}
                  <col class="data-col" style:--col-w="{column.width}px" />
                {/each}
              </colgroup>
              <thead>
                <tr>
                  <td></td>
                  {#each data.grid.columns as column (column.letter)}
                    <th scope="col">{column.letter}</th>
                  {/each}
                </tr>
              </thead>
              <tbody>
                {#each data.grid.rows as row (row.number)}
                  <tr>
                    <th scope="row">{row.number}</th>
                    {#each row.cells as cell, col (col)}
                      <td
                        class:numeric={cell.numeric}
                        class:bold={cell.bold}
                        class:rule-above={cell.ruleAbove}
                        class:selected={selected?.row === row.number && selected.col === col}
                        style:color={cell.color}
                        style:background-color={cell.fill}
                      >
                        {cell.text}
                      </td>
                    {/each}
                  </tr>
                {/each}
              </tbody>
            </table>
          </div>
          <div class="sheet-tabs"><span>{data.grid.sheetName}</span></div>
        </div>
      </figure>
    </div>
    <div class="stage-foot">
      <p>
        That sheet is real output. The code shown with it built the workbook, the library saved it and
        read the bytes back, and <code>getCellDisplayText</code> put each value through its number format.
        Download it and open it in Excel: the totals are live formulas.
      </p>
      <div class="stage-action">
        <button
          type="button"
          class="btn stage-btn"
          onclick={downloadWorkbook}
          disabled={download.state === 'working'}
        >
          {download.state === 'working' ? 'Building the workbook…' : 'Download the .xlsx'}
        </button>
        <p class="stage-note" aria-live="polite">{download.note}</p>
      </div>
    </div>
  </div>
</section>

<section class="band">
  <div class="frame">
    <ul class="paths">
      {#each paths as p (p.title)}
        <li>
          <code class="path-code">{p.code}</code>
          <h2>{p.title}</h2>
          <p>{p.body}</p>
          <a href="{base}{p.href}">{p.link}</a>
        </li>
      {/each}
    </ul>
  </div>
</section>

<section class="band">
  <div class="frame split">
    <div class="split-text">
      <h2>Most spreadsheets that matter already exist</h2>
      <p>
        The report finance sends every month, the template with the pivot table, the .xlsm someone’s
        macros depend on. Load it, write the cells that change, and save. The parts the library does
        not model go back into the file exactly as they came out.
      </p>
      <p>
        Reading is as typed as writing. A cell value is a discriminated union, not
        <code>any</code>; <code>getCellDisplayText</code> gives the text Excel shows under the cell’s
        number format, and <code>getCellDate</code> reads a date serial through the workbook’s epoch.
      </p>
    </div>
    <figure class="split-code">
      <figcaption>basic-read-write.ts</figcaption>
      {@html data.roundTripCode}
    </figure>
  </div>
</section>

<section class="band">
  <div class="frame split reverse">
    <div class="split-text">
      <h2>A million rows without a million rows in memory</h2>
      <p>
        The streaming writer pushes each row through deflate as it arrives and forwards the chunk to
        the sink, so row buffering stays near 64 KiB however long the sheet gets. Shared strings are
        capped at 100,000 entries and 8 MiB; past that, new strings are written inline.
      </p>
      <p>
        The reader inflates a worksheet chunk by chunk into a SAX parser, so the inflated sheet is
        never fully resident. One honest limit: ZIP needs its central directory, so the compressed
        archive itself is loaded up front.
      </p>
    </div>
    <figure class="split-code">
      <figcaption>streaming-write.ts</figcaption>
      {@html data.streamingCode}
    </figure>
  </div>
</section>

<section class="band" aria-labelledby="entries-title">
  <div class="frame entries">
    <div class="section-head">
      <h2 id="entries-title">Import the part you use</h2>
      <p>
        There is no root import. The package has 15 subpaths, every export has exactly one home, and
        nothing has side effects, so a bundler keeps only what you call. Sizes are minified and
        brotli-compressed with dependencies included, measured at 0.18.0; the cap is what CI enforces.
      </p>
    </div>
    <div class="table-scroll">
      <table>
        <thead>
          <tr>
            <th scope="col">Subpath</th>
            <th scope="col">What it holds</th>
            <th scope="col" class="num">Size</th>
            <th scope="col" class="num">CI cap</th>
          </tr>
        </thead>
        <tbody>
          {#each entries as entry (entry.path)}
            <tr>
              <th scope="row"><code>{currentProduct.pkg}/{entry.path}</code></th>
              <td class="what">{entry.what}</td>
              <td class="num" data-label="Size">{entry.size}</td>
              <td class="num" data-label="CI cap">{entry.cap}</td>
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
    <p class="entries-more">
      The other ten cover cells, charts, chartsheets, drawings, Node file helpers, and the packaging,
      schema, XML, ZIP, and utility layers. <a href="{base}/api">The API reference</a> lists them all.
    </p>
  </div>
</section>

<section class="band proof" aria-labelledby="proof-title">
  <div class="frame proof-inner">
    <h2 id="proof-title">Valid files, checked by machines</h2>
    <p class="proof-lede">
      “Excel happens to open it today” is not the bar. The bytes have to be valid OOXML, and a
      machine has to say so on every commit.
    </p>
    <dl>
      {#each proof as item (item.claim)}
        <div class="proof-row">
          <dt>{item.claim}</dt>
          <dd>{item.how}</dd>
        </div>
      {/each}
    </dl>
  </div>
</section>

<section class="band" aria-labelledby="caps-title">
  <div class="frame caps">
    <div class="section-head">
      <h2 id="caps-title">What you can build today</h2>
      <p>
        The library is pre-1.0 and says so. This is what works now, and what does not.
        <a href="{base}/api">The API reference</a> lists every function.
      </p>
    </div>
    <dl class="caps-grid">
      {#each capabilities as c (c.area)}
        <div>
          <dt>{c.area}</dt>
          <dd>{c.items}</dd>
        </div>
      {/each}
    </dl>
    <div class="not-yet">
      <h3>Not supported</h3>
      <ul>
        {#each notYet as item (item)}
          <li>{item}</li>
        {/each}
      </ul>
    </div>
  </div>
</section>

<section class="band agents">
  <div class="frame agents-inner">
    <div>
      <h2>Written to be driven by AI agents too</h2>
      <p>
        Spreadsheets are increasingly written by agents, so the docs are built for them as well as
        for you. Because every function has one import path, a model that knows the name knows the
        import.
      </p>
    </div>
    <ul>
      <li>
        <a href="{base}/llms.txt"><code>/llms.txt</code></a>
        <span>A self-contained guide to the API, with an index of every docs page.</span>
      </li>
      <li>
        <a href="{base}/llms-full.txt"><code>/llms-full.txt</code></a>
        <span>The whole documentation in one file.</span>
      </li>
      <li>
        <a href="{base}/docs/getting-started.md"><code>any-docs-page.md</code></a>
        <span>Add .md to a docs URL to get the raw Markdown.</span>
      </li>
    </ul>
  </div>
</section>

<section class="band" aria-labelledby="family-title">
  <div class="frame">
    <div class="section-head family-head">
      <h2 id="family-title">One kit, three file formats</h2>
      <p>
        Office Kit is a family of libraries built on the same rules: the ECMA-376 spec is the source
        of truth, output has to validate, and one ESM build has to run everywhere.
      </p>
    </div>
    <FamilyGrid product="xlsx" />
  </div>
</section>

<style>
  /* Hero. */
  .hero-inner {
    padding: clamp(3rem, 6vw, 4.75rem) var(--gutter) clamp(2.75rem, 5vw, 3.75rem);
    text-align: center;
  }

  h1 {
    max-width: 19ch;
    margin: 0 auto;
    font-size: clamp(2.35rem, 6vw, 4.25rem);
    font-weight: 650;
    line-height: 1.02;
    letter-spacing: -0.04em;
  }

  .lede {
    max-width: 62ch;
    margin: 1.75rem auto 0;
    color: var(--ink-2);
    font-size: clamp(1.05rem, 1.6vw, 1.2rem);
    line-height: 1.55;
    text-wrap: balance;
  }

  .cta {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 0.6rem;
    margin-top: 2.25rem;
  }

  /* Stage: the product accent as a field, ruled like a worksheet's cell grid. */
  .stage {
    --stage-rule: rgb(255 255 255 / 0.14);
    background-color: var(--accent);
    background-image:
      linear-gradient(var(--stage-rule) 1px, transparent 1px),
      linear-gradient(90deg, var(--stage-rule) 1px, transparent 1px);
    background-size: 72px 24px;
  }

  .stage-inner {
    border-inline-color: rgb(255 255 255 / 0.22);
    padding: clamp(1.5rem, 4vw, 3.5rem) var(--gutter) 0;
  }

  .stage-grid {
    display: grid;
    grid-template-columns: minmax(0, 6fr) minmax(0, 6fr);
    gap: clamp(1rem, 2vw, 1.75rem);
    align-items: stretch;
  }

  figure {
    margin: 0;
  }

  .code-pane {
    display: flex;
    flex-direction: column;
    min-height: 0;
    border-radius: var(--radius);
    background: var(--night);
    box-shadow: var(--shadow-pop);
    overflow: hidden;
  }

  .code-pane figcaption,
  .split-code figcaption {
    flex: none;
    padding: 0.6rem 1rem;
    border-bottom: 1px solid var(--night-line);
    color: var(--night-ink-2);
    font-family: var(--mono);
    font-size: 0.8rem;
  }

  /* The pane's height comes from the sheet beside it; the code scrolls inside. */
  .code-scroll {
    position: relative;
    flex: 1;
    min-height: 0;
    overflow: auto;
  }

  .code-scroll :global(pre) {
    position: absolute;
    inset: 0 auto auto 0;
    min-width: 100%;
    margin: 0;
    border: none;
    border-radius: 0;
    overflow: visible;
    font-size: 0.8rem;
  }

  /* The sheet is a document, so like paper it stays light in both themes. */
  .sheet {
    --sheet-paper: #ffffff;
    --sheet-chrome: #f3f4f6;
    --sheet-rule: #e2e4e9;
    --sheet-rule-strong: #c9cdd6;
    --sheet-ink: #15171c;
    --sheet-ink-2: #5b616e;
    --sheet-select: #168a4f;
    --sheet-row-h: 30px;
    --sheet-gutter-w: 38px;
    --sheet-scale: 1;

    display: flex;
    flex-direction: column;
    height: 100%;
    border-radius: 4px;
    background: var(--sheet-paper);
    color: var(--sheet-ink);
    box-shadow: var(--shadow-pop);
    overflow: hidden;
    font-size: 0.84rem;
    line-height: 1.2;
  }

  .formula-bar {
    flex: none;
    display: flex;
    align-items: center;
    gap: 0.6rem;
    padding: 0.45rem 0.6rem;
    border-bottom: 1px solid var(--sheet-rule-strong);
    background: var(--sheet-chrome);
    font-family: var(--mono);
    font-size: 0.78rem;
  }

  .name-box,
  .formula {
    padding: 0.3rem 0.55rem;
    border: 1px solid var(--sheet-rule-strong);
    border-radius: 3px;
    background: var(--sheet-paper);
  }

  .name-box {
    min-width: 3.25rem;
  }

  .formula {
    flex: 1;
    min-width: 0;
  }

  .fx {
    color: var(--sheet-ink-2);
    font-style: italic;
  }

  .sheet-scroll {
    flex: 1;
    overflow-x: auto;
  }

  .sheet table {
    width: max-content;
    min-width: 100%;
    margin: 0;
    table-layout: fixed;
    border-collapse: separate;
    border-spacing: 0;
    font-size: inherit;
    font-variant-numeric: tabular-nums;
  }

  .gutter-col {
    width: var(--sheet-gutter-w);
  }

  .data-col {
    width: calc(var(--col-w) * var(--sheet-scale));
  }

  .sheet th,
  .sheet td {
    height: var(--sheet-row-h);
    padding: 0 0.45rem;
    border: none;
    border-right: 1px solid var(--sheet-rule);
    border-bottom: 1px solid var(--sheet-rule);
    vertical-align: middle;
    white-space: nowrap;
    overflow: hidden;
  }

  .sheet th,
  .sheet thead td {
    background: var(--sheet-chrome);
    border-color: var(--sheet-rule-strong);
    color: var(--sheet-ink-2);
    font-size: 0.74rem;
    font-weight: 500;
    text-align: center;
  }

  .sheet td.numeric {
    text-align: right;
  }

  .sheet td.bold {
    font-weight: 700;
  }

  .sheet td.rule-above {
    box-shadow: inset 0 1px 0 var(--sheet-ink);
  }

  .sheet td.selected {
    outline: 2px solid var(--sheet-select);
    outline-offset: -2px;
  }

  .sheet-tabs {
    flex: none;
    padding: 0 0.6rem;
    border-top: 1px solid var(--sheet-rule-strong);
    background: var(--sheet-chrome);
    font-size: 0.78rem;
  }

  .sheet-tabs span {
    display: inline-block;
    margin-top: -1px;
    padding: 0.4rem 0.9rem;
    border: 1px solid var(--sheet-rule-strong);
    border-top-color: var(--sheet-paper);
    border-bottom: 2px solid var(--sheet-select);
    background: var(--sheet-paper);
    font-weight: 600;
  }

  .stage-foot {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 1.5rem 3rem;
    padding: clamp(1.25rem, 3vw, 2rem) 0 clamp(1.5rem, 3vw, 2.5rem);
    color: var(--on-accent);
  }

  .stage-foot p {
    max-width: 62ch;
    margin: 0;
    font-size: 1rem;
    line-height: 1.55;
  }

  .stage-foot code {
    white-space: nowrap;
    background: rgb(255 255 255 / 0.16);
    border-color: rgb(255 255 255 / 0.3);
    color: inherit;
  }

  .stage-action {
    flex: none;
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 0.5rem;
    max-width: 22rem;
  }

  .stage-btn {
    background: #fff;
    border-color: #fff;
    color: #15171c;
  }

  .stage-btn:hover {
    background: #15171c;
    border-color: #15171c;
    color: #fff;
  }

  .stage-btn:disabled {
    cursor: progress;
    opacity: 0.8;
  }

  .stage-foot .stage-note {
    font-size: 0.88rem;
    text-align: right;
  }

  .stage :global(:focus-visible) {
    outline-color: #fff;
  }

  /* Three ways in. */
  .paths {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    list-style: none;
    margin: 0;
    padding: 0;
  }

  .paths li {
    display: flex;
    flex-direction: column;
    margin: 0;
    padding: 2.5rem var(--gutter) 2.25rem;
    border-right: 1px solid var(--line);
  }

  .paths li:last-child {
    border-right: none;
  }

  .path-code {
    align-self: flex-start;
    background: var(--accent-wash);
    border-color: transparent;
    color: var(--accent-ink);
    font-size: 0.82rem;
  }

  .paths h2 {
    margin: 1.1rem 0 0.6rem;
    font-size: 1.35rem;
  }

  .paths p {
    margin: 0 0 1.25rem;
    color: var(--ink-2);
    font-size: 0.97rem;
  }

  .paths a {
    margin-top: auto;
    font-weight: 550;
  }

  /* Text beside code; `.reverse` puts the code first so two in a row alternate. */
  .split {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    gap: clamp(2rem, 5vw, 4.5rem);
    align-items: center;
    padding: clamp(3rem, 7vw, 5.5rem) var(--gutter);
  }

  .split.reverse .split-text {
    order: 1;
  }

  .split-text h2 {
    margin: 0 0 1.25rem;
    font-size: clamp(1.7rem, 3.4vw, 2.5rem);
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .split-text p {
    color: var(--ink-2);
    max-width: 52ch;
  }

  .split-code {
    border-radius: var(--radius);
    background: var(--night);
    border: 1px solid var(--night-line);
    overflow: hidden;
  }

  .split-code :global(pre) {
    margin: 0;
    border: none;
    border-radius: 0;
  }

  /* Section heads shared by the lower sections. */
  .section-head {
    max-width: 70ch;
    padding: clamp(3rem, 7vw, 5rem) var(--gutter) 0;
  }

  .section-head h2 {
    margin: 0 0 1rem;
    font-size: clamp(1.7rem, 3.4vw, 2.5rem);
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .section-head p {
    margin: 0;
    color: var(--ink-2);
    font-size: 1.04rem;
  }

  /* Subpath entries. */
  .entries {
    padding-bottom: clamp(3rem, 6vw, 4.5rem);
  }

  .entries .table-scroll {
    margin-top: 2.25rem;
    padding: 0 var(--gutter);
  }

  .entries table {
    min-width: 600px;
    margin: 0;
  }

  .entries thead th {
    padding-bottom: 0.8rem;
    font-size: 0.85rem;
  }

  .entries tbody th {
    border-bottom-color: var(--line);
    font-weight: 400;
  }

  .entries tbody th code {
    white-space: nowrap;
    background: none;
    border: none;
    padding: 0;
    color: var(--accent-ink);
    font-weight: 500;
  }

  .entries td {
    color: var(--ink-2);
  }

  .entries .num {
    padding-right: 0;
    text-align: right;
    font-variant-numeric: tabular-nums;
    white-space: nowrap;
  }

  .entries th.num + th.num,
  .entries td.num + td.num {
    padding-left: 1.5rem;
  }

  .entries-more {
    max-width: 70ch;
    margin: 1.5rem 0 0;
    padding: 0 var(--gutter);
    color: var(--ink-2);
  }

  /* Proof: the one dark band on the page. */
  .proof {
    background: var(--night);
    color: var(--night-ink);
    border-bottom-color: var(--night-line);
  }

  .proof-inner {
    border-inline-color: var(--night-line);
    padding: clamp(3rem, 7vw, 5.5rem) var(--gutter);
  }

  .proof h2 {
    margin: 0;
    color: #fff;
    font-size: clamp(1.9rem, 4.2vw, 3.1rem);
    letter-spacing: -0.035em;
    line-height: 1.05;
  }

  .proof-lede {
    max-width: 58ch;
    margin: 1.1rem 0 2.75rem;
    color: var(--night-ink-2);
    font-size: 1.08rem;
  }

  .proof dl {
    margin: 0;
    border-top: 1px solid var(--night-line);
  }

  .proof-row {
    display: grid;
    grid-template-columns: minmax(0, 5fr) minmax(0, 6fr);
    gap: 0.5rem 3rem;
    padding: 1.5rem 0;
    border-bottom: 1px solid var(--night-line);
  }

  .proof dt {
    font-family: var(--display);
    font-size: 1.25rem;
    font-weight: 550;
    line-height: 1.25;
    letter-spacing: -0.015em;
    color: #fff;
    text-wrap: balance;
  }

  .proof dd {
    margin: 0;
    color: var(--night-ink-2);
    font-size: 0.98rem;
  }

  /* Capabilities. */
  .caps-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    margin: 2.5rem 0 0;
    border-top: 1px solid var(--line);
  }

  .caps-grid > div {
    padding: 1.5rem var(--gutter) 1.6rem;
    border-right: 1px solid var(--line);
    border-bottom: 1px solid var(--line);
  }

  .caps-grid > div:nth-child(4n) {
    border-right: none;
  }

  .caps-grid dt {
    font-family: var(--display);
    font-size: 1.12rem;
    font-weight: 600;
    letter-spacing: -0.01em;
  }

  .caps-grid dd {
    margin: 0.45rem 0 0;
    color: var(--ink-2);
    font-size: 0.93rem;
    line-height: 1.5;
  }

  .not-yet {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 3fr);
    gap: 0.75rem 2rem;
    padding: 1.75rem var(--gutter) 2.25rem;
    background: var(--wash);
  }

  .not-yet h3 {
    margin: 0;
    font-size: 1.12rem;
  }

  .not-yet ul {
    margin: 0;
    padding: 0;
    list-style: none;
    columns: 2;
    column-gap: 2.5rem;
    color: var(--ink-2);
    font-size: 0.93rem;
  }

  .not-yet li {
    margin: 0 0 0.45rem;
    break-inside: avoid;
  }

  /* Agents. */
  .agents-inner {
    display: grid;
    grid-template-columns: minmax(0, 5fr) minmax(0, 6fr);
    gap: clamp(2rem, 5vw, 4.5rem);
    padding: clamp(3rem, 7vw, 5rem) var(--gutter);
  }

  .agents h2 {
    margin: 0 0 1rem;
    font-size: clamp(1.7rem, 3.4vw, 2.5rem);
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .agents p {
    color: var(--ink-2);
    max-width: 50ch;
  }

  .agents ul {
    list-style: none;
    margin: 0;
    padding: 0;
    border-top: 1px solid var(--line);
  }

  .agents li {
    display: grid;
    grid-template-columns: 11.5rem minmax(0, 1fr);
    gap: 0.25rem 1rem;
    align-items: baseline;
    margin: 0;
    padding: 0.95rem 0;
    border-bottom: 1px solid var(--line);
  }

  .agents li span {
    color: var(--ink-2);
    font-size: 0.95rem;
  }

  .family-head {
    padding-bottom: 2.5rem;
    border-bottom: 1px solid var(--line);
    max-width: none;
  }

  .family-head p {
    max-width: 66ch;
  }

  @media (max-width: 960px) {
    .stage-grid {
      grid-template-columns: minmax(0, 1fr);
    }

    /* Sheet first on narrow screens: it is the payoff, the code is the proof. */
    .sheet-pane {
      order: -1;
    }

    .code-scroll {
      height: 300px;
      flex: none;
    }

    .paths {
      grid-template-columns: 1fr;
    }

    .paths li {
      border-right: none;
      border-bottom: 1px solid var(--line);
      padding-block: 2rem;
    }

    .paths li:last-child {
      border-bottom: none;
    }

    .split,
    .agents-inner {
      grid-template-columns: minmax(0, 1fr);
    }

    .split.reverse .split-text {
      order: 0;
    }

    .caps-grid {
      grid-template-columns: repeat(2, 1fr);
    }

    .caps-grid > div:nth-child(2n) {
      border-right: none;
    }
  }

  @media (max-width: 720px) {
    .stage-foot {
      flex-direction: column;
    }

    .stage-action {
      align-items: flex-start;
      max-width: none;
    }

    .stage-foot .stage-note {
      text-align: left;
    }

    .proof-row {
      grid-template-columns: 1fr;
    }

    .not-yet {
      grid-template-columns: 1fr;
    }

    /* A four-column table would scroll its sizes off a phone, and the sizes
     * are the point, so each row stacks and labels its own numbers. */
    .entries table {
      display: block;
      min-width: 0;
    }

    .entries thead {
      position: absolute;
      width: 1px;
      height: 1px;
      overflow: hidden;
      clip-path: inset(50%);
    }

    .entries tbody {
      display: block;
      border-top: 1px solid var(--line-strong);
    }

    .entries tbody tr {
      display: grid;
      grid-template-columns: auto 1fr;
      gap: 0.2rem 1.5rem;
      padding: 0.9rem 0;
      border-bottom: 1px solid var(--line);
    }

    .entries tbody th,
    .entries tbody td {
      padding: 0;
      border: none;
    }

    .entries tbody th,
    .entries .what {
      grid-column: 1 / -1;
    }

    .entries tbody .num {
      padding-left: 0;
      text-align: left;
    }

    .entries .num::before {
      content: attr(data-label) ' ';
      color: var(--ink-3);
    }

    .not-yet ul {
      columns: 1;
    }
  }

  @media (max-width: 520px) {
    /* Scaled down so the four columns that hold data fit a phone without scrolling. */
    .sheet {
      --sheet-scale: 0.8;
      --sheet-gutter-w: 30px;
      font-size: 0.78rem;
    }

    .sheet th,
    .sheet td {
      padding: 0 0.35rem;
    }

    .cta .btn {
      flex: 1 1 100%;
    }

    .caps-grid {
      grid-template-columns: 1fr;
    }

    .caps-grid > div {
      border-right: none;
    }

    .agents li {
      grid-template-columns: 1fr;
    }
  }
</style>
