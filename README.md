<p align="center">
  <img src="./assets/logo.png" alt="@office-kit/xlsx" width="180" height="180" />
</p>

# @office-kit/xlsx

A TypeScript library for reading and writing Excel `.xlsx` workbooks
from Node 22+ and modern browsers, with no runtime dependencies on
Python or Excel. Inspired by [openpyxl](https://openpyxl.readthedocs.io/).

> **Status: pre-1.0 alpha.** The core read / write / streaming pipeline is
> in place and round-trips real-world fixtures (including pivot tables and
> macro-enabled `.xlsm`), but APIs may shift before `1.0`.

## Why @office-kit/xlsx?

The JavaScript xlsx ecosystem in 2026 is split between **commercial upsell
tiers** and **stalled open-source projects**. SheetJS Community Edition
deliberately omits styling, charts, images, pivots, conditional formatting,
and data validation on write — those live in [SheetJS Pro][sjs-pro], a paid
tier. ExcelJS is MIT but
[has not had a meaningful release since October 2023][exceljs-discussion]
and its maintainers explicitly call it inactive; the dependency footprint
unpacks to 21.8 MB. excel4node was [archived in 2022][e4n-archive].
xlsx-js-style is frozen at a 2022 SheetJS fork.

@office-kit/xlsx is the third option: an actively-developed, pure-MIT,
TypeScript-first library with no Pro tier and no missing features behind a
paywall.

[sjs-pro]: https://sheetjs.com/pro/
[exceljs-discussion]: https://github.com/exceljs/exceljs/discussions/2987
[e4n-archive]: https://github.com/natergj/excel4node

| Concern                | Other libraries                                                                  | @office-kit/xlsx                                                              |
|------------------------|----------------------------------------------------------------------------------|-----------------------------------------------------------------------|
| TypeScript types       | hand-written `.d.ts` retrofitted (SheetJS) or community typings (xlsx-populate, excel4node) | first-party, written in TS under `exactOptionalPropertyTypes` + `noUncheckedIndexedAccess` |
| Bundle size            | ExcelJS unpacks to 21.8 MB; xlsx ~7.5 MB                                         | full lib ≤120 KB min+brotli (currently ~85 KB); streaming entry ~49 KB |
| Streaming              | SheetJS docs explicitly note the zip central-directory layout prevents true streaming; ExcelJS supports both directions but the lib is heavy | both read iter and write append, with fixed-memory budget for tens of millions of rows |
| Charts (write)         | none in ExcelJS, xlsx-js-style, SheetJS CE; gated behind SheetJS Pro             | 16 legacy `c:` + 8 modern `cx:` chart kinds (Sunburst, Treemap, Waterfall, Histogram, Pareto, Funnel, BoxWhisker, RegionMap) |
| Pivots / VBA / OLE     | ExcelJS drops pivot tables on read ([#261][exceljs-pivot]); others vary           | byte-identical passthrough so Excel 365 still renders parts we don't model |
| Maintenance            | ExcelJS stalled since 2023; excel4node archived 2022; xlsx-js-style frozen 2022; SheetJS npm artifact frozen 2022 (still distributed via private CDN) | active                                                                |
| License                | SheetJS CE strips features for Pro upsell; SheetJS Pro pricing not published      | MIT, single tier, no upsell                                            |
| Conformance            | none of the major libraries validate against ECMA-376 in CI                       | 3-tier validator (OPC structure + ECMA-376 XSD + semantic invariants) gates every CI build, including a fast-check property-based oracle |
| Modules                | monolithic root barrel                                                           | subpath imports — `@office-kit/xlsx/io`, `/streaming`, `/cell`, `/styles`, etc., each independently tree-shakable |

[exceljs-pivot]: https://github.com/exceljs/exceljs/issues/261

### Where each existing library still wins

- **Read simple xlsx in the browser** → [`read-excel-file`][rexf] is excellent.
- **Write simple xlsx with images** → [`write-excel-file`][wexf] is excellent.
- **Template-based fidelity preservation with password protection** → `xlsx-populate`.
- **Non-xlsx formats** (XLS / XLSB / ODS / CSV / HTML) → SheetJS Community.
- **Commercial budget + long shopping list** → SheetJS Pro.

[rexf]: https://www.npmjs.com/package/read-excel-file
[wexf]: https://www.npmjs.com/package/write-excel-file

### @office-kit/xlsx's home turf

- You write modern TypeScript and want types that actually behave under
  strict mode (cell values are a discriminated union, not `any`).
- You produce **large** xlsx files (tens of millions of cells) and care
  about heap budget.
- You need **charts**, conditional formatting, data validation, defined
  names, tables, ZIP64 (entry-count overflow; see "What's supported" for the
  4 GiB-per-entry caveat) — and want them in MIT.
- You round-trip xlsx files that contain pivot tables, VBA macros,
  threaded comments, Power Query metadata, or customXml — and need them
  preserved byte-for-byte.
- You want **proof** that the bytes you emit are valid OOXML, not "Excel
  happens to open them today."

### When NOT to use @office-kit/xlsx

Honest list:

- **Pre-1.0**: API may shift before 1.0. Pin the version for long-running
  projects.
- **`.xlsx` only**: no `.xls` (BIFF), `.xlsb`, `.ods`, or `.csv`. Use
  SheetJS for those.
- **Node 22+ required**: relies on built-in `Web Streams`, `Blob`, and
  `fetch`. Node 18 / 20 (EOL) are not supported.
- **Browser stress-test history is shorter** than ExcelJS's. If you ship
  to millions of browser users today, run your own benchmark first.
- **Visual QA in Excel 365** is on the human-verification list; the schema
  gate proves spec compliance, not that every chart renders pixel-perfect.

### Motivation

The reasons @office-kit/xlsx exists, written down so future contributors don't
relitigate them:

1. **The reference implementation is in Python.** [openpyxl][openpyxl] has
   spent 15 years collecting Excel / LibreOffice corner cases. @office-kit/xlsx
   consumes its fixture corpus directly (`reference/openpyxl/` is a git
   submodule), so edge cases the Python world solved years ago don't get
   re-discovered painfully in JS.
2. **The 2010-era JS stack is heavy.** Most existing libraries pull in
   `jszip`, `lodash`, `archiver`, `xmlbuilder`, `sax`. In 2026 we have
   `fflate`, `fast-xml-parser`, and `saxes` — the toolchain is an order
   of magnitude lighter. @office-kit/xlsx ships with three runtime dependencies.
3. **TypeScript-first changes the API surface.** A library authored in TS
   under strict-mode flags from day one exposes different ergonomics than
   `.d.ts` typings retrofitted onto an old JS codebase.
4. **"Schema-valid" should be a CI gate, not a vibe.** ECMA-376 is
   downloadable; xmllint is free; vendoring the schemas costs <1 MB. There
   is no good reason a 2026 library shouldn't validate every byte it
   emits against the spec.
5. **No Pro tier.** Charts, pivots passthrough, conditional formatting,
   ZIP64 write — all MIT. Nothing held back.

[openpyxl]: https://openpyxl.readthedocs.io/

## Install

```sh
pnpm add @office-kit/xlsx   # or npm / yarn / bun
```

Requires Node `>=22` for the built-in `Web Streams`, `Blob`, and `fetch`
globals.

## Subpath entries

The package has no root barrel — every export lives behind a section
subpath, so your editor's autocomplete only surfaces what's relevant to
the area you're working in. Each export has exactly one home (no
convenience re-exports).

| Import                 | Use case                                          |
|------------------------|---------------------------------------------------|
| `@office-kit/xlsx/io`           | `loadWorkbook` / `saveWorkbook` / `workbookToBytes` plus byte-level Source/Sink + browser helpers (Blob/Response/Stream) |
| `@office-kit/xlsx/node`         | Node fs glue (`fromFile` / `toFile` / `fromBuffer` / `toBuffer` / `fromReadable` / `toWritable`) |
| `@office-kit/xlsx/streaming`    | Read-only iter (`loadWorkbookStream`) + write-only append (`createWriteOnlyWorkbook`) |
| `@office-kit/xlsx/workbook`     | `createWorkbook`, `addWorksheet`, defined names   |
| `@office-kit/xlsx/worksheet`    | `setCell`, `getCell`, `mergeCells`, tables, …     |
| `@office-kit/xlsx/cell`         | Cell value-model + inline rich text               |
| `@office-kit/xlsx/styles`       | Fonts, fills, borders, alignment, number formats  |
| `@office-kit/xlsx/chart`        | `c:` and `cx:` chart kinds                        |
| `@office-kit/xlsx/chartsheet`   | Standalone chartsheets                            |
| `@office-kit/xlsx/drawing`      | Anchors, images, chart placement                  |

Other subpaths: `@office-kit/xlsx/packaging`, `@office-kit/xlsx/utils`, `@office-kit/xlsx/xml`,
`@office-kit/xlsx/zip`, `@office-kit/xlsx/schema`. All exports are tree-shakable
(`"sideEffects": false`).

Bundle budgets (min + brotli):

- `@office-kit/xlsx/streaming` ≤ 80 KB    (currently ~49 KB)
- `@office-kit/xlsx/io` ≤ 120 KB           (currently ~85 KB)

## Quick examples

For a one-page lookup of task → exact functions to import and call, see
the [Cheatsheet](https://baseballyama.github.io/@office-kit/xlsx/docs/cheatsheet).
For prose-style worked examples (styling, charts, validation, streaming),
see the [Recipes](https://baseballyama.github.io/@office-kit/xlsx/docs/recipes).

### Read + edit + write

```ts
import { loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
import { setCell } from '@office-kit/xlsx/worksheet';
import { fromBuffer } from '@office-kit/xlsx/node';
import { readFile, writeFile } from 'node:fs/promises';

const wb = await loadWorkbook(fromBuffer(await readFile('input.xlsx')));
const sheet = wb.sheets[0];
if (sheet?.kind === 'worksheet') {
  setCell(sheet.sheet, /* row */ 1, /* col */ 1, 'Hello from @office-kit/xlsx');
}
await writeFile('output.xlsx', await workbookToBytes(wb));
```

### Read directly from disk (Node)

```ts
import { loadWorkbook, saveWorkbook } from '@office-kit/xlsx/io';
import { fromFile, toFile } from '@office-kit/xlsx/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// …mutate wb…
await saveWorkbook(wb, toFile('output.xlsx'));
```

### Read directly from a `fetch` response (browser)

```ts
import { fromResponse, loadWorkbook } from '@office-kit/xlsx/io';

const response = await fetch('/sheet.xlsx');
const wb = await loadWorkbook(fromResponse(response));
```

### Streaming write — millions of rows in a fixed memory budget

```ts
import { createWriteOnlyWorkbook } from '@office-kit/xlsx/streaming';
import { toFile } from '@office-kit/xlsx/node';

const sink = toFile('big.xlsx');
const wb = await createWriteOnlyWorkbook(sink);
const ws = await wb.addWorksheet('Data');
ws.setColumnWidth(1, 24); // must precede the first appendRow
for (let r = 0; r < 10_000_000; r++) {
  await ws.appendRow([r, `row-${r}`, r * Math.PI]);
}
await ws.close();
await wb.finalize();
```

The streaming writer pushes each row through deflate as it arrives, and
`toFile` forwards each deflated chunk to disk (honouring write-stream
backpressure) — peak memory stays at one pending-row buffer plus deflate
scratch, regardless of total archive size. The same is true of `toWritable`;
buffered sinks (`toBuffer` / `toBlob` / `toArrayBuffer`) instead keep the
full archive resident so `result()` can hand it back in one piece.

### Streaming read — iterate row-by-row without materialising the sheet

```ts
import { loadWorkbookStream } from '@office-kit/xlsx/streaming';
import { fromFile } from '@office-kit/xlsx/node';

const wb = await loadWorkbookStream(fromFile('big.xlsx'));
const sheet = wb.openWorksheet(wb.sheetNames[0] ?? '');
for await (const row of sheet.iterRows({ minRow: 1, maxRow: 100 })) {
  console.log(row.map((c) => c.value));
}
await wb.close();
```

The whole-sheet iteration path (default / `minRow <= 1`) inflates the
worksheet entry chunk-by-chunk straight into the SAX parser, so the inflated
worksheet body is never fully resident. Note: ZIP requires random access to
its central directory, so the **compressed archive bytes are loaded up
front**. A 200 MB compressed xlsx therefore needs ~200 MB resident, plus the
inflate window + SAX state per active iterator — not the multi-GB inflated
worksheet payload. Band queries (`minRow > 1`) build a row-offset index once
per sheet, which does materialise that sheet's inflated bytes; subsequent
band queries reuse the cached index.

### Migrating from openpyxl

@office-kit/xlsx is shaped after openpyxl, but a few defaults differ. The most
common surprise for direct ports:

- **`createWorkbook()` returns an empty workbook with no sheets.**
  `openpyxl.Workbook()` creates a default sheet named `Sheet` that
  callers usually remove with `wb.remove(wb.active)`. @office-kit/xlsx skips
  that step — call `addWorksheet(wb, 'Data')` directly. Translating a
  `remove(active)` call literally produces a no-op (or, worse, a guard
  that hides a real bug elsewhere).
- **`setCell(ws, row, col, value)`** is the @office-kit/xlsx equivalent of
  openpyxl's `ws.cell(row=r, column=c, value=v)`. Coordinates are
  1-based on both sides.
- **`makeBorder({ left: makeSide({ style: 'thin' }) })`** is the
  @office-kit/xlsx equivalent of openpyxl's
  `Border(left=Side(style='thin'))`. Same with `makeFill`, `makeFont`,
  etc. — every style primitive has a `make*` constructor under
  `@office-kit/xlsx/styles`.

## What's supported

- ✅ Cell values: number, string (sharedStrings), boolean, error, formulas
  (normal / array / shared / dataTable), inline rich text
- ✅ Styles: Font, Fill, Border, Alignment, Protection, NumberFormat, full
  Stylesheet pool with dedup, named styles + DXF
- ✅ Worksheet rich features: mergedCells, sheetView/freezePanes, columnDims,
  rowDims, hyperlinks, defined names, data validations, autoFilter, Tables,
  legacy comments, conditional formatting
- ✅ Drawings: anchors, images (PNG/JPEG/GIF/BMP/WebP/TIFF/SVG/EMF/WMF) with
  format + dimension auto-detection, picture frames in worksheets and charts
- ✅ Charts: 16 legacy `c:` chart kinds + 8 `cx:` chartex kinds (Sunburst,
  Treemap, Waterfall, Histogram, Pareto, Funnel, BoxWhisker, RegionMap),
  spPr / txPr / dLbls / trendline / errBars wiring, chartsheets, UserShapes
- ✅ Pivot tables / VBA / OLE / threaded comments / external links / Power
  Query metadata / customXml / customUI: byte-identical passthrough so
  Excel 365 still renders parts we don't model. The `<workbook>` body
  extras and per-sheet rels chain are preserved end-to-end.
- ✅ Encrypted xlsx detection (CFB Compound Document magic): clear error
  pointing at `msoffcrypto-tool` for decryption.
- ✅ ZIP64 write — partial: workbooks with > 65 535 entries get a ZIP64 EOCD
  record + locator spliced into the final chunk. Read works too. **Limit:**
  individual entry sizes and the central-directory offset must still fit in
  32 bits (≤ 4 GiB each); xlsx archives never approach that in practice, but
  if you genuinely need a single >4 GiB entry the writer will throw. Tracked
  in [`src/zip/zip64-patch.ts`][zip64-patch].

[zip64-patch]: ./src/zip/zip64-patch.ts

## Development

The test suite reads fixtures from the `reference/openpyxl` git submodule, so
clone with submodules (or run `pnpm install`, which auto-inits via the
`prepare` script):

```sh
git clone --recursive https://github.com/office-kit/xlsx.git
# or, if you already cloned without --recursive:
git submodule update --init --recursive

pnpm install
pnpm typecheck
pnpm lint
pnpm test          # vitest, ~2100 tests
pnpm test:perf     # write-only throughput + heap-budget bench
pnpm build         # tsdown + tsc → dist/
pnpm size          # size-limit guards on each bundle
```

[Nix flake](flake.nix) included — `nix develop` (or [direnv](https://direnv.net/)
with `use flake`) gives a pinned Node 22 + pnpm 10 + Python 3 environment.

## License

MIT — see [LICENSE](LICENSE) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
