# @office-kit/xlsx

## 0.9.2

### Patch Changes

- [#108](https://github.com/office-kit/xlsx/pull/108) [`b28e960`](https://github.com/office-kit/xlsx/commit/b28e960474d56ac28abcc0bb86ea68d6d1140c49) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: round-tripping a sheet with form controls produced a package Excel refused ([#105](https://github.com/office-kit/xlsx/issues/105))

  Loading a worksheet that carries a form control (`<legacyDrawing>` + the
  x14-gated `<controls>` block Excel 2010+ writes) and saving it again wrote a
  file Excel reported as corrupt:

  - `<legacyDrawing r:id>` and its `vmlDrawing` relationship were dropped on
    sheets without comments — only the comment VML was ever re-linked — leaving
    the control's VML shape orphaned;
  - Excel's `<mc:AlternateContent><mc:Choice Requires="x14">` wrapper around
    `<controls>` / `<oleObjects>` was passed through as an opaque node, landing
    before `<drawing>` (out of `CT_Worksheet` order) with the `x14` prefix
    undeclared.

  `Worksheet` now exposes `legacyDrawingRId` next to `legacyDrawingHFRId`; any
  VML that is not a comment overlay keeps its relationship and part across a
  round-trip, and the `Requires="x14"` wrapper is read into the typed
  `controls` / `oleObjects` model and written back in schema order with
  `xmlns:x14` declared on the worksheet root.

## 0.9.1

### Patch Changes

- [#106](https://github.com/office-kit/xlsx/pull/106) [`994935e`](https://github.com/office-kit/xlsx/commit/994935ee6c6b736fe608b035a4e771103d524029) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: round-tripping a sheet with a header/footer picture (`<legacyDrawingHF>`) produced a file Excel refused to open ([#104](https://github.com/office-kit/xlsx/issues/104))

  `loadWorkbook` → `saveWorkbook` re-emitted `<legacyDrawingHF r:id="…"/>` but dropped the `vmlDrawing` relationship it points at, the VML part's own `.rels`, and the image behind it, and the surviving `.vml` part had no content type. Parts referenced by relationships the writer re-emits verbatim are now carried over together with their own relationships (transitively), and `<Default>` content types from the source manifest are preserved for such parts.

## 0.9.0

### Minor Changes

- [#100](https://github.com/office-kit/xlsx/pull/100) [`c74f9d0`](https://github.com/office-kit/xlsx/commit/c74f9d057f766a14c6dc26b41e57682d611e17ca) Thanks [@baseballyama](https://github.com/baseballyama)! - Renamed the package from `xlsx-kit` to `@office-kit/xlsx`. The old `xlsx-kit`
  package on npm is deprecated and will receive no further releases — install
  `@office-kit/xlsx` and update every subpath import (`xlsx-kit/io` →
  `@office-kit/xlsx/io`, and likewise for `/streaming`, `/workbook`, `/worksheet`,
  `/cell`, `/styles`, `/chart`, `/drawing`, `/node`, and the low-level `/xml`,
  `/zip`, `/packaging`, `/schema` escape hatches). No runtime behaviour changed.

### Patch Changes

- [#102](https://github.com/office-kit/xlsx/pull/102) [`9b1feba`](https://github.com/office-kit/xlsx/commit/9b1febad6bb7a1f6d0279e0bb9fceea2eaa2a36b) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: `loadWorkbook` rejected borders with `style="none"` ([#99](https://github.com/office-kit/xlsx/issues/99))

  Files written by OnlyOffice (and any producer that emits the explicit no-border
  value) set border sides to `style="none"`, which is the first value of
  ECMA-376's `ST_BorderStyle` enumeration. Loading such a file failed with
  `expected one of [thin, medium, ...]; got "none"`. `none` is now accepted and
  round-trips faithfully; it draws no stroke in HTML/SVG preview.

- [#101](https://github.com/office-kit/xlsx/pull/101) [`06dbfd8`](https://github.com/office-kit/xlsx/commit/06dbfd868e30eb6056114f696c8fbe8932b93ceb) Thanks [@michaelcradock76-spec](https://github.com/michaelcradock76-spec)! - Preserve the `_xlfn.` future-function prefix on read so dynamic-array formulas
  survive a load → save round-trip. The reader stripped `_xlfn.` / `_xlfn._xlws.`
  into the model, and the writer emits formula text verbatim, so a loaded
  `_xlfn.SCAN(...)` was written back as bare `SCAN(...)` — an unknown name that
  Excel renders as `#NAME?`. Formula text is now kept verbatim (matching openpyxl),
  fixing every future function (SCAN, BYCOL, LAMBDA, XLOOKUP, FILTER, SEQUENCE,
  LET, ANCHORARRAY, …).

## 0.8.0

### Minor Changes

- [#78](https://github.com/office-kit/xlsx/pull/78) [`2a46931`](https://github.com/office-kit/xlsx/commit/2a469317466508c8f18fd6f8181f70bacf42b051) Thanks [@baseballyama](https://github.com/baseballyama)! - **Breaking**: `Chartsheet.properties.tabColor` replaces `tabColorRgb`.

  Worksheets already expose `SheetProperties.tabColor` as a full `Color` (rgb / indexed / theme / auto / tint). Chartsheets carried a stringly-typed `tabColorRgb` instead, which forced callers to special-case the two sheet kinds and silently dropped every non-RGB colour attribute Excel produces.

  The new field is a `Color` object so both sheet kinds share one tab-colour model. Migration:

  ```ts
  // Before
  cs.properties = { tabColorRgb: "FF8800" };

  // After
  cs.properties = { tabColor: { rgb: "FF8800" } };
  ```

  Reads recover the additional `indexed` / `theme` / `auto` / `tint` attributes that the old shape discarded.

- [#78](https://github.com/office-kit/xlsx/pull/78) [`2a46931`](https://github.com/office-kit/xlsx/commit/2a469317466508c8f18fd6f8181f70bacf42b051) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: release streaming sinks when serialization fails, and bound the streaming write-only buffer against CJK payloads.

  When `saveWorkbook` or `createWriteOnlyWorkbook().finalize()` threw partway through serialization, the underlying sink stayed open — `toFile` would leave a half-written `.xlsx` on disk that callers could mistake for a successful save. `BufferedSinkWriter` now exposes an optional `abort(cause?)` hook; the ZIP writer + workbook writers call it from a surrounding `catch` so streaming destinations (`toFile` / `toWritable`) are released and the partial file is best-effort removed. The write-only workbook also exposes its own `abort(cause?)` for callers driving it from a custom pipeline.

  The streaming write-only worksheet's pending-byte counter and the XML stream writer's flush threshold now use an accurate UTF-8 length (`utf8ByteLength`) instead of `string.length`. The previous accounting undercounted CJK text by ~3×, letting the in-flight buffer grow well past the configured flush threshold; Japanese / Chinese workloads now flush at the documented ~64 KB ceiling.

  API:

  - `BufferedSinkWriter.abort(cause?)` is optional. Custom sinks don't need to implement it, but doing so lets `saveWorkbook` clean up streaming destinations on failure.
  - `WriteOnlyWorkbook.abort(cause?)` is the matching escape hatch on the write-only API.
  - `XlsxSink.toBytes` is now required at the type level. It was already required in practice — the writer threw at runtime when it was missing — and built-in sinks (`toBuffer` / `toBlob` / `toArrayBuffer` / `toFile` / `toWritable`) already implement it.

- [#78](https://github.com/office-kit/xlsx/pull/78) [`2a46931`](https://github.com/office-kit/xlsx/commit/2a469317466508c8f18fd6f8181f70bacf42b051) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: tighten three worksheet / cell mutation paths that previously produced silently-bad workbooks.

  - `copyRange` / `moveRange` no longer carry `hyperlinkId` / `commentId` across worksheets. Those fields are indexes into the source sheet's `hyperlinks` / `legacyComments` arrays and would point at unrelated records (or out of bounds) on the destination. Same-sheet copy / move still preserves them.

  - `setSheetState` refuses to hide the last visible sheet. Excel rejects workbooks with every sheet hidden ("Excel cannot use the object linking and embedding features…"); raising at mutation time keeps the workbook recoverable instead of producing a save Excel will reject.

  - `makeTextRun` throws `OpenXmlSchemaError` instead of `TypeError` so every public error path uses the documented `OpenXmlError` subclass hierarchy.

### Patch Changes

- [#78](https://github.com/office-kit/xlsx/pull/78) [`2a46931`](https://github.com/office-kit/xlsx/commit/2a469317466508c8f18fd6f8181f70bacf42b051) Thanks [@baseballyama](https://github.com/baseballyama)! - chore(ci): tighten the CI gate and surface CLAUDE.md hard rules in the linter.

  - The test matrix now runs on macOS and Windows alongside Ubuntu. Path-separator and TextDecoder differences that ubuntu-only CI would silently miss are now exercised on every PR. Each OS installs `libxml2-utils` / `libxml2` so the ECMA-376 conformance gate never silently degrades to "no schema validation ran".

  - A dedicated `perf` job runs `pnpm test:perf` with `PERF_GATE=1`, promoting the throughput / heap thresholds in `tests/perf/` from informational to fatal. CI runners are noisier than the M1 baseline the gates target — if a specific Node minor turns flaky, retune the thresholds in `vitest.perf.config.ts` rather than reverting this gate.

  - `typescript/no-explicit-any` is `error` in `src/` and `off` in `tests/` (where `as any` is a legitimate way to exercise error paths). The remaining intentional `any` in `src/schema/core.ts` is annotated with `// oxlint-disable-next-line typescript/no-explicit-any` and an explanation comment.

- [#78](https://github.com/office-kit/xlsx/pull/78) [`2a46931`](https://github.com/office-kit/xlsx/commit/2a469317466508c8f18fd6f8181f70bacf42b051) Thanks [@baseballyama](https://github.com/baseballyama)! - refactor: route every XML writer through one canonical `escapeXmlAttr` / `escapeXmlText` pair in `src/utils/escape.ts`.

  Before, twelve files each carried their own near-identical escape regex — `src/io/save.ts` deliberately skipped `>` while the rest escaped it, and three of them also escaped `\r` / `\n` / `\t` via numeric character references. The discrepancies were quiet correctness bugs (attribute values containing `]]>` rendered differently across writers) and a maintenance hazard.

  The unified helpers escape `&`, `<`, `>`, and `"`; whitespace bytes stay literal because our parser (`fast-xml-parser`) does not decode numeric character references and would otherwise break the round-trip.

  No user-visible behaviour change beyond consistent attribute escaping across every writer.

- [#78](https://github.com/office-kit/xlsx/pull/78) [`2a46931`](https://github.com/office-kit/xlsx/commit/2a469317466508c8f18fd6f8181f70bacf42b051) Thanks [@baseballyama](https://github.com/baseballyama)! - perf: drop quadratic / linear-scan patterns on four read/write hot paths.

  - The `loadWorkbook` resolver indexes each sheet's rels file once via the new `indexRelsById` helper instead of running a fresh `Array.find(r => r.id === relId)` per table / comments / drawing / chart / picture cross-reference. Worksheets with many drawings or pivot tables load in O(refs) instead of O(refs × rels).

  - `containsCommentMarker` in the VML drawing classifier replaces a byte-by-byte JS loop with a single latin1 `String.indexOf` — multi-megabyte legacy VML drawings load noticeably faster.

  - The streaming `iterParse` queue uses a head pointer instead of `Array#shift()`. A single SAX batch with hundreds of events (typical for a wide `<row>`) is now O(N) instead of O(N²); a stale dead-code branch in the prologue gate is gone too.

  - `serializeHyperlinks` allocates rIds via a `Set` index instead of nested `Array.some()` calls. Worksheets with hundreds of hyperlinks (dashboards / link-heavy index sheets) finish save in O(N) rather than O(N²).

- [#78](https://github.com/office-kit/xlsx/pull/78) [`2a46931`](https://github.com/office-kit/xlsx/commit/2a469317466508c8f18fd6f8181f70bacf42b051) Thanks [@baseballyama](https://github.com/baseballyama)! - refactor: collapse `describeWorkbook` into a single pass, harden `getRowValues`, and document `normalizePath`'s `..` handling.

  `describeWorkbook` used to walk every cell three times — once for `getWorkbookStats`, once for `getWorkbookCellsByKind`, and once again for the per-sheet counts. The new implementation fuses all three into one pass and shares the value-kind classifier with `countCellsByKind` via the newly exported `classifyCellValue` (so the two never drift on edge cases like `Date` vs `duration`).

  `getRowValues` no longer derives `maxCol` via `Math.max(...rowMap.keys())`. The spread is limited by V8's argument-count cap (~125 k), which `getRowValues` would silently blow past on a dense row near MAX_COL — the call now uses a linear scan to derive the max.

  `normalizePath` picks up an explanatory comment for the `..`-when-`out`-is-empty case, so the next reader doesn't have to second-guess whether path-traversal is possible (it isn't — the archive lookup catches escape attempts naturally).

- [#78](https://github.com/office-kit/xlsx/pull/78) [`2a46931`](https://github.com/office-kit/xlsx/commit/2a469317466508c8f18fd6f8181f70bacf42b051) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: harden the ZIP64 read path against decompression bombs.

  Previously, archives whose End-of-Central-Directory carried ZIP64 sentinel values fell back to `fflate.unzipSync`, which inflates every entry up front. A crafted ZIP64-shaped xlsx could exhaust memory before the per-entry / archive-total caps ran. The random-access reader now parses ZIP64 EOCD + the Zip64 Extended Information extra field directly, so the existing decompression-bomb guards apply to ZIP64 archives the same way they apply to ZIP32.

  `loadWorkbook(decompressionLimits)` defaults are unchanged; only the underlying enforcement path is stricter.

## 0.7.1

### Patch Changes

- [#76](https://github.com/office-kit/xlsx/pull/76) [`e646e1c`](https://github.com/office-kit/xlsx/commit/e646e1c7217ce07a1d6357b5327f006002d86c21) Thanks [@baseballyama](https://github.com/baseballyama)! - `addAutoFilterColumn`, `makeHyperlink`, and `setPrintTitles` now throw
  `OpenXmlSchemaError` instead of the generic `Error` when their preconditions
  are violated. Existing catch blocks that check `err instanceof OpenXmlError`
  now match these errors uniformly with the rest of the library.

## 0.7.0

### Minor Changes

- [#74](https://github.com/office-kit/xlsx/pull/74) [`fb57bae`](https://github.com/office-kit/xlsx/commit/fb57baef237b13c3873a541ff6662ef6040d84bb) Thanks [@baseballyama](https://github.com/baseballyama)! - Harden the reader against decompression-bomb attacks and tighten release
  hygiene ahead of `1.0`:

  - `loadWorkbook` / `loadWorkbookStream` now apply a `decompressionLimits`
    guard by default (per-entry size cap, total archive cap, compression-ratio
    cap). The new `OpenXmlDecompressionBombError` (a subclass of
    `OpenXmlIoError`) is thrown when an archive trips the limit. Pass
    `decompressionLimits: false` to disable, or supply a partial override to
    tighten or loosen specific bounds.
  - `saveWorkbook` / `workbookToBytes` now validate sheet titles against
    Excel's rules (1–31 chars, forbidden `: \ / ? * [ ]`, no leading/trailing
    apostrophe, reserved `History`, case-insensitive uniqueness) at save time,
    catching invalid names that were introduced by direct mutation of
    `ws.title` after `addWorksheet`.
  - `size-limit` now tracks the minified parse size (no brotli, no gzip) of
    `@office-kit/xlsx/streaming` and `@office-kit/xlsx/io` in addition to the existing
    min+brotli budgets, so transitive bundle growth (e.g. the stylesheet
    writer chunk) is caught at PR time.
  - New `SECURITY.md` documents the supported versions, the private security
    advisory reporting process, and `decompressionLimits` recommendations for
    consumers.
  - New `CONTRIBUTING.md`, GitHub Issue / PR templates, a
    `template-compliance` workflow, and a project-specific `CLAUDE.md` for
    contributors and AI agents working in the repository.
  - `docs/migrate-from-openpyxl.md` realigned to the 0.6.x API surface
    (`iterRows`, `setCellByCoord`, `addWorksheet` returning an empty workbook,
    ZIP64 entry-count support, the current passthrough part list).

## 0.6.0

### Minor Changes

- [#71](https://github.com/office-kit/xlsx/pull/71) [`551901c`](https://github.com/office-kit/xlsx/commit/551901cbbd22952e8b26d5421e9371df08721130) Thanks [@baseballyama](https://github.com/baseballyama)! - Re-export DML colour, fill, and text-body primitives from `@office-kit/xlsx/drawing`. Chart styling reaches `<a:srgbClr>` / `<a:solidFill>` / `<a:bodyPr>…<a:p>…<a:r>` through `ShapeProperties.fill`, `Series.spPr`, `Axis.txPr`, etc., but the building blocks (`DmlColor`, `DmlColorWithMods`, `Fill`, `TextBody`, `TextParagraph`, `RunProperties`, …) and their constructors (`makeColor`, `makeSrgbColor`, `makeSchemeColor`, `makeSolidFill`, `makeTextBody`, `makeParagraph`, `makeRun`, `makeRunProperties`, …) previously had no public home. They now ship as part of `@office-kit/xlsx/drawing` alongside `makeShapeProperties`. Closes [#55](https://github.com/office-kit/xlsx/issues/55), closes [#56](https://github.com/office-kit/xlsx/issues/56).

- [#72](https://github.com/office-kit/xlsx/pull/72) [`80a06bf`](https://github.com/office-kit/xlsx/commit/80a06bff737d8034ed0e5de89686c0a3f6d953d3) Thanks [@baseballyama](https://github.com/baseballyama)! - `AxisShared.majorGridlines` and `AxisShared.minorGridlines` now accept `boolean | Gridlines` instead of just `boolean`. The `Gridlines` shape carries a `ShapeProperties`, so `<c:majorGridlines><c:spPr><a:ln>…</a:ln></c:spPr></c:majorGridlines>` can be emitted to colour / dash / weight the gridline (e.g. corporate-style light grey `D9D9D9`). The plain `true` form keeps emitting `<c:majorGridlines/>` so all existing call sites stay unchanged. Round-trip through `parseChartXml` is preserved for both forms. Closes [#57](https://github.com/office-kit/xlsx/issues/57).

### Patch Changes

- [#69](https://github.com/office-kit/xlsx/pull/69) [`2317545`](https://github.com/office-kit/xlsx/commit/2317545ae784c772bc65f088a0c0fb9063904c35) Thanks [@baseballyama](https://github.com/baseballyama)! - Rename the chart-internal `NumberFormat` interface to `ChartNumberFormat` and re-export it from `@office-kit/xlsx/chart`. The interface was already part of the public surface through `AxisShared.numFmt` and `DataLabelList.numFmt`, but the type itself was not exported — callers building axis / data-label options had to write the literal inline. The new name also disambiguates from the cell-stylesheet `NumberFormat` exported from `@office-kit/xlsx/styles`, which is a different shape (`{ numFmtId, formatCode }`). Closes [#58](https://github.com/office-kit/xlsx/issues/58).

- [#68](https://github.com/office-kit/xlsx/pull/68) [`060f436`](https://github.com/office-kit/xlsx/commit/060f436b46e79cd6d3ecf9613ce3a04278bb641c) Thanks [@baseballyama](https://github.com/baseballyama)! - Harden DrawingML `Fill` serializer against two natural mis-uses. (1) Passing a colour without `mods` (e.g. `{ base: { kind: 'srgb', value: 'FF0000' } }` instead of `{ base, mods: [] }`) no longer crashes the chart serializer with `Cannot read properties of undefined (reading 'map')`; the missing modifier list is now treated as empty. (2) Passing a `Fill` with an unknown `kind` (e.g. `'solid'` instead of `'solidFill'`) used to silently emit an empty `<c:spPr></c:spPr>` and lose the caller's styling intent; the serializer now throws `OpenXmlSchemaError` so the mistake surfaces immediately.

- [#73](https://github.com/office-kit/xlsx/pull/73) [`f376354`](https://github.com/office-kit/xlsx/commit/f376354bf305db80a615a2a5ba9597bd276e167a) Thanks [@baseballyama](https://github.com/baseballyama)! - Document the small set of openpyxl → @office-kit/xlsx defaults that differ, in particular that `createWorkbook()` returns an empty workbook with no sheets (unlike `openpyxl.Workbook()` which creates a default `'Sheet'`). Direct ports of openpyxl code that include a `wb.remove(wb.active)` call after `Workbook()` were translating that into a no-op `removeSheet(wb, 'Sheet')` — the new README "Migrating from openpyxl" subsection calls this out alongside the `setCell` and `makeBorder` / `makeSide` equivalents. Closes [#62](https://github.com/office-kit/xlsx/issues/62).

## 0.5.0

### Minor Changes

- [#66](https://github.com/office-kit/xlsx/pull/66) [`afecdc3`](https://github.com/office-kit/xlsx/commit/afecdc3b0b822f4d3ab3ecd16458c7a76a847f3e) Thanks [@baseballyama](https://github.com/baseballyama)! - Remove the silently-ignored `readOnly` / `keepLinks` / `keepVba` / `dataOnly` / `richText` placeholders from `LoadOptions`. They were declared on the public surface but the loader (`src/io/load.ts`) accepted them via `_opts` and dropped them on the floor, so production callers expecting `dataOnly: true` to suppress formulas — or `readOnly: true` to enable a special path — got the default behaviour instead. `LoadOptions` is now an empty type until the underlying behaviour ships; future toggles will land here once they actually do something. The `loadWorkbook(source, opts)` signature is unchanged.

### Patch Changes

- [#64](https://github.com/office-kit/xlsx/pull/64) [`b613607`](https://github.com/office-kit/xlsx/commit/b61360774e4f3b1423985ee5cf924093a991e32d) Thanks [@baseballyama](https://github.com/baseballyama)! - Treat sheet names as case-insensitive for uniqueness, matching Excel. Previously `addWorksheet(wb, 'Data')` followed by `addWorksheet(wb, 'data')` succeeded locally but produced a workbook Excel and LibreOffice refuse to open. `addWorksheet`, `addChartsheet`, `duplicateSheet`, `renameSheet`, and `pickUniqueSheetTitle` now compare titles case-insensitively. A case-only rename of the same sheet (`renameSheet(wb, 'Data', 'data')`) is allowed.

- [#51](https://github.com/office-kit/xlsx/pull/51) [`1cf8d0c`](https://github.com/office-kit/xlsx/commit/1cf8d0cc1f896366608c5d60aa40b4efc682bed9) Thanks [@baseballyama](https://github.com/baseballyama)! - Tighten the streaming I/O surface so the README's "fixed-memory" claims hold up
  in practice.

  - `toFile().toBytes().finish()` no longer re-reads the just-written file. The
    previous code called `fs.readFile(path)` from `finish()` and returned the
    full archive bytes, defeating the chunk-streamed write — a 10M-row workbook
    ended its save by reloading the entire output into memory. `finish()` now
    resolves with an empty `Uint8Array` once the underlying write stream has
    flushed; callers that need the bytes should `fs.readFile()` the path
    themselves.
  - `toFile` and `toWritable` honour write-stream backpressure: when `write()`
    returns `false`, subsequent chunks chain off a `drain` event before
    proceeding, so peak memory tracks the writable's `highWaterMark` rather
    than the producer's pace.
  - `workbookToBytes` no longer depends on `Buffer`. Browser bundles that omit
    the Node `Buffer` polyfill previously broke at `toBuffer().result()`
    because the in-memory sink ended its result with `Buffer.from(...)`. The
    helper now uses a `Uint8Array`-only sink; a regression test
    (`tests/phase-1/io/browser.test.ts`) saves a workbook with `globalThis.Buffer`
    shadowed to `undefined`.
  - The streaming read path inflates worksheet entries chunk-by-chunk. A new
    `ZipArchive.readStream(path)` returns a `ReadableStream<Uint8Array>` that
    drives fflate's `Inflate` incrementally, and `loadWorkbookStream`'s
    whole-sheet `iterRows()` feeds the SAX parser directly off that stream so
    the inflated worksheet body is never fully resident. Band queries
    (`minRow > 1`) still materialise the inflated sheet to build the row-offset
    index — that trade-off is unchanged.
  - Documentation: the `XlsxSink` / `BufferedSinkWriter` JSDoc no longer
    describes `toBytes()` as the "buffered mode" — that name was historical;
    the underlying object can either accumulate (buffered sinks) or forward
    chunks as they arrive (streaming sinks). The README also clarifies that
    the streaming reader still loads the compressed archive up front (ZIP
    needs random access to the central directory) — the win is that the
    inflated worksheet payload is never fully resident.

- [#63](https://github.com/office-kit/xlsx/pull/63) [`fa73fc5`](https://github.com/office-kit/xlsx/commit/fa73fc5a7e4c03a69eecc36acfb3a3f6482e525d) Thanks [@baseballyama](https://github.com/baseballyama)! - Tighten sheet-title validation on the streaming write path. The `createWriteOnlyWorkbook` `addWorksheet` call now applies the same rules as the buffered `addWorksheet` (no `: \ / ? * [ ]`, no leading / trailing apostrophe, not the reserved name `History`) and rejects duplicate titles case-insensitively so the streaming path can't produce a workbook Excel refuses to open.

## 0.4.0

### Minor Changes

- [`f9f273d`](https://github.com/office-kit/xlsx/commit/f9f273d390b034ac58a9d545b12ef650fa9a583a) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose the full ECMA-376 axis attribute surface on `CategoryAxis` and
  `ValueAxis`. Previously the serializer emitted fixed defaults for several
  elements; these are now driven by typed fields, unblocking horizontal-bar
  reversal (`scaling.orientation: 'maxMin'`), 100 %-stacked axis caps
  (`scaling.max`), value-axis crossing rules, custom tick formatting, axis
  titles, and more.

  Newly exposed shared fields: `scaling` (`orientation`/`min`/`max`/`logBase`),
  `crosses`, `crossesAt`, `numFmt`, `majorTickMark`, `minorTickMark`,
  `tickLblPos`, `title`, `minorGridlines`. `ValueAxis` gains `crossBetween`,
  `majorUnit`, `minorUnit`. `CategoryAxis` gains `auto`, `lblAlgn`,
  `lblOffset`, `noMultiLvlLbl`. All previously-emitted defaults remain the
  output when fields are unset, so existing files are unchanged.

  New type exports: `AxisCrossBetween`, `AxisCrosses`, `AxisOrientation`,
  `AxisScaling`, `CategoryLabelAlignment`, `TickLabelPosition`, `TickMark`.

  Closes [#46](https://github.com/office-kit/xlsx/issues/46).

- [`2e5e460`](https://github.com/office-kit/xlsx/commit/2e5e4606f098ba9822bc4aaef76db324b51eeea9) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose `overlap?: number` on `BarChart` (and `makeBarChart`). The serializer
  now emits `<c:overlap val="N"/>` (range -100..100) inside `<c:barChart>` when
  set, unblocking flush stacking (`overlap: 100`) and negative-space clustered
  bars. When unset, the serializer continues to emit the prior default of
  `<c:overlap val="100"/>` for `stacked` / `percentStacked` grouping so existing
  output is unchanged.

  Closes [#45](https://github.com/office-kit/xlsx/issues/45).

- [`19e8368`](https://github.com/office-kit/xlsx/commit/19e8368582c70ad89e0d6ec0265e8a7cb756ded1) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose `style?: number` on `ChartSpace` (and `makeChartSpace`). The serializer
  emits `<c:style val="N"/>` (range 1..48) between `<c:roundedCorners>` and
  `<c:chart>`, selecting one of Excel's built-in "Chart Styles" gallery presets
  — the same single attribute openpyxl writes via `chart.style = N`.

  Closes [#48](https://github.com/office-kit/xlsx/issues/48).

- [`1541291`](https://github.com/office-kit/xlsx/commit/154129136198a2f22beef0a9796f2a2ba16fcaac) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `DateAxis` and `SeriesAxis` types and `dateAx?` / `serAx?` slots on
  `PlotArea`. `DateAxis` carries `auto`, `lblOffset`, `baseTimeUnit`,
  `majorUnit`, `majorTimeUnit`, `minorUnit`, `minorTimeUnit` on top of the
  shared axis surface — unblocking time-series charts (`<c:dateAx>`).
  `SeriesAxis` adds `tickLblSkip` and `tickMarkSkip`, used by surface charts
  (`<c:serAx>`). The serializer emits both inside `<c:plotArea>` between the
  inferred cat/val axes and `<c:spPr>`; the parser round-trips them.

  New type exports: `DateAxis`, `SeriesAxis`, `TimeUnit`.

- [`0708aa8`](https://github.com/office-kit/xlsx/commit/0708aa81106db5f4276d1d214fc5714e25996fb3) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `Layout` / `ManualLayout` types and expose `layout?: Layout` on
  `ChartTitle`, `PlotArea`, and `Legend`. The serializer emits
  `<c:layout><c:manualLayout>` with `layoutTarget`, `xMode` / `yMode` /
  `wMode` / `hMode`, and `x` / `y` / `w` / `h` when set, falling back to the
  existing empty `<c:layout/>` placeholder when unset — so output is unchanged
  for charts that don't configure manual layout. Parser round-trips both
  forms.

  New type exports: `Layout`, `LayoutMode`, `LayoutTarget`, `ManualLayout`.

- [`0989eec`](https://github.com/office-kit/xlsx/commit/0989eec45f879d05a7707da8402fd734f4a3208b) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose per-point `dPt?: DataPoint[]` on `BarSeries` (used by bar / line /
  area / pie / doughnut / radar / stock / surface), `ScatterSeries`, and
  `BubbleSeries`, with the new `DataPoint` type carrying `idx`,
  `invertIfNegative?`, `marker?`, `bubble3D?`, `explosion?`, and `spPr?`.
  The serializer emits `<c:dPt>` children between the series'
  `<c:marker>`/`<c:spPr>` and `<c:dLbls>` per ECMA-376 sequence — unblocking
  per-slice colours on pie / doughnut charts, per-bar colours on single-series
  bar charts, and per-point styling on line / scatter / bubble.

  Closes [#44](https://github.com/office-kit/xlsx/issues/44).

- [`7f9e143`](https://github.com/office-kit/xlsx/commit/7f9e1430c32ad685b14e382c1a17abac41f24b4f) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `invertIfNegative?: boolean` and `explosion?: number` to `BarSeries`
  (used by bar / line / area / pie / doughnut / radar / stock / surface) and
  `invertIfNegative?: boolean` to `BubbleSeries`. The serializer emits
  `<c:invertIfNegative>` and `<c:explosion>` between `<c:spPr>` and `<c:dPt>`
  per ECMA-376 sequence — unblocking per-series colour inversion on negative
  values and pie/doughnut slice explosion at the series level (in addition to
  the per-point `DataPoint.explosion`).

- [`ffa777c`](https://github.com/office-kit/xlsx/commit/ffa777caaddbf00f2cffc7fedce8d021a2a584f6) Thanks [@baseballyama](https://github.com/baseballyama)! - Expose `marker?: Marker` on `LineSeries` and `ScatterSeries` (with the new
  `Marker` / `MarkerSymbol` types). The serializer emits `<c:marker>` between
  the series' `<c:spPr>` and `<c:dLbls>` per ECMA-376 sequence, carrying
  `<c:symbol>`, `<c:size>`, and an optional nested `<c:spPr>` for marker
  fill / line colour — matching openpyxl's `series.marker = Marker(...)`.

  Closes [#47](https://github.com/office-kit/xlsx/issues/47).

- [`70a2f17`](https://github.com/office-kit/xlsx/commit/70a2f17bfe6a1106df04702de5e11e2ac16cd596) Thanks [@baseballyama](https://github.com/baseballyama)! - Extend `StockChart.hiLowLines` and `StockChart.upDownBars` to accept a
  detailed object form in addition to the existing boolean flag. The
  detailed form lets callers style the lines (`HiLowLines.spPr`) and the
  up/down bars (`UpDownBars.gapWidth` + `upBars.spPr` + `downBars.spPr`)
  with per-element shape properties.

  The boolean form (`hiLowLines: true`) keeps its existing meaning and
  output, so existing callers are unaffected. Parser round-trips both
  forms, picking the boolean form when no detail is found.

  New type exports: `BarFrame`, `HiLowLines`, `UpDownBars`.

- [`7cd181c`](https://github.com/office-kit/xlsx/commit/7cd181c9276dfe8c37675d3c4b1c77e020b82b64) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `view3D?: View3D` and `floor?` / `sideWall?` / `backWall?` (typed
  `SurfaceFrame`) to `ChartSpace` (and `makeChartSpace`). The serializer
  emits `<c:view3D>` (with `rotX`, `rotY`, `depthPercent`, `hPercent`,
  `rAngAx`, `perspective`) and `<c:floor>` / `<c:sideWall>` / `<c:backWall>`
  (with `thickness` and `spPr`) between `<c:autoTitleDeleted>` and
  `<c:plotArea>` per ECMA-376 sequence — unblocking real 3-D chart viewpoints
  and wall styling for `bar3DChart` / `line3DChart` / `pie3DChart` /
  `area3DChart` / `surface3DChart`.

## 0.3.1

### Patch Changes

- [#41](https://github.com/office-kit/xlsx/pull/41) [`a04f645`](https://github.com/office-kit/xlsx/commit/a04f6459113810c638adf4247d1a201e34123d1c) Thanks [@baseballyama](https://github.com/baseballyama)! - Relax `engines.node` from `>=24.15.0` back to `>=22.0.0` so the published
  package installs on every active Node LTS line (22.x, 24.x) plus current
  (26.x), matching the CI matrix. 0.3.0 inadvertently shipped a Node 24+
  floor that excluded the still-supported 22.x LTS; this restores broader
  LTS coverage. The library does not rely on any Node 24-only API.

## 0.3.0

### Minor Changes

- [#32](https://github.com/office-kit/xlsx/pull/32) [`87a0051`](https://github.com/office-kit/xlsx/commit/87a005104f4d54b5fd0a1a747acc515d6cf9171e) Thanks [@baseballyama](https://github.com/baseballyama)! - **Breaking**: `iterRows` / `iterValues` (in `@office-kit/xlsx/worksheet`) now
  iterate **rectangularly** over the populated bounding box rather than
  skipping empty rows and gaps. `iterRows` yields
  `(Cell | undefined)[]` (one entry per `[minCol, maxCol]` position);
  `iterValues` yields `CellValue[]` with `null` filling the gaps.

  Default extent switches from `MAX_ROW`/`MAX_COL` (the 1M × 16K sheet
  limit) to `getMaxRow(ws)` / `getMaxCol(ws)` (the populated bounding
  box). The `IterRowsOptions.valuesOnly` flag is removed — it was already
  unread.

  Migration:

  - Aggregation callers that want populated rows only:
    `[...iterRows(ws)].filter((row) => row.some((c) => c !== undefined))`.
  - Cell-by-cell streaming over populated cells only: keep using `iterCells`
    (unchanged).

  Closes [#24](https://github.com/office-kit/xlsx/issues/24).

- [#31](https://github.com/office-kit/xlsx/pull/31) [`9297c46`](https://github.com/office-kit/xlsx/commit/9297c46aed63f566b081db97fc1f84ca9a24b3c0) Thanks [@baseballyama](https://github.com/baseballyama)! - Extend `cellValueAsString` (in `@office-kit/xlsx/cell`) with optional
  `dateFormat` / `emptyText` overrides and add a sibling
  `cellValueAsPrimitive` that maps a `CellValue` to the most natural JS
  primitive (`string | number | boolean | Date | null`) without forcing a
  single target type. Closes [#25](https://github.com/office-kit/xlsx/issues/25).

- [#30](https://github.com/office-kit/xlsx/pull/30) [`78d04fd`](https://github.com/office-kit/xlsx/commit/78d04fd8cee4f9b2b0991e14f471ede39adbef2f) Thanks [@baseballyama](https://github.com/baseballyama)! - Add `workbookToBuffer` to `@office-kit/xlsx/node`. One-shot Node-flavored helper
  that returns a `Buffer` directly, paralleling the existing `fromBuffer`
  source. Closes [#28](https://github.com/office-kit/xlsx/issues/28).

## 0.2.0

### Minor Changes

- [`b36ca45`](https://github.com/office-kit/xlsx/commit/b36ca453b08c91981baac42b3b5bc4aeeeef6ec0) Thanks [@baseballyama](https://github.com/baseballyama)! - Hardening and docs release.

  - Add a 3-tier ECMA-376 conformance validator and broaden conformance coverage to the writer surface, real-world fixtures, and fast-check property tests.
  - Add `knip` to CI to keep the public export surface tight; prune unused exports flagged by it.
  - Refresh the docs site: redesigned landing and docs UI with a new typography system, new logo and favicons, and new "Why @office-kit/xlsx" / comparison / motivation sections in the README.
  - Tighten release / dependency automation: pin dependencies, drop EOL Node 18/20 from the test matrix and add Node 26, bump the project Node engine to 22.22.2.
