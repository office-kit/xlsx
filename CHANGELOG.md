# @office-kit/xlsx

## 0.14.0

### Minor Changes

- [#147](https://github.com/office-kit/xlsx/pull/147) [`5074ac5`](https://github.com/office-kit/xlsx/commit/5074ac582024ee339e83a44c5bbf7f0594c9c3c2) Thanks [@kibertoad](https://github.com/kibertoad)! - fix: reading an entry twice no longer counts it twice against the decompression budget

  `readStream()` followed by `read()` on the same archive entry charged its payload
  to `maxTotalUncompressedBytes` twice, so a legitimate workbook could be rejected
  as a decompression bomb. Each entry is now charged the largest amount any single
  inflate of it produced, so reading one again costs nothing however the reads
  overlap. The per-entry size and ratio caps are unchanged, and an archive whose
  distinct entries genuinely exceed the total is still rejected.

  An entry rejected for exceeding the archive total remains rejected on repeated
  sync or streaming reads, including archives with understated directory sizes.

  perf!: loading a workbook no longer holds every part it has read

  Inflated entries were cached for the lifetime of the archive, which put the whole
  uncompressed package in memory. The cache now takes entries of 64 KB or less, 4 MB
  of them in total, and drops the least recently used first, which still covers the
  `.rels` parts a load re-reads. Loading a 10-sheet, 8.3 MB workbook holds about
  21 MB of inflated bytes instead of about 65 MB.

  `openZip().read(path)` returns a fresh array on every call as part of this. It
  previously handed back the same array once an entry had been read, so mutating one
  read's result changed what later reads of that path returned. Code that relied on
  that aliasing, or on two reads yielding the identical object, has to keep its own
  reference now.

  Repeated row-band queries retain their worksheet bytes and row index without
  re-inflating the part. Closing the workbook releases these caches and prevents
  subsequent band queries from using stale bytes.

- [#145](https://github.com/office-kit/xlsx/pull/145) [`45d6ad1`](https://github.com/office-kit/xlsx/commit/45d6ad1e03fd775e133a8ca6e4e1ac06b0e1ee7e) Thanks [@kibertoad](https://github.com/kibertoad)! - perf: saveWorkbook no longer holds whole worksheets in memory

  Worksheets are serialised straight into their ZIP entry instead of being built
  as a string, encoded, and held until every sheet is done. Saving a single
  500k-cell sheet now completes under an 88 MB heap cap where it previously needed
  192 MB; an eight-sheet, 2M-cell workbook needs 320 MB instead of 384 MB.

  **Saved bytes change.** `xl/_rels/workbook.xml.rels` is now written after the
  worksheet parts rather than before them, because its contents depend on whether
  serialising the sheets produced any shared strings. Code that reads the package
  through `loadWorkbook`, a zip library, or Excel is unaffected, since OPC resolves
  parts by name and archive position carries no meaning. Byte-for-byte consumers
  are: if you pinned `mtime` for reproducible output (added in 0.13.0) and store
  golden files or content hashes of saved workbooks, every one of them changes
  with this release and has to be regenerated.

- [#144](https://github.com/office-kit/xlsx/pull/144) [`ec45506`](https://github.com/office-kit/xlsx/commit/ec455065629df01aa84efd830c3f2b19657d4747) Thanks [@kibertoad](https://github.com/kibertoad)! - perf: streamed sheet reads no longer buffer a whole chunk's parse events

  Walking a worksheet queued every SAX event a single parser write produced
  before yielding the first row, so peak heap tracked whatever the parser was
  handed rather than staying flat. How much that was depended on the shape of
  the input: a materialised 40 MB sheet body went in whole, a streamed one at
  the zip reader's inflate granularity of roughly 450 KB. Both are now decoded
  and fed in 64 KB slices, whatever the producer hands over, and the event queue
  drains between writes.

  On a 200k-row, 5-column sheet, peak heap for `iterRows({ minRow: 2 })` drops
  from about 1240 MB to about 70 MB, and for a full-sheet walk from about
  200 MB to about 60 MB.

  Three changes in behavior come with it:

  - `iterRows({ minRow })` above row 1 no longer fails on sheets written by
    Excel. The band was replayed inside a rebuilt `<sheetData>` envelope that
    declared only the default namespace, so the `x14ac:dyDescent` Excel puts on
    nearly every `<row>` raised an unbound-prefix error. The replay now carries
    the worksheet's and sheetData's own namespace declarations.
  - Malformed XML throws `OpenXmlSchemaError` with the parser's own error as
    `cause`, instead of surfacing the raw `saxes` `Error`.
  - Streamed input is scanned for DTD and entity declarations across the whole
    document instead of only its first 256 characters. A `<!DOCTYPE` or
    `<!ENTITY` token further into the payload, inside a comment or a CDATA
    section for instance, is now rejected where it previously parsed.

### Patch Changes

- [#145](https://github.com/office-kit/xlsx/pull/145) [`45d6ad1`](https://github.com/office-kit/xlsx/commit/45d6ad1e03fd775e133a8ca6e4e1ac06b0e1ee7e) Thanks [@kibertoad](https://github.com/kibertoad)! - fix: a failed `saveWorkbook` to a file path no longer leaves the partial file behind

  `toFile` deletes the half-written file when a save fails, but the delete raced
  the write stream's own file-handle close and lost on Windows, so a truncated
  `.xlsx` stayed at the destination. The cleanup now waits for the handle to
  close, and `saveWorkbook` (and the write-only `finalize()`) wait for the cleanup
  before rejecting, so the path is clear by the time the error reaches the caller.

## 0.13.0

### Minor Changes

- [#137](https://github.com/office-kit/xlsx/pull/137) [`6d29a1a`](https://github.com/office-kit/xlsx/commit/6d29a1a9cdc0ae921899cb982c67c76745a704b1) Thanks [@kibertoad](https://github.com/kibertoad)! - feat!: `SaveOptions.mtime` for byte-identical output, and `compressionLevel` is now wired through

  ZIP has no "no timestamp" encoding, so fflate stamped the wall clock into every
  entry's local header and central-directory record. Two saves of the same workbook
  therefore differed in bytes, which ruled out golden-file tests and
  content-addressed caching for anything this library writes.

  `saveWorkbook` / `workbookToBytes` now accept `mtime?: Date`, applied to every
  entry. `createWriteOnlyWorkbook` takes the same option. With `mtime` pinned and the
  core properties set from the caller's own data rather than the clock, identical
  input produces identical bytes.

  The stamp records the date's UTC wall time, so the archive is the same on every
  machine. ZIP's DOS date field carries no timezone, and fflate reads a `Date`
  through local-time getters, which would have made the bytes depend on the writer's
  `TZ`: a golden file committed from a laptop would not match the one CI renders
  from the same input. A date whose year falls outside 1980-2099, the range supported by
  the ZIP backend, is rejected with an `OpenXmlIoError` before anything is written
  rather than part-way through the first entry. Resolution is two seconds, per the
  format.

  `SaveOptions.compressionLevel` was declared but documented as "Reserved" and never
  read; it now reaches fflate's deflate stream on both the buffered and the streaming
  writer, and is typed `0 | 1 | ... | 9` rather than `number`. Two things change for
  callers who already set it: the type narrows, so `{ compressionLevel: someNumber }`
  no longer typechecks, and the value is now honoured, so output that silently came
  back at fflate's default level 6 changes in size and in bytes. A level outside
  `0..9` throws instead of falling back to fflate's default, which is what the old
  "Reserved" option effectively did. `CompressionLevel` is exported from
  `@office-kit/xlsx/zip`, alongside `ZipWriterOptions`.

  UTC timestamps remain identical across daylight-saving transitions, including
  local times that do not exist in the writer’s timezone.

## 0.12.0

### Minor Changes

- [#139](https://github.com/office-kit/xlsx/pull/139) [`9b3c46c`](https://github.com/office-kit/xlsx/commit/9b3c46cb654369b340deec6682f91a53840fc6e1) Thanks [@kibertoad](https://github.com/kibertoad)! - feat!: `ensureCell` for get-or-create, and `setCell`'s `value` is now mandatory

  `setCell(ws, row, col)` read like "reach the cell at (row, col)" and wrote `null`. A
  styling pass that walked already-populated rows therefore erased the values and
  formulas it touched, and nothing in the signature or the docstring said so. The
  cheatsheet and the formula recipe both taught the no-value form as the way to reach a
  cell, so the trap was the documented path.

  `value` is now required on `setCell` and `setCellByCoord`, which turns the mistake
  into a compile error rather than a wrong file. `ensureCell(ws, row, col)` returns the
  cell at a coordinate and allocates a blank one only when it does not exist, leaving
  an existing value untouched; `ensureCellByCoord(ws, 'B5')` is the A1-addressed form.
  That pattern already existed inside `applyToRange`, `setRangeStyle`,
  `setRangeWrapText`, `setRangeAlignment` and `setRangeBorderBox`; those now call
  `ensureCell` instead of open-coding it.

  Migration is mechanical, and the compiler points at every site:

  - `setCell(ws, r, c)` becomes `ensureCell(ws, r, c)`, and `setCellByCoord(ws, 'B5')`
    becomes `ensureCellByCoord(ws, 'B5')`
  - emptying a cell stays available and is now explicit. `null` is a `CellValue`,
    so `setCell(ws, r, c, null)` clears the value and leaves the cell in the sheet
    with its fill, border and number format, the way Excel's Delete key does.
    `deleteCell` drops the cell outright and `clearRange` does the same across a
    rectangle.

  `mergeCells` drops the cells underneath a merge, and reaching one of those
  coordinates with `ensureCell` allocates it again, so the written `<sheetData>` carries
  a blank `<c>` under the merge. Address the top-left coordinate when the merged block
  is what you mean.

  Also removed: `setCellFormula`, `setCellArrayFormula` and `setCellRichText` in
  `src/worksheet/worksheet.ts`. They were dropped from the public subpaths in an earlier
  "one way per task" trim and have been unreachable since. `ensureCell` plus
  `setFormula` / `setArrayFormula` covers the two formula wrappers. Rich text is
  `setCell(ws, r, c, { kind: 'rich-text', runs: makeRichText(runs) })`: `makeRichText`
  builds the runs and the `kind` wrapper is spelled out at the call site, where
  `makeErrorValue` / `makeDurationValue` hand back a `CellValue` outright.

- [#136](https://github.com/office-kit/xlsx/pull/136) [`ef50829`](https://github.com/office-kit/xlsx/commit/ef50829e8d21a379d1fb8cc062d718349c74a6b0) Thanks [@kibertoad](https://github.com/kibertoad)! - fix: formula text starting with `=` produced a workbook Excel calls damaged

  OOXML stores formula text without the leading `=` (ECMA-376 §18.3.1.40), but no
  setter stripped it. `setFormula(cell, '=SUM(A1:A3)')` emitted
  `<f>=SUM(A1:A3)</f>`, which Excel reports as an unreadable-content error on open.
  The natural spelling was the broken one, and nothing in the types or the docs said
  so.

  Formula text is now normalised (leading `=`, plus the whitespace around it) both
  where it enters the model and where it is serialised, so a hand-built
  `FormulaValue` passed straight to `setCell` cannot produce a damaged file either.
  The other elements that carry OOXML formula text get the same treatment:
  `<formula1>` / `<formula2>` on a data validation, `<formula>` on a
  conditional-formatting rule, and a defined name's value.

  Two behaviour changes to watch for on upgrade:

  - Loading a workbook whose `<f>` carried a leading `=` (including any file this
    library wrote before this release) now gives `getFormulaText(cell)` as
    `'SUM(A1:A3)'` where it returned `'=SUM(A1:A3)'`, and re-saving writes the
    normalised text.
  - `makeDataValidation`, `makeCfRule`, `addDefinedName` and the builders over them
    (`addListValidation`, `addCustomValidation`, `addFormulaRule`, …) store the
    normalised text, so reading `dv.formula1` back returns it without the `=`.

  `makeFormula` and `makeArrayFormula` now throw `OpenXmlSchemaError` when the text
  normalises to nothing (`''` or `'='`). That case used to emit an empty `<f/>`,
  which Excel rejects as well, and loading a file that already contains one now
  throws rather than carrying a formula cell with no expression.

- [#136](https://github.com/office-kit/xlsx/pull/136) [`ef50829`](https://github.com/office-kit/xlsx/commit/ef50829e8d21a379d1fb8cc062d718349c74a6b0) Thanks [@kibertoad](https://github.com/kibertoad)! - feat: formula value constructors, and the `<calcPr>` setters are now exported

  Placing a formula took two steps: reach or create a cell, then mutate it with
  `setFormula`. `makeFormula(text, { cachedValue })` returns the `CellValue`, so
  `setCell(ws, row, col, makeFormula('SUM(B5:I5)'), styleId)` is the whole write.
  `makeArrayFormula`, `makeSharedFormula` and `makeDataTableFormula` do the same
  for the other `<f>` kinds, and `setFormula` / `setArrayFormula` /
  `setSharedFormula` / `setDataTableFormula` stay as the form that applies the same
  value to a cell you already hold.

  Five helpers over the workbook's `<calcPr>` existed but none of them was
  reachable. `setCalcMode`, `setIterativeCalc`, `setCalcOnSave`, `setFullCalcOnLoad`
  and `setFullPrecision` are now exported from `@office-kit/xlsx/workbook`.

  `setFullCalcOnLoad(wb, true)` asks a calculating app to recompute the workbook on
  open instead of trusting the cached values in the file: reach for it when you
  wrote formulas this library cannot evaluate for you, or when the values you did
  cache may be stale. It does nothing for viewers that never calculate (Quick Look,
  Outlook and SharePoint previews, most thumbnailers), which show a `cachedValue`
  or an empty cell, so keep supplying one wherever the producer can compute it.

- [#141](https://github.com/office-kit/xlsx/pull/141) [`d26cd23`](https://github.com/office-kit/xlsx/commit/d26cd23d231cbf2e1523f77d52ea553debadc2c7) Thanks [@kibertoad](https://github.com/kibertoad)! - feat!: numeric coordinates wherever an A1 string was required, and one way to freeze panes

  Every range-taking helper insisted on an A1 string, so code that tracks rows and
  columns as integers had to format `"A4:H20"` for the callee to parse straight back
  into the numbers it started with. These now take `string | { minRow, minCol, maxRow,
maxCol }`, a union named `RangeRef` in `@office-kit/xlsx/utils`:

  - `@office-kit/xlsx/styles`: `setRangeStyle`, `setRangeFont`, `setRangeBackgroundColor`,
    `setRangeNumberFormat`, `setRangeAlignment`, `setRangeWrapText`, `setRangeProtection`,
    `setRangeBorderBox`, `formatAsHeader`, `clearRangeStyle`
  - `@office-kit/xlsx/worksheet`: `setRangeValues`, `getRangeValues`, `applyToRange`,
    `clearRange`, `getCellsInRange`, `replaceInRange`, `getRangeAddress`, `copyRange`,
    `moveRange`, `mergeCells`, `unmergeCells`

  `writeRange` takes a `{ row, col }` anchor alongside the A1 form. Bounds are validated
  against the sheet grid and inverted bounds are normalised, so `{ minRow: 5, maxRow: 1 }`
  covers the same cells as `"A5:A1"` instead of iterating nothing, and a fractional or
  off-grid bound throws before any cell is touched. `mergeCells` stores a normalised copy,
  so mutating a bounds object after the call no longer rewrites a merge that is already on
  the sheet.

  `setSelectedRange` keeps its string parameter, because an `sqref` can hold several
  ranges, and the `*Str` helpers (`shiftRangeStr`, `rangeAreaStr`, `expandRangeStr` and
  friends) stay string-in / string-out by definition.

  This widens one parameter rather than adding a second function, so each capability still
  has a single canonical helper that reads either spelling, the way `setFreezePanes` now
  reads either.

  **Breaking:** `setRangeValues` clips to the range it was given. Values past the bottom or
  right edge of `range` are dropped instead of written outside it, which is what
  `copyRange` already does against a smaller target, and it makes `getRangeValues` a true
  inverse. `setRangeValues(ws, 'A1', rows)` used to lay down a whole block from a one-cell
  range; `writeRange(ws, 'A1', rows)` is that behaviour, and it returns the bounding box it
  wrote.

  Formula text had the same string-concatenation problem. `tupleToCoordinate` and
  `boundariesToRangeString` gained `absoluteCol` / `absoluteRow`, so `$B$5` and
  `$A$4:$H$20` come out of the helpers.

  **Breaking:** `freezePanes(ws, rows, cols)` is removed. `setFreezePanes` now accepts
  `'B2' | { rows, cols } | undefined`, which covers both spellings through one function.
  The numeric form is also strictly more capable: `freezePanes` required both counts to be
  at least 1, so "freeze two rows and no columns" could not be expressed. Replace
  `freezePanes(ws, 1, 1)` with `setFreezePanes(ws, { rows: 1, cols: 1 })`.

- [#140](https://github.com/office-kit/xlsx/pull/140) [`3af5f73`](https://github.com/office-kit/xlsx/commit/3af5f735172c8cafaadc3ab42d473e9de5af24a9) Thanks [@kibertoad](https://github.com/kibertoad)! - feat: `registerCellStyle` and `patchCellFont`, so styling a report is not a second pass

  Two gaps made formatting a generated sheet cost far more calls than it should.

  `setCell` already accepted a `styleId` and the stylesheet already deduped xf records,
  but there was no way to obtain a `styleId` from a style spec without first having a
  cell to hang it on. `registerCellStyle(wb, spec)` returns one:

  ```ts
  const INT = registerCellStyle(wb, { numberFormat: "#,##0", border: THIN });
  setCell(ws, row, col, value, INT);
  ```

  The id names a complete style rather than a patch over the target cell: an axis
  missing from `spec` renders as the workbook default even where the target cell had
  something there. `setCellStyle` and `setRangeStyle` remain the patch-an-existing-cell
  paths, so reach for those when the cells already carry formatting you want to keep.

  `appendRow` and `appendRows` take matching `{ styleIds }`, positionally aligned with
  the values and reused for every row of an `appendRows` call:

  ```ts
  appendRows(ws, rows, { styleIds: [TEXT, INT, INT] });
  ```

  A column with a style id is written even when its value is empty, so a
  bordered-but-blank input column survives the append. Ids past a row's last value
  therefore add styled blank cells and widen the sheet; trim the array per row
  (`styleIds: columnStyles.slice(0, values.length)`) when the input is ragged.

  Saving now throws `OpenXmlSchemaError` when a cell's `styleId` names no entry in its
  workbook's `cellXfs` pool, naming the sheet and cell. Excel drops such a sheet behind
  the repair dialog, and an id reused across two workbooks is the easy way to get there.

  `patchCellFont(wb, cell, patch)` merges a partial font over the cell's current one.
  `setCellFont(wb, c, makeFont({ bold: true }))` is a legal call that registers a font
  with no `<name>` and no `<sz>`, after which Excel, LibreOffice and Sheets each
  substitute a different default and nothing warns. A field set to `undefined` is
  removed rather than kept, so `setBold`, `setItalic`, `setStrikethrough`,
  `setUnderline`, `setFontSize`, `setFontName` and `setFontColor` are all now this
  function with one field filled in, rather than seven hand-rolled merges.

- [#138](https://github.com/office-kit/xlsx/pull/138) [`502d305`](https://github.com/office-kit/xlsx/commit/502d3050f077c0d5107b742ebb23ddb564bfa27d) Thanks [@kibertoad](https://github.com/kibertoad)! - feat!: `addTable` / `addExcelTable` now reject a table that disagrees with the sheet under it

  A table whose `columns` count did not match the width of its `ref`, whose `ref` could not contain its
  header and totals rows, whose column names were empty or duplicated, or whose header
  cells did not hold the column names it declared, produced a workbook Excel treats as
  damaged. Excel "repairs" it by dropping the table, so the mistake surfaced as missing
  filters and a broken structured reference in the delivered file, a long way from the call
  that caused it.

  All of those now throw `OpenXmlSchemaError` at the call, naming the table and the offending
  header cell. A header cell has to hold text (a string, rich text, or a formula caching a
  string), since that is what Excel keeps in `tableColumn/@name`: write a numeric or date
  header as a string. Write the header row before adding the table, or pass
  `headerRowCount: 0` for a genuinely header-less table. Header-only tables with no
  data rows remain valid. Header and totals row counts must be unsigned 32-bit integers.

  `loadWorkbook` and `saveWorkbook` are unchanged: a mismatched table read from an input file
  still loads, and still saves, so read-modify-write of someone else's file keeps working.

  This can newly throw for code that previously appeared to work. Those are exactly the files
  Excel was already repairing.

## 0.11.1

### Patch Changes

- [#132](https://github.com/office-kit/xlsx/pull/132) [`6244117`](https://github.com/office-kit/xlsx/commit/6244117e0c142c3c65ab1be80624e3f6be070dd6) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: `loadWorkbook` returned numeric character references such as `&#20219;` as literal text instead of decoding them ([#131](https://github.com/office-kit/xlsx/issues/131)). Decimal and hexadecimal references are now decoded in text and attributes, including the inline and shared strings openpyxl writes. A reference to a character XML does not allow (for example `&#0;`) now fails the load with `OpenXmlSchemaError` instead of being kept as literal text.

- [#133](https://github.com/office-kit/xlsx/pull/133) [`dcbc31a`](https://github.com/office-kit/xlsx/commit/dcbc31a5f7a9408301c4726617eab43cadf3e340) Thanks [@kibertoad](https://github.com/kibertoad)! - fix: types were silently lost on `moduleResolution: node16` / `nodenext`

  Relative imports inside the shipped `.d.ts` files had no file extension
  (`from './load'`), which Node's ESM rules reject. Consumers on
  `moduleResolution: node16` or `nodenext` got TS2834 inside `node_modules`,
  where the usual `skipLibCheck: true` discarded it, so every symbol imported
  from `@office-kit/xlsx/*` degraded to an error type: no autocomplete, and no
  type errors reported against the library's API.

  The declarations now carry `.js` extensions, so all of `node16`, `nodenext` and
  `bundler` resolve the full type graph with `skipLibCheck: false`. No runtime
  behaviour, export name or type signature changed.

  `moduleResolution: node10` remains unsupported, since the subpaths are declared
  only through `exports`.

## 0.11.0

### Minor Changes

- [#129](https://github.com/office-kit/xlsx/pull/129) [`7c2e3f9`](https://github.com/office-kit/xlsx/commit/7c2e3f98cad33a00fb2b7c7b4fe891affbcbbdd5) Thanks [@matt-felicity](https://github.com/matt-felicity)! - Expose the workbook `date1904` flag from `loadWorkbookStream` so streamed Excel date serials can be converted with the correct epoch.

## 0.10.0

### Minor Changes

- [#124](https://github.com/office-kit/xlsx/pull/124) [`3a8db65`](https://github.com/office-kit/xlsx/commit/3a8db654fb9b218d68c6ab9683e9d95b13a11995) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: an invalid hyperlink target is now rejected instead of producing a file Excel refuses ([#117](https://github.com/office-kit/xlsx/issues/117))

  `setHyperlink` accepted any string as `target` and the writer emitted it
  verbatim into the worksheet rels, so `{ target: 'https:// www.example.com' }`
  saved without complaint and Excel then reported "We found a problem with some
  content in …" and dropped content on repair. One bad row poisoned an entire
  export.

  The rels `Target` attribute is `xsd:anyURI`. Authoring a hyperlink now throws
  an `OpenXmlSchemaError` when the target is empty, contains whitespace or a
  control character, or has a non-ASCII host (percent-encoding does not rescue
  that one — the host has to be punycode). The check runs in `makeHyperlink`, so
  `setHyperlink`, `addUrlHyperlink` and `addMailtoHyperlink` are all covered.

  Reading is deliberately unaffected: a workbook whose rels already carry a
  broken target still loads, so it can be inspected and repaired.

  **Behaviour change:** `makeHyperlink`, `setHyperlink`, `addUrlHyperlink` and
  `addMailtoHyperlink` now throw on a target they previously accepted. That is
  the point of the change — the alternative was a workbook Excel refuses — but
  it can surface at a call site that used to succeed.

- [#122](https://github.com/office-kit/xlsx/pull/122) [`3bd3940`](https://github.com/office-kit/xlsx/commit/3bd39400e5731178ffd739a6da9bcd37ba76714e) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: rich text was flattened to plain text on save ([#114](https://github.com/office-kit/xlsx/issues/114))

  A shared string built from `<r>` runs came back as a single plain `<t>`: the
  loader collapsed rich-text `<si>` entries into their concatenated text before
  the worksheet reader saw them, so a cell with a bold first half round-tripped
  uniformly unformatted. Inline strings (`<c t="inlineStr"><is>`) lost their runs
  the same way on read.

  Rich-text `<si>` and `<is>` bodies now go through one CT_Rst parser and become
  `{ kind: 'rich-text' }` cell values, so per-run fonts survive. Cells holding
  rich text are written into `xl/sharedStrings.xml` as `<si><r>…</r></si>` —
  where Excel itself stores them — instead of being inlined per cell, so a
  formatted string repeated across cells costs one entry rather than one copy
  each. `<r>` runs with no `<rPr>` stay separate runs, since that is how Excel
  writes the unformatted half of a rich string.

  **Behaviour change:** a cell whose shared string is rich text now reads back
  as `{ kind: 'rich-text', runs }` where it used to read back as the concatenated
  plain string, and rich-text cells serialise into `xl/sharedStrings.xml` rather
  than as `t="inlineStr"`. Code that assumed `getCell(...).value` was a string
  for those cells needs to handle the union.

### Patch Changes

- [#125](https://github.com/office-kit/xlsx/pull/125) [`c5d706f`](https://github.com/office-kit/xlsx/commit/c5d706fd81275b1c14fa0997dcc3bd4ec6f93e41) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: drawing shapes were rewritten as chart graphicFrames with an empty r:id ([#110](https://github.com/office-kit/xlsx/issues/110))

  An anchor holding anything the model doesn't cover — a shape, a group, a
  connector — was written back as a chart `<xdr:graphicFrame>` carrying
  `<c:chart r:id=""/>`, so a rectangle came back as "Chart 1" with a dangling
  reference, its geometry gone and no drawing rels part written at all. Anchors
  sitting inside Excel's `<mc:AlternateContent>` wrapper were dropped outright,
  leaving an empty `<xdr:wsDr/>`. Together those are what still broke worksheets
  carrying form controls after [#105](https://github.com/office-kit/xlsx/issues/105): the worksheet side was fixed, but the shapes
  that draw the controls live in `drawing1.xml`.

  Unmodeled drawing content is now kept as the verbatim source XML and written
  back untouched — the whole anchor element, so `editAs` and the
  `<xdr:clientData fLocksWithSheet="0">` flags form controls rely on survive too
  — and `<mc:AlternateContent>` wrappers keep their place in document order,
  which is z-order. Relationships those nodes reference keep their original ids
  and their target parts, and the writer allocates ids for modeled charts and
  pictures around them.

- [#121](https://github.com/office-kit/xlsx/pull/121) [`2f7e6c4`](https://github.com/office-kit/xlsx/commit/2f7e6c4e2c3a9bb8051d74d83187513331fa1b58) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: an empty cached formula value and its `t="str"` were dropped on save ([#115](https://github.com/office-kit/xlsx/issues/115))

  `<c r="A1" t="str"><f>[1]Extern!$A$1</f><v/></c>` came back as
  `<c r="A1"><f>[1]Extern!$A$1</f></c>`: the reader collapsed an empty `<v/>`
  and an absent `<v>` into the same "no cached value" state, so the writer had
  nothing left to emit the type from.

  It matters for formulas that reference another workbook — Excel cannot
  recalculate those without opening the other file, so the cached value is the
  only thing it has to display. An empty `<v/>` is now read as a cached empty
  string and written back with its `t="str"`; a genuinely absent `<v>` still
  means "no cached value", and an empty `<v/>` on a numeric cell still carries
  no number.

- [#123](https://github.com/office-kit/xlsx/pull/123) [`b3dd9c0`](https://github.com/office-kit/xlsx/commit/b3dd9c0650d7399d07b919444cf4e30290753e94) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: cells without a value were dropped on save, leaving defined names dangling ([#111](https://github.com/office-kit/xlsx/issues/111))

  `<c r="B1"/>` and `<c r="C1" s="0"/>` disappeared from the saved sheet: the
  writer skipped any cell whose value was `null` unless it also carried a
  non-default style. Dropping a genuinely unreferenced empty cell is harmless,
  but an empty cell can still be the target of a `definedName`, a form control's
  `fmlaLink`, or a drawing or comment anchor — and those were left pointing at
  nothing.

  A valueless cell now always emits. `deleteCell` — not `setCell(…, null)` — is
  how a caller says the cell is gone; cells that were never in the model are
  still not written.

- [#119](https://github.com/office-kit/xlsx/pull/119) [`48b0d4d`](https://github.com/office-kit/xlsx/commit/48b0d4d7b91d90954b21488edeed97867fe36a23) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: styles.xml lost dxfs, tableStyles, colors and extLst on save ([#113](https://github.com/office-kit/xlsx/issues/113))

  Everything after `<cellStyles>` in `xl/styles.xml` was dropped when a loaded
  workbook was saved again. `<dxfs>` was only written when the pool held
  entries, and `<tableStyles>`, `<colors>` and `<extLst>` — where Excel keeps
  custom table styles, the MRU colour palette and the x14/x15 slicer and
  timeline styles — were never read in the first place. Every workbook Excel
  writes carries at least `<dxfs>` and `<tableStyles>`, so this hit essentially
  any real file.

  `Stylesheet` now carries the unmodeled tail (`stylesXmlTail`) verbatim and the
  writer re-emits it after `<dxfs>`, where `CT_Stylesheet` (ECMA-376 §18.8.39)
  puts it. `<dxfs count="0"/>` is now always written, matching Excel — a
  conditional-formatting rule's `dxfId` is an index into that list.

- [#118](https://github.com/office-kit/xlsx/pull/118) [`c97258e`](https://github.com/office-kit/xlsx/commit/c97258e8aef0eea93ebd11877663f26cdd4408d1) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: workbook.xml was written out of CT_Workbook order ([#112](https://github.com/office-kit/xlsx/issues/112))

  `CT_Workbook` (ECMA-376 §18.2.27) is an `xsd:sequence`, so the children of
  `<workbook>` have a normative order. Saving a workbook wrote `<definedNames>`
  straight after `<sheets>` — ahead of `<functionGroups>` and
  `<externalReferences>` — and `<pivotCaches>` ahead of `<calcPr>`. A workbook
  carrying both an external reference and a defined name therefore came back as
  a package Excel refuses to open.

  `workbook.xml` is now emitted in schema order: `functionGroups`,
  `externalReferences`, `definedNames`, `calcPr`, `oleSize`,
  `customWorkbookViews`, `pivotCaches`, `smartTagPr`, `smartTagTypes`,
  `fileRecoveryPr`, then any unmodeled tail (`webPublishing`,
  `webPublishObjects`, `extLst`) that was carried over from the loaded file.

- [#126](https://github.com/office-kit/xlsx/pull/126) [`21c4a26`](https://github.com/office-kit/xlsx/commit/21c4a260c6be6ea82378742b25967ed449f13249) Thanks [@baseballyama](https://github.com/baseballyama)! - fix: workbook.xml lost mc:Ignorable and its namespace declarations on save ([#116](https://github.com/office-kit/xlsx/issues/116))

  Saving a workbook whose `xl/workbook.xml` carries the markup-compatibility
  block Excel writes into essentially every file stripped `mc:Ignorable` and the
  namespace declarations it names off the root, renamed the child prefixes
  (`x15ac:` → a generated `ns0:`, `xr:` → `x16:`), dropped the `xr2:uid` on
  `<workbookView>`, and moved `<workbookPr>` behind the `mc:AlternateContent`
  block. Excel reported the result as corrupt. This is the workbook-level sibling
  of [#105](https://github.com/office-kit/xlsx/issues/105), which fixed the same rewrite on the worksheet.

  The root's namespace declarations and `mc:Ignorable` are now carried through
  verbatim, and every captured child is written with the prefixes the root
  declares — `mc:Ignorable` names them by prefix, so the two have to agree.
  `<workbookView>` keeps namespaced attributes it doesn't model (`xr2:uid`),
  and the unmodeled head is emitted after `<workbookPr>`, where Excel puts it.

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
