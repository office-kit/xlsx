---
'@office-kit/xlsx': minor
---

feat: make generated reports cheaper to write, and their bytes reproducible

Building a styled, formula-carrying workbook from scratch took more calls than
it should, and two of them behaved in ways that quietly destroyed work. This
closes both, adds the pieces the ergonomic path was missing, and pins the
remaining source of non-determinism in the output.

**Breaking**

- `setCell(ws, row, col, value)` and `setCellByCoord(ws, coord, value)` now
  require `value`. The three-argument form read like "reach the cell at
  (row, col)" but wrote `null`, so a border or number-format pass over
  already-populated rows erased the values and formulas it walked over. Use the
  new `ensureCell(ws, row, col)` for get-or-create, or pass the value you mean
  (including an explicit `null` to blank a cell).
- `freezePanes(ws, rows, cols)` is gone; `setFreezePanes` now takes
  `'B2' | { rows, cols } | undefined`. One function covers both input shapes,
  and unlike `freezePanes` the numeric form accepts a zero on either axis, so
  "freeze two rows and no columns" is expressible: `{ rows: 2, cols: 0 }`.
- `WriteOnlyOptions.estimatedMaxRow` is gone. It was documented as reserved and
  was never read.

**Added**

- `ensureCell(ws, row, col)` in `/worksheet`: returns the cell at a coordinate,
  allocating a blank one only when it does not exist. The internal range
  helpers now use it too.
- `registerCellStyle(wb, spec)` in `/styles` returns a `styleId` for a complete
  style, which `setCell`, `setCellByCoord` and `appendRow` accept. Build each
  look once and the write carries the formatting, instead of a styling pass per
  cell afterwards. `appendRow` / `appendRows` take `{ styleIds }`, positionally
  aligned with the values; a column with an id is written even when its value is
  empty, so a bordered-but-blank input column survives the append.
- `patchCellFont(wb, cell, patch)` in `/styles` merges a partial font over the
  cell's current one. `setCellFont(wb, c, makeFont({ bold: true }))` is a legal
  call that registers a font with no name and no size, and Excel, LibreOffice
  and Sheets each substitute a different one; `setBold` and the other
  single-field setters are now this function with one field filled in.
- `makeFormula(text, { cachedValue })` in `/cell` builds a formula as a
  `CellValue`, so a formatted formula is one `setCell` write.
  `setFullCalcOnLoad` is now exported from `/workbook`, which is what a
  generated workbook wants when it leaves formulas uncached.
- `SaveOptions.mtime` (and the same option on `createWriteOnlyWorkbook`) pins
  the timestamp stamped into every ZIP entry. Without it fflate uses the wall
  clock, so two saves of the same workbook differed in bytes and golden-file
  tests were impossible. `SaveOptions.compressionLevel` is now actually wired
  through to the deflate stream instead of being documented as reserved.
- Range parameters across `/worksheet` and `/styles` accept pre-computed bounds
  (`{ minRow, minCol, maxRow, maxCol }`) as well as an A1 string, so code that
  tracks integers no longer formats a string for the callee to parse straight
  back. `writeRange` takes a `{ row, col }` anchor. `tupleToCoordinate` and
  `boundariesToRangeString` gained `absoluteCol` / `absoluteRow`, for the
  `$B$5` and `$A$4:$H$20` forms formula text needs.

**Fixed**

- Every formula setter now strips a leading `=`. OOXML stores `<f>` without it,
  so `setFormula(c, '=SUM(A1:A3)')` previously emitted `<f>=SUM(A1:A3)</f>` and
  Excel reported the file as damaged.
- `addExcelTable` now rejects a table whose column count does not match the
  range width, or whose header cells do not hold their column names. Excel
  treats both as a damaged file and repairs it by dropping the table, which
  surfaced as silent data loss far from the call that caused it. Pass
  `headerRowCount: 0` for a header-less table.

**Removed (internal)**

`setCellFormula`, `setCellArrayFormula` and `setCellRichText` were dropped from
`src/worksheet/worksheet.ts`. They were unreachable through the public subpaths
and are covered by `setCell` plus `makeFormula` / `makeRichText`.
