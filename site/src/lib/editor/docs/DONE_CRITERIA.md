# Editor: definition of done and verification status

A feature is **done** only when every column below is `PASS`. Status words:

- `PASS (stale)`: passed earlier, but the code it covers has changed since, so it
  proves nothing about the current source.
- `NOT RUN`: written or changed, never executed. Not a pass.
  Every `NOT RUN` in the Browser and Real Excel columns is also `HOLD`: those
  checks are blocked until the verification hold lifts.
- `PASS (2026-10-03)`: executed on the current source on that date (UTC); see
  "Run log" below for the exact commands.
- `FAIL`: executed and failed.
- `n/a`: the check doesn't apply.

Only the targeted runs in "Run log" have been executed on the current source.
The full suites, `check:consumer`, the browser and real Excel are still on hold.

Common bar for every feature:

- **Model**: every change goes through `doc.transact` with its parts declared
  before mutating, so it is one undo step and a protection refusal rolls the
  whole step back.
- **Save**: the saved file opens in desktop Excel without a repair prompt and
  shows the same content.
- **Reopen**: load → save → load gives the same model, and parts the editor
  doesn't model survive byte for byte.
- **Large sheets**: the feature does no work proportional to the sheet size
  on each edit or frame.

| Feature | Unit test | Browser | Real Excel | Notes / open gaps |
| --- | --- | --- | --- | --- |
| Cell entry, formulas, recalculation | PASS (stale) | PASS (stale) | PASS (stale) | |
| `_xlfn.` / `_xlpm.` storage of newer functions | PASS (2026-10-03) (`core/save-roundtrip.test.ts`) | PASS (2026-10-03): XLOOKUP shown bare, saved prefixed | NOT RUN | Stored prefixed from cell entry, Find/Replace, defined names, data validation and conditional formatting; shown bare in the formula bar, Show Formulas, Name Manager, Watch Window and `FORMULATEXT`. Evaluate Formula may still show `_xlpm.` on LET/LAMBDA parameters. |
| Save twice / reload idempotence | PASS (2026-10-03) (`tests/io/edit-save-idempotent.test.ts`, `core/save-roundtrip.test.ts`, `scripts/qa-public-api-resave.mjs` on `dist/`, including f(f(input))); `tests/consumer/src/main.ts` NOT RUN (`check:consumer` installs a tarball) | n/a | n/a | Byte equality needs the zip `mtime` and the document dates pinned, and the same input file: the persons and threaded comments get random GUIDs, so two separately built sources differ. |
| Protection refusal is atomic | PASS (2026-10-03) (`core/save-roundtrip.test.ts`); `core/protection.test.ts` PASS (stale) | PASS (2026-10-03): typing on a protected sheet shows the refusal, adds no undo step, and the saved file is unchanged there | Password hashes PASS (stale) | |
| Unknown OOXML parts | Library passthrough tests PASS (stale) | n/a | n/a | Not exercised through the editor's own load/save. |
| Fill, Auto Fill Options | PASS (stale) | PASS (stale) | NOT RUN | |
| Data validation alerts | PASS (stale) | PASS (stale) | NOT RUN | |
| Formula AutoComplete, argument tips, range finder | PASS (stale) | PASS (stale) | n/a | |
| Page Layout view | PASS (stale) | PASS (stale) | n/a | Drawings are placed through the paged axes; the Size fields use the raw axes on purpose, since a size must not include page gutters. |
| Focus Cell, outline symbols | PASS (stale) | PASS (stale) | NOT RUN | |
| Evaluate Formula | PASS (stale) | PASS (2026-10-03): opens on an XLOOKUP cell, one Evaluate step, Close | n/a | |
| Symbol, Workbook Statistics | none | NOT RUN | n/a | |
| Watch Window | none | NOT RUN | n/a | Add Watch now stops at 1,000 cells without walking a whole-column selection. |
| Conditional formatting | PASS (stale) | PASS (stale) | PASS (stale) | Excel 2010 extension (solid bars, negative axis) can't be written. |
| Charts (classic) | PASS (stale) | PASS (stale) | PASS (stale) | |
| Chart sheets on screen | none | NOT RUN | n/a | The sheet tab shows the chart full-window; editing still targets the last worksheet. |
| Charts (chartex: treemap etc.) | n/a | n/a | FAIL (Excel repairs) | Hidden from the UI until the library writer is fixed. |
| Shapes, text boxes, sparklines | PASS (stale) | PASS (stale) | PASS (stale) | |
| PivotTables | PASS (stale); date items PASS (2026-10-03) (`tests/worksheet/pivot-table.test.ts` › date items) | PASS (stale) | PASS (stale); date items NOT RUN | No pivot styling; captions English until Excel refreshes. |
| Threaded comments | PASS (stale) | PASS (stale) | PASS (stale) | `Workbook.persons` added; `Workbook.authors` kept as deprecated (changeset). |
| Dialogs (Format Cells, Sort, Find, Names, Page Setup, …) | PASS (stale) | Partial, stale | NOT RUN | Unprotect dialog now receives its `target` argument. |
| Table Design tab, built-in table styles | PASS (stale) | PASS (stale) | PASS (stale) | Custom table styles not rendered. |
| Subtotal, Consolidate, Advanced Filter, What-If | PASS (stale) | PASS (stale) | Subtotal PASS (stale); others NOT RUN | Consolidate links not implemented; Advanced Filter state is per session. |
| Page Setup, header/footer, print titles | PASS (stale) | PASS (stale) | PASS (stale) | `&P+n` offsets ignored when printing. |
| Navigation pane (View ▸ Navigation); Ctrl+Arrow / current region / Ctrl+End | PASS (2026-10-03) (`core/navigation.test.ts`) | PASS (2026-10-03): Ctrl+↓ A1→A8→A12→A1048576, Ctrl+→, Ctrl+End B12, Ctrl+A A1:B8, pane lists both sheets and switches to one | n/a | Both live in `core/navigation.ts`. |
| Message tables have no duplicate keys | PASS (stale) (`i18n/i18n.test.ts`; keys added since) | n/a | n/a | |

Checks not run on the current source: `svelte-check`, `pnpm test:editor`,
`pnpm test`, `pnpm typecheck`, `pnpm lint`, `pnpm knip`, `pnpm build`,
`pnpm size`, `check:attw`, `check:consumer`, `check:api` (the API baseline
still lists `authors`), and the perf gates (the last run failed under load
average ~150).

## Source traces

Rows in `EXCEL_PARITY.md` marked `SOURCE WIRED / NOT RUN`, with the path that
was followed. Tracing proves the code is connected, not that it works.

| Row | UI | Action | Model change and undo | Save / load |
| --- | --- | --- | --- | --- |
| Review ▸ Protect Workbook (structure) | `ribbon/ReviewTab.svelte` button | `core/actions.ts` `toggleWorkbookProtection` → `dialogs/ProtectWorkbookDialog.svelte`; removal through `dialogs/UnprotectDialog.svelte` (password hash check) | `doc.transact` with `tx.workbook('workbookProtection')`, so one undo step; the guard in `core/protection.ts` refuses sheet structure edits | `src/io/save.ts` `serializeWorkbookProtection`; `src/io/load.ts` reads `workbookProtection`. Password hashes opened correctly in real Excel earlier: stale. |
| Home ▸ Bold / Italic | `ribbon/HomeTab.svelte` buttons; Cmd+B / Cmd+2 in `core/shortcuts.ts` | `core/actions.ts` `toggleBold` / `toggleItalic` → `format` → `core/commands.ts` `formatRanges` | `doc.transact` with `tx.cells` per range (plus column/row dimensions for whole columns/rows); `applyStyle` adds the style to the workbook's style pool | Cell `styleId` and the stylesheet are written by the library's worksheet and stylesheet writers. |
| Home ▸ Underline / Double underline | `ribbon/HomeTab.svelte` split button | `toggleUnderline(ctl, 'single' \| 'double')` → `format` → `formatRanges` | as Bold | as Bold |
| Home ▸ Merge & Center / Across / Cells / Unmerge | `ribbon/HomeTab.svelte` split button | `core/actions.ts` `merge` → `core/commands.ts` `mergeRanges` | `doc.transact` with `tx.sheet(ws, 'mergedCells')` and `tx.cells`; overlapping merges dissolved first | `src/worksheet/writer.ts` writes `<mergeCells>`. |
| View ▸ Freeze Panes / Top Row / First Column | `ribbon/ViewTab.svelte` menu | `core/actions.ts` `freezePanes` → `core/commands.ts` `freeze` | `doc.transact` with `tx.sheet(ws, 'views')`; library `setFreezePanes` | `src/worksheet/writer.ts` `serializeSheetViews` writes the pane. |
| Home ▸ Fill color | `ribbon/HomeTab.svelte` split button with `ui/ColorGrid.svelte` (theme tints, standard colours, No Fill, More Colors via a colour input) | `core/actions.ts` `setFillColor` → `format` → `formatRanges` | as Bold | as Bold |
| Home ▸ Font color | same split button pattern | `setFontColor` → `format` → `formatRanges` | as Bold | as Bold. "Automatic" writes theme colour 1 (dk1), as the default font does; whether Excel writes `auto` here instead is HOLD (needs real Excel). |
| Home ▸ Increase / Decrease Decimal | `ribbon/HomeTab.svelte` buttons | `stepDecimals` rewrites the number format code → `format` | as Bold | Number format codes are written to the stylesheet. |
| Review ▸ Protect Sheet | `ribbon/ReviewTab.svelte`, Home ▸ Format menu | `dialogs/ProtectSheetDialog.svelte` (SHA-512 hash, then commit) | `doc.transact` with `tx.sheet(ws, 'sheetProtection')`; the guard in `core/protection.ts` then refuses edits to locked cells | `src/worksheet/writer.ts` `serializeSheetProtection`. Hashes opened correctly in real Excel earlier: stale. |
| Home ▸ AutoSum (Sum, Average, Count Numbers, Max, Min, More Functions…) | `ribbon/HomeTab.svelte` split button; More Functions opens `insertFunction` | `core/actions.ts` `autoSum`: one cell starts an edit with the proposed range (committed through `commitEdit` → `parseInput` → `doc.transact`); a block writes formulas with `putFormula` inside `doc.transact` | `tx.cells` on the target row/column | Formulas are saved by the library worksheet writer. The Formulas-tab AutoSum menu has no More Functions entry, so that row stays `todo`. |
| Data ▸ Sort A to Z / Z to A | `ribbon/DataTab.svelte` | `core/data.ts` `quickSort` → `sortRange` (AutoFilter range, else the current region with header detection) | `doc.transact('Sort')` with `tx.cells` on the body | cell values and styles saved as usual |
| Data ▸ Filter | `ribbon/DataTab.svelte`; Cmd+Shift+F / Cmd+Shift+L in `core/shortcuts.ts` | `ctl.toggleFilter` → `core/filter.ts` `toggleAutoFilter` | `doc.transact('Filter')` with `tx.sheet(ws, 'autoFilter', 'rowDimensions')` | `src/worksheet/writer.ts` `serializeAutoFilter` |
| Home ▸ Font family, Font size, Increase / Decrease font size | `ribbon/HomeTab.svelte` combos and buttons (size 1–409) | `setFontName` / `setFontSize` / `stepFontSize` → `format` → `formatRanges` | as Bold | as Bold. The Cmd+Shift+> / < shortcuts were not traced. |
| Home ▸ Borders presets (all 13) | `ribbon/HomeTab.svelte` `BORDERS` list; More Borders opens Format Cells › Border | `applyBorder` → `format` | as Bold | Borders are written to the stylesheet. |
| Home ▸ Top / Middle / Bottom, Left / Center / Right | `ribbon/HomeTab.svelte` toggles | `setVAlign` / `setHAlign` (pressing the active one clears it) → `format` | as Bold | as Bold |
| Home ▸ Orientation (5 angles, none, Format Cell Alignment) | `ribbon/HomeTab.svelte` menu | `setRotation` → `format`; the last item opens Format Cells › Alignment | as Bold | as Bold |
| Home ▸ Decrease / Increase Indent | `ribbon/HomeTab.svelte` buttons | `stepIndent` → `format` (General alignment becomes Left, as in Excel) | as Bold | as Bold |
| Home ▸ Percent / Comma style | `ribbon/HomeTab.svelte` buttons; Cmd+Shift+% in `core/shortcuts.ts` | `setNumberFormat` → `format` | as Bold | as Bold |
| Home ▸ Insert (Cells…, Sheet Rows, Sheet Columns, Sheet) | `ribbon/HomeTab.svelte` split button | `insertCellsSmart` / Insert Cells dialog, `insertLines`, `insertSheet` | `doc.transact`; rows/columns declare a structural snapshot once (`declareStructural`) and then shift every span | Library writers. Delete stays `todo`: the menu has no Delete Table Rows / Columns. |
| Formulas ▸ Name Manager (New, Edit, Delete) | `ribbon/FormulasTab.svelte` | `dialogs/NameManagerDialog.svelte`; New/Edit open `dialogs/DefineNameDialog.svelte` (validated with `validateName`) | `doc.transact` with `tx.workbook('definedNames')` (structural, so formulas recalculate) | `src/io/save.ts` writes `<definedNames>`. |
| Data ▸ Sort… | `ribbon/DataTab.svelte` | `dialogs/SortDialog.svelte` → `core/data.ts` `sortRange` (rows or columns) | `doc.transact('Sort')` with `tx.cells` on the body | cells saved as usual |
| Data ▸ Clear / Reapply | `ribbon/DataTab.svelte` | `core/filter.ts` `clearAllFilters` / `reapplyFilters`; `core/advanced-filter.ts` `clearAdvancedFilter` | `doc.transact` with `tx.sheet(ws, 'rowDimensions')` (hidden rows) | AutoFilter and row hidden flags written by the worksheet writer |
| Data ▸ Advanced Filter | `ribbon/DataTab.svelte` | `dialogs/AdvancedFilterDialog.svelte` → `core/advanced-filter.ts` | in place: `tx.sheet(ws, 'rowDimensions')`; copy: `tx.cells` on the output | as above. In-place state is per session (known gap). |
| Data ▸ Text to Columns | `ribbon/DataTab.svelte` | `dialogs/TextToColumnsDialog.svelte` → `core/text-to-columns.ts` `textToColumns` | `doc.transact('Text to Columns')` | cells saved as usual |
| Data ▸ Remove Duplicates | `ribbon/DataTab.svelte`, `ribbon/TableDesignTab.svelte` | `dialogs/RemoveDuplicatesDialog.svelte` → `core/data.ts` `removeDuplicates` | `doc.transact('Remove Duplicates')`, refusal (outline / SUBTOTAL) before the transaction | cells, table ref + autoFilter, hyperlinks, comments |
| Data ▸ Consolidate | `ribbon/DataTab.svelte` | `dialogs/ConsolidateDialog.svelte` → `core/consolidate.ts` | `doc.transact` with `tx.cells(target)` and `tx.sheet(ws, 'dataConsolidate')` | `src/worksheet/writer.ts` `serializeDataConsolidate`. "Create links to source data" not implemented. |
| Data ▸ What-If (Scenario Manager, Goal Seek, Data Table) | `ribbon/DataTab.svelte` menu | `dialogs/ScenarioManagerDialog.svelte`, `GoalSeekDialog.svelte`, `DataTableDialog.svelte` → `core/what-if.ts` | `doc.transact` per command (`tx.cells` on the changed cells / table body) | scenarios via `serializeScenarioList`; data table `TABLE()` formulas with values computed by the editor |
| Data ▸ Subtotal | `ribbon/DataTab.svelte` | `dialogs/SubtotalDialog.svelte` → `core/subtotal.ts` | `doc.transact('Subtotal')`, structural edits declared once; Remove All in its own step | SUBTOTAL formulas and outline levels; opened in real Excel earlier: stale |
| Page Layout ▸ Margins (Normal / Wide / Narrow, Custom), Orientation, Size, Scale to Fit (width, height, scale) | `ribbon/PageLayoutTab.svelte` menus and fields; Custom Margins opens Page Setup › Margins | `core/actions.ts` `setMargins` / `setOrientation` / `setPaperSize` / `setFitTo` / `setPrintScale` → `editPage` | `doc.transact` with `tx.sheet(ws, 'pageSetup', 'pageMargins', 'printOptions', 'rowBreaks', 'colBreaks', 'sheetProperties')` | `src/worksheet/writer.ts` `serializePageMargins` / `serializePageSetup` |
| Page Layout ▸ Print Area (Set / Clear) | `ribbon/PageLayoutTab.svelte` | `setPrintArea` | `doc.transact` with `tx.workbook('definedNames')` (`_xlnm.Print_Area`) | `<definedNames>` in `src/io/save.ts` |
| Page Layout ▸ Print Titles, Page Setup dialog | `ribbon/PageLayoutTab.svelte`; Insert ▸ Header & Footer opens the same dialog | `dialogs/PageSetupDialog.svelte` → `dialogs/page-setup.ts` | `doc.transact('Page Setup')` with the page fields plus `tx.workbook('definedNames')` (`_xlnm.Print_Titles`) | page fields and names as above; Print_Titles opened in real Excel earlier: stale |
| Insert ▸ PivotTable | `ribbon/InsertTab.svelte` | `dialogs/CreatePivotTableDialog.svelte` → `core/pivot.ts` | `doc.transact` with `tx.sheet(ws, 'pivotTables')`, `tx.cells` on the report, and `tx.workbook('sheets', …)` for a new sheet | `src/worksheet/pivot-xml.ts` via `src/io/save.ts`; opened in real Excel earlier: stale |
| Insert ▸ Table | `ribbon/InsertTab.svelte` | `dialogs/CreateTableDialog.svelte` | `doc.transact` with `tx.sheet(ws, 'tables', 'autoFilter')` and `tx.cells` on the header row | table parts written by the library |
| Insert ▸ Link | `ribbon/InsertTab.svelte` | `dialogs/HyperlinkDialog.svelte` | `doc.transact` with `tx.cells` and `tx.sheet(ws, 'hyperlinks')` | `src/worksheet/writer.ts` `serializeHyperlinks`. The Hyperlink cell style is not applied (known gap). |
| Insert ▸ Symbol | `ribbon/InsertTab.svelte` | `dialogs/SymbolDialog.svelte` inserts into the cell edit | committed through `commitEdit` → `parseInput` → `doc.transact` | cell text saved as usual |
| Insert ▸ Sparklines | `ribbon/InsertTab.svelte` menu | `dialogs/CreateSparklinesDialog.svelte` → `core/sparklines.ts` | `doc.transact` with `tx.sheet(ws, 'sparklineGroups')` | `serializeSparklineExt` in the worksheet extLst; opened in real Excel earlier: stale |
| Insert ▸ Recommended Charts, Column/Bar, Line/Area, Pie/Doughnut, Scatter/Bubble | `ribbon/InsertTab.svelte` galleries (`ChartMenuButton`) and `dialogs/InsertChartDialog.svelte` | `core/charts.ts` `insertChart` | `doc.transact` with `tx.sheet(ws, 'drawing')` | drawing and chart parts written by the library; opened in real Excel earlier: stale. Chartex kinds are hidden (Excel repairs them). |
| Home ▸ Conditional Formatting menu (Highlight, Top/Bottom, Data Bars, Color Scales, Icon Sets, New Rule, Clear Rules, Manage Rules) | `ribbon/HomeTab.svelte` menu | `dialogs/ConditionalFormattingDialog.svelte` with a `preset` → `core/conditional-format.ts` | `doc.transact` with `tx.sheet(ws, 'conditionalFormatting')`; new formats go into the append-only dxf pool | conditional formatting written by the worksheet writer; opened in real Excel earlier: stale. The Excel 2010 extension (solid bars, negative axis) can't be written. |
| Data ▸ Show Detail / Hide Detail | `ribbon/DataTab.svelte` | `core/actions.ts` `showDetail` → `core/outline.ts` `toggleRun` | `doc.transact` with `tx.sheet(ws, 'rowDimensions' \| 'columnDimensions')` | outline levels and hidden flags written with the row/column dimensions |
| Data ▸ Data Validation…, Circle Invalid Data, Clear Validation Circles | `ribbon/DataTab.svelte` menu | `dialogs/DataValidationDialog.svelte` → `dialogs/data-validation.ts` `applyValidation`; circles are a screen overlay (`circleInvalid`), as in Excel they are not saved | `doc.transact('Data Validation')` with `tx.sheet(ws, 'dataValidations')` | `src/worksheet/writer.ts` writes `<dataValidations>` |

## Do the new regression tests fail without the fix?

Checked by reading the code, not by running it:

| Test | Without the change it fails because |
| --- | --- |
| `save-roundtrip` › `_xlfn.` prefix | `parseInput` stored `XLOOKUP(…)` bare, so the expected `_xlfn.XLOOKUP(…)` differs. |
| `save-roundtrip` › not prefixed twice | Guards `toStorageFormula`'s own check; passes either way and only catches a regression in it. |
| `save-roundtrip` › save again changes nothing | Characterization: no known bug it reproduces. |
| `save-roundtrip` › refused transaction | Without `tx.rollback()` B2 keeps 7. |
| `edit-save-idempotent` | Characterization of the save path; it fails if anything in it starts depending on the clock or on randomness. |
| `pivot-table` › date items | Without the change the label cell has the General format. |
| `navigation` | New feature; no earlier implementation. |

## Not implemented yet

Each needs a browser or real-Excel check to be trusted:

- Themes / theme colours and fonts (Page Layout ▸ Themes). A new workbook ships
  no theme part, so this means writing a full theme XML that Excel accepts.
- Grouped sheet editing (several tabs selected, one edit applied to all).
- Split panes (View ▸ Split).
- Custom Views, Spelling, cell checkboxes.
- `&P+n` page-number offsets when printing.
- Chart sheet editing (Chart Design tabs on a chart sheet).
- Chartex charts (treemap, sunburst, histogram, box & whisker, waterfall, funnel) in the library writer.
- `#CALC!` is written as-is; Excel stores it as `#VALUE!`.
- Home ▸ Delete ▾ has no Delete Table Columns, so the Delete row stays `todo`. Not done because it is not a small fix: the band delete shrinks the table's `ref` but not `table.columns`, and structured references (`Sales[Region]`) to the deleted column are not rewritten to `#REF!`; both need new table-column code. The Formulas-tab AutoSum menu has no More Functions….
- Added; regressions PASS (2026-10-03), browser and real Excel NOT RUN (`core/tables.test.ts` "Delete Table Rows", `core/controller.test.ts` "Shrink Text to Fit" / "Remove Hyperlinks" / "Remove Page Break" / "Create from Selection"):
  - Home ▸ Delete ▸ Delete Table Rows (`core/tables.ts` `deleteTableRows`, shown when the active cell is in a table): Delete Cells ▸ shift up over the table's columns, so `ref` and its AutoFilter shrink; deleting every data row keeps one empty row. Whether Excel also shifts cells below the table in those columns is HOLD (needs real Excel).
  - Home ▸ Wrap Text ▾ ▸ Shrink Text to Fit (`A.toggleShrink`, `alignment.shrinkToFit` through `formatRanges`).
  - Home ▸ Clear ▾ ▸ Remove Hyperlinks and the cell context menu's Remove Hyperlink (`ClearKind` `removeHyperlinks`): removes the links and resets the style of the linked cells only; Clear Hyperlinks still keeps the style. Which cells Excel resets is HOLD.
  - Page Layout ▸ Breaks ▸ Remove Page Break (`A.removePageBreak`): removes the breaks on the active cell's top and left edges; no-op without a transaction when there is none.
  - Each declares its parts before mutating, so a protection refusal throws before any change and leaves no undo step; none of them touch passthrough parts.
- **Major, fixed (regression PASS 2026-10-03):** Formulas ▸ Create from Selection (`core/actions.ts` `createNamesFromSelection`) now runs `validateName`: a label such as `Q1`, `R` or `C` gets a leading `_`, labels that are still invalid or already taken are skipped, and nothing is written when no name is left. Whether Excel renames such labels the same way is HOLD (needs real Excel). Excel also asks which edge holds the labels (dialog); this one still picks the top row or the left column on its own, so the row stays `todo`.
- Formulas ▸ Calculation Options has no "Automatic except for data tables", and Calculate Sheet recalculates the whole workbook.
- Table Design tab has no Summarize with PivotTable or Insert Slicer; Sparkline tab has no Edit Data, Axis or Group/Ungroup; the sheet-tab menu has no Select All Sheets (grouped sheets are not implemented). These rows stay `todo`.
- View ▸ Normal / Page Layout / Page Break Preview and the view flags (gridlines, headings, show formulas) change sheet-view state outside `doc.transact`: they are saved but are not undo steps. They are not eligible for `SOURCE WIRED` under its definition and stay `todo`.
- Data ▸ Group has no Auto Outline and Ungroup has no Clear Outline; Formulas ▸ function library has no Recently Used menu; Insert Function is a dialog that inserts the function name, not Excel's Formula Builder with argument fields and a live result.
- Home ▸ Format ▾ has no Tab Color (the sheet-tab path exists separately), so the Format row stays `todo`; Home ▸ Fill ▾ has no Across Worksheets.
- Insert/Delete Rows/Columns over several selected spans used to snapshot the whole workbook once per span; it now snapshots once (`core/actions.ts` `insertLines` / `deleteLines`). NOT RUN.
- API baseline (`tests/consumer/api-baseline.json`) for the new exports and the `authors` → `persons` change.

## Run log

Node v26.6.0, worktree `xlsx-editor`, run one at a time, 2026-10-03 (UTC).

| Start (UTC) | Command | Exit | Result | Time |
| --- | --- | --- | --- | --- |
| 19:07:43 | `pnpm exec vitest run tests/io/edit-save-idempotent.test.ts tests/worksheet/pivot-table.test.ts` | 0 | 2 files, 15 tests passed | 3 s |
| 19:07:50 | `pnpm exec vitest run --root site --config vitest.editor.config.ts …/save-roundtrip …/navigation …/controller …/tables` | 0 | 4 files, 29 tests passed | 2 s |
| 19:07:56 | `pnpm typecheck` | 0 | no errors | 4 s |
| 19:08:03 | `pnpm build` | 0 | built `dist/` | 3 s |
| 19:08:10 | `pnpm --dir site check` | 1 | **FAIL**, 14 errors: see below | 7 s |
| 19:11:06 | `pnpm exec vitest run --root site --config vitest.editor.config.ts …/navigation …/controller …/evaluate-formula …/charts …/data-tools` | 0 | 5 files, 41 tests passed | 3 s |
| 19:11:16 | `pnpm doc:api` | 0 | generated the gitignored `site/src/lib/server/api-data.json` (as CI's site build does) | 4 s |
| 19:11:23 | `pnpm --dir site check` | 0 | no errors | 8 s |
| 19:12:09 | public-API script, `edit` twice in two Node processes on one input file | 0, 0 | identical output bytes; reload + resave byte-identical; `_xlfn.XLOOKUP`, the edit, the threaded comment and the person survive | <1 s each |
| 19:12:48 | `pnpm exec vitest run --root site --config vitest.editor.config.ts …/save-roundtrip …/tables` (again, after the `navigation.ts` restore) | 0 | 2 files, 16 tests passed | 3 s |
| 19:15:00 | `node scripts/qa-public-api-resave.mjs` (= `pnpm qa:public-api-resave`), twice in two processes, on the 19:08 `dist/` (no `src/` file newer) | 0, 0 | both runs: input `51feaf6a…`, f(input) `a4494011…`, f(f(input)) `a4494011…`; f(input) run twice is identical, f(f(input)) equals f(input), reload + resave unchanged; A2 `_xlfn.XLOOKUP`, A3, one threaded comment and one person with their pinned ids, no duplicates | <1 s each |

The first `site check` failure:
- `core/navigation.ts` had been overwritten by the Navigation pane change, losing `dataEdge`, `currentRegion` and `lastUsedCell`. The editor tests still passed because Vite turns a missing named import into `undefined` and nothing called it. The original was recovered from the session log, checked line by line against two later reads, and merged with the pane code. `core/navigation.test.ts` now exercises the three functions.
- `dialogs/EvaluateFormulaDialog.svelte` named a variable `state`, which makes `$state` read as a store subscription. Renamed to `evaluation`.
- `api-data.json` was missing: generated by `pnpm doc:api`.

The 19:12 public-API script (scratch copy) imports each subpath through `package.json` `exports` → `dist/`, with no install. It is not `check:consumer`, which packs and installs a tarball and was not run.

`scripts/qa-public-api-resave.mjs` is the reusable form: it needs `pnpm build`
first, imports only the published subpaths through `package.json` `exports`,
pins the zip mtime, the document dates and the person / comment ids, and fails
with a non-zero exit on any mismatch. f(f(input)) feeds the first saved output
back in as the second run's input.

### Browser run, 2026-10-03 (UTC)

One Chromium (Playwright MCP, already installed) against a dev server this run
started on `127.0.0.1:5231` from this worktree and stopped afterwards. The
existing preview on 5199 was left alone: its parent had exited, so its owner
could not be confirmed, and another project's proxy listens on the same port.
The fixture was built with the public API on `dist/`: `_xlfn.XLOOKUP` in
Data!A2, a threaded comment and person on B1, an unknown `customXml/item1.xml`
part, and a protected sheet `Locked`. Page errors and console errors/warnings
on that page: 0.

| UTC | Case | Result |
| --- | --- | --- |
| 19:29:59 | open → A2 shows `=XLOOKUP(…)` → type `=SUM(B1:B2)` in A3 → Undo (A3 empty) → Redo → save (download) → reopen with the public API | PASS: A2 `_xlfn.XLOOKUP` kept, A3 saved, comment and person ids unchanged, `customXml/item1.xml` byte-identical, `Locked` still protected |
| 19:30–19:31 | Ctrl+Arrow, Ctrl+End, Ctrl+A, View ▸ Navigation, click `Locked` in the pane | PASS (values in the table row above) |
| 19:31:54 | type on `Locked`!A1 → refusal dialog → OK → one Undo takes back Data!A3 (the last real edit) → Redo → save → reopen | PASS: `Locked`!A1 still `keep`, no extra undo step, file otherwise as before |
| 19:32 | Formulas ▸ Evaluate Formula on A2, Evaluate once, Close | PASS |

Found and fixed during the run: after Enter in the Name Box, focus went to
`<body>`, so typing did nothing (Excel puts the keys into the cell). The grid
now also refocuses on `ctl.gridFocusRequest`, which the Name Box bumps
(`core/controller.svelte.ts`, `grid/GridView.svelte`, `ui/FormulaBar.svelte`).
Re-run after the fix: `pnpm --dir site check` exit 0 (19:32:52, 9 s) and the
controller / navigation / save-roundtrip tests, 3 files and 20 tests, exit 0
(19:33:01, 2 s). The fix has no unit test: the editor tests run without a DOM;
it was checked in the browser run above.

Screenshots, the fixture, the saved files and the dev-server log are in this
session's scratchpad under `browser-qa/`. Playwright MCP also wrote one empty
snapshot and a copy of the download into the main checkout's gitignored
`.playwright-mcp/`; those two files were moved to the scratchpad.

### Keyboard focus after ribbon / Name Box controls, 2026-10-03 (UTC)

After these controls, focus stayed on the control or fell to `<body>`, and the
grid only listens on its own textarea, so typing, arrows, Undo and Save did
nothing until a click. With the number format list focused, the next typed
letter even picked another format (`f` → Fraction). Each path now bumps
`ctl.gridFocusRequest`:

- Name Box Escape and a pick from the defined-names list (`ui/FormulaBar.svelte`);
- Enter in the font name and font size boxes (`ribbon/HomeTab.svelte`), which
  also call `preventDefault()` now: without it the Enter landed in the
  refocused cell editor as a leading newline;
- a change in the number format list (`ribbon/HomeTab.svelte`).

No unit test: the editor tests run without a DOM and no DOM library is
installed, so the regression is the browser run below, the same steps before
and after the fix. Dev server started by this run on `127.0.0.1:5231` and
stopped afterwards; 5199 untouched. Fixture: the browser-QA fixture plus the
defined name `Rate` = `Data!$D$5`.

| UTC | Run | Result |
| --- | --- | --- |
| 20:56:21 / 20:56:31 | before the fix | FAIL on all 5: focus `BODY` after Name Box Escape, names list, font name Enter, font size Enter; focus left on `Number Format` after a change, and typing `fmt2` switched it to Fraction |
| 20:57:02 | after the focus change only | 3 PASS; font name / size Enter: focus returned but the cell got `"\n…"` |
| 20:58:06 | after `preventDefault()` too | PASS on all 5: focus `Cell editor`, the typed text in the cell; then ArrowDown E6→E7, Undo empties E6, Redo restores it, Save downloads; page errors 0 |
| 20:58:23 | reopen the save with the public API | E1/D5/E3/E4/E6 hold the typed text; E3 font Arial, E4 size 14, E6 `"$"#,##0.00`; `Rate`, `_xlfn.XLOOKUP`, the comment and person, `customXml/item1.xml` and the protected sheet kept |
| 20:58:52 | `pnpm --dir site check` | exit 0, 9 s |
| 20:59:05 | `vitest … --maxWorkers=1 --minWorkers=1` | exit 1: Vitest 4 has no `--minWorkers`; nothing ran |
| 20:59:15 | `vitest run --root site --config vitest.editor.config.ts --maxWorkers=1` controller / navigation / save-roundtrip / i18n | exit 0, 4 files, 22 tests, 3 s |

Not changed in this run (held): Delete Table Columns, a Create from Selection
dialog, data-table recalculation.

Output filtering: every command in the run logs above (19:07–19:33 and
20:56–20:59) was shown through `grep -v WARN | tail`, without a full log kept.
The exit codes and pass counts are as printed, but the number of warnings
those commands emitted was not recorded and is not claimed to be zero. The
lines removed included pnpm's notice that the `pnpm` field in `package.json`
is no longer read; whether anything else was removed is unknown.

Correction: an earlier status note said the worktree HEAD had moved. It had
not. The comparison used the main checkout's git status from the start of the
session (`<repo root>`, `5a9fe27`), not this worktree's HEAD.
This worktree (`<repo root>/.claude/worktrees/xlsx-editor`) was at
`137883274a17f6f670dffa7210b684cd7dba721e` on `feat/xlsx-editor` at 20:53Z, the
same as the coordinator's baseline, with nothing staged. No git command that
changes state was run in this session.


### Remove Duplicates, 2026-10-04 (UTC)

Rules followed (Microsoft support, "Filter for unique values or remove
duplicate values"): a duplicate is decided on the displayed values of the
chosen columns, the first occurrence is kept, the whole row of the range goes
(all of its columns, not only the compared ones), values outside the range or
table don't change or move, and outlined or subtotalled data is refused.
Apache OpenOffice Calc has no in-place command, only Standard Filter ▸
Options ▸ "No duplications" (`helpcontent2/source/text/scalc/01/12090104.xhp`),
so nothing there was copied.

Inside a table the target is the whole table body; the table ref and its
autoFilter shrink by the removed rows and the total row moves up. Hyperlinks
and comments on kept rows move with them; those on removed rows go.
Comparison stays case-insensitive as before (not verified against Excel).

| UTC | Command | Result |
| --- | --- | --- |
| 02:16:24 | `vitest … --maxWorkers=1 core/data-tools.test.ts -t "Remove Duplicates"` before the fix | exit 1: 3 failed (table shrink, hyperlink remap, outline refusal), 3 passed |
| 02:17:16 | same file after the fix | exit 0, 17 tests |
| 02:18:06 | `pnpm --dir site check` | exit 0: 0 errors, 2 warnings (`CreatePivotTableDialog.svelte`, `UnprotectDialog.svelte`, both `state_referenced_locally`, not touched here) |
| 02:18:28 | `vitest run --root site --config vitest.editor.config.ts --maxWorkers=1` | exit 0, 27 files, 947 tests, 13 s |
| 02:20:16 | browser, own `vite dev` on 5231, 1280×800: table `Sales` A1:B6, Remove Duplicates on Region only | A2:A4 East / West / North, B4 5 (the Qty of the kept row), D3 and A9 sentinels unchanged; Undo restores A4:A6, Redo removes again; console 0 errors / 0 warnings, 0 page errors |
| 02:20 | reopen the saved file in the editor and with the public API | same cells; table `A1:B4`, autoFilter `A1:B4`, the North link moved A6 → A4 |

Full logs are kept outside the repo; nothing above was filtered. Screenshots:
`docs/qa/xlsx-editor-20261004/remove-duplicates-{before,after}.png`.

### Repository gate, 2026-10-04 (UTC), one command at a time

Full logs kept outside the repo, unfiltered. Not run: perf gate, browser /
conformance / render jobs (need dotnet, LibreOffice or Playwright browsers
not installed here).

| UTC | Command | Result |
| --- | --- | --- |
| 02:22:16 | `pnpm typecheck` | exit 0 |
| 02:25:23 | `pnpm lint` | **FAIL** exit 1: 282 findings, all in `site/src/lib/editor`; 256 are `prefer-const` / `no-unassigned-vars` on `.svelte` variables that the template reassigns (`bind:`), which oxlint cannot see. Not fixed: a lint-config override was refused as a CI bypass, and the choice is left to the maintainer. |
| 02:25:23 | `pnpm qa:matrix`, `qa:matrix-calibration` | exit 0, exit 0 |
| 02:25:58 | `pnpm knip` | exit 0 after un-exporting `SPARKLINE_EXT_URI` and `PivotFieldStats` (used only in their own files) |
| 02:26:06 | `pnpm build` | exit 0; tsdown warns that `external` is deprecated |
| 02:26:09 | `pnpm size` | **FAIL**: `@office-kit/xlsx/io` 122.72 kB brotli (limit 120, +2.72) and 468.31 kB minified parse (limit 435, +33.31). Budget not raised. Largest new code in that graph: `pivot-table.ts` (35 kB source, via `computePivotTable` in `save.ts`) and `pivot-reader.ts` (16 kB, via `load.ts`). |
| 02:26:26 | `pnpm test --maxWorkers=1` | exit 0: 377 files, 3516 tests, 97 s |
| 02:28:09 | `pnpm check:attw`, `check:consumer`, `check:api-calibration` | exit 0, exit 0, exit 0 (7 pass) |
| 02:28:33 | `pnpm check:api` | exit 1 (declaration drift), then after review and `--write` exit 0, hash `303ad266…` |
| 02:28:34 | `pnpm qa:public-api-resave` | exit 0, once = twice `a4494011a7a4869c` |

API review: every removed declaration line is either a re-ordered export list
or `resolveRelTarget`, which no subpath exported. `Workbook.authors` is kept as
deprecated. Object literals typed as the old `Workbook` (with `authors`,
without `persons`) no longer compile (TS2741, checked against `dist/`); the same
holds for `Worksheet` literals, which now need `threadedComments` and
`sparklineGroups`. Reading `wb.authors`, spreading a factory result and the
factories themselves still compile. `tests/consumer/api-change.json` marks the
change `breaking` because `DrawingItem.content` gains `shape`.

Correction to the gate notes above: Playwright's chromium, firefox and webkit
builds, `soffice` and `pdftoppm` are present on this machine. `pnpm
test:browser`, `qa:office`, `qa:render` and the perf gate were not run for lack
of a verification slot, not because anything is missing; they are NOT RUN, not
failed. Only the `dotnet run --project tests/conformance/sdk` steps cannot run
here (no `dotnet`). The CI render job also installs Linux-only
`fonts-liberation2`, so a local macOS render would not match CI.

Fixtures: `tests/fixtures/pivot/excel-saved-filter-pivot.xlsx` and
`tests/fixtures/threaded-comments/excel-mac.xlsx` are new in this change and
carried a real name in `docProps/core.xml` `lastModifiedBy`. It is now `Demo
User`; every other zip entry has the same content and order. No test compares
these files byte for byte. To verify in the next slot:
`tests/worksheet/threaded-comments.test.ts` and
`tests/worksheet/pivot-table.test.ts` (NOT RUN after the edit).

IO size (unmet, recorded only): `saveWorkbook` calls `computePivotTable` on
purpose, so the cache and item indices always match the current source data.
Keeping a computed result on the model instead would let a stale cache be
written after a source edit, so that idea is dropped. Candidate that keeps the
contract: move report-cell layout (`layoutReport`, only `refreshPivotTable`
needs it) out of the module `save.ts` imports, and prove saved bytes unchanged
with the pivot tests and the resave idempotence test. Not measured yet.

### Second round, 2026-10-04 (UTC): the four gate fixes and sheet copy

Every command's exit code was written to a `.exit` file right after it ran;
full logs are kept outside the repo, unfiltered. The 3516-test run above was on
older source; the counts here are for this round's source.

- Optional fields: `Workbook.persons`, `Worksheet.threadedComments` and
  `Worksheet.sparklineGroups` are optional; `createWorkbook` / `addWorksheet`
  still set `[]`, and the library and editor read them with `?? []` / `??=`.
  The same old-style `Workbook` literal compiled against `dist/` with
  `--strict --exactOptionalPropertyTypes`: TS2741 (exit 2) on the earlier
  source, exit 0 now. `tests/workbook/pre-persons-literals.test.ts` edits,
  duplicates, saves, reloads and re-saves a workbook without those fields
  (not run against the earlier source).
- IO size: `saveWorkbook` now calls `computePivotStructure`, which skips the
  aggregation, report cells and date-format probe that only a refresh needs.
  Saved bytes of five pivot shapes (first save, save after a source edit
  without refresh, after refresh, after reopen) are identical before and after
  the split. io went from 122.72 to 121.70 kB brotli and 468.31 to 465.26 kB
  parse — **still over the 120 / 435 kB budgets, which were not changed**.
  Against HEAD's io bundle (432,670 B minified), the remaining +29.5 kB is the
  new features themselves: pivot model + save (pivot-table 7.5 kB, pivot-reader
  5.9 kB, pivot-xml 5.2 kB), threaded comments 3.5 kB, sparklines 3.1 kB,
  shapes in drawing-xml +2.5 kB, save.ts +2.6 kB. No large transitive
  dependency was pulled in.
- Lint: oxlint reads only the `<script>` of a `.svelte` file. For
  `site/**/*.svelte`, `prefer-const` and `no-unassigned-vars` now run in ESLint
  with eslint-plugin-svelte's parser, which reads the template; every other
  rule stays in oxlint, and `pnpm lint` runs both. On the same probe file
  ESLint reports only the truly constant variable while oxlint also reports the
  one written by `bind:value`. Over the editor, the template-aware rules found
  7 real `prefer-const` cases (fixed) instead of 256. The 26 other oxlint
  findings were fixed in code. The recommended eslint-plugin-svelte set (70
  findings, mostly in the docs site) was not adopted.
- Fixtures: `lastModifiedBy` set to `Demo User`; the two tests that read them
  pass.
- Sheet copy: the editor's Copy Sheet kept the source's threaded-comment ids,
  which Excel rejects across sheets. It now gives the copy fresh thread and
  mention ids (replies follow their root, authors unchanged), as the library's
  `duplicateSheet` does; that one now renews mention ids too. Regressions fail
  on the earlier source and pass now, including undo, redo, save, reopen and a
  second copy.

| UTC | Command | Result |
| --- | --- | --- |
| 07:04:10 | fixture tests (threaded-comments, pivot-table) | exit 0, 22 tests |
| 07:21:47 | `pnpm typecheck` | exit 0 |
| 07:22:27 | `pnpm lint` (oxlint + site ESLint) | exit 0 |
| 07:21:56 | `pnpm --dir site check` | exit 0, 0 errors, 2 warnings (existing `state_referenced_locally`) |
| 07:22:24 | `pnpm knip` | exit 0 |
| 07:22:08 | `pnpm build` | exit 0 |
| 07:22:11 | `pnpm size` | **exit 1**: io 121.70 kB (+1.70), parse 465.26 kB (+30.26) |
| 07:22:41 | `pnpm test --maxWorkers=1` | exit 0, 378 files, 3518 tests |
| 07:24:15 | editor `vitest --maxWorkers=1` | exit 0, 27 files, 948 tests |
| 07:24:59 | `check:attw`, `check:consumer`, `check:api-calibration` | exit 0 each |
| 07:25:12 | `pnpm check:api` | exit 1 (drift), reviewed, `--write`, then exit 0, hash `b76fe392…` |
| 07:25:13 | `pnpm qa:public-api-resave` | exit 0, once = twice `a4494011a7a4869c` |
| 07:26:58 | browser, own 5231, 1280×800: fill-handle drag A1:A2 → A5 | A3:A5 = 3, 4, 5; Undo clears them; 0 page errors |

Not run: perf gate, `test:browser`, `qa:office`, `qa:render`, the dotnet
conformance steps (no `dotnet` here). Not checked in real Excel.

### IO size: Major, needs a budget review (nothing below is applied)

Measured on this round's build: io 121.70 kB brotli (budget 120, +1.70) and
465.26 kB minified parse (budget 435, +30.26). Splitting the pivot model into
an opt-in subpath (about 18.7 kB minified, per the esbuild metafile) would
leave about 446.6 kB, still +11.6 kB over the parse budget, and it changes the
public API and the default load / save behaviour. So it does not solve the
budget and is not proposed as the fix. Even all of the new feature code in io
(about 29.5 kB minified: pivots 18.7, threaded comments 3.5, sparklines 3.1,
shapes +2.5, save.ts +2.6) is about the size of the overrun, and the esbuild
figures are not size-limit's own measurement, so only a budget review settles
this. Candidates, none applied: opt-in pivot subpath (-18.7 kB, breaking);
the same for threaded comments and sparklines (-6.6 kB, breaking); or raising
the io budgets with this measurement as the reason. Kept: the save-time
pivot recompute, unknown-XML passthrough, threaded comments, sparklines and
shapes.

Public API: `duplicateSheet` was not added to any subpath. The editor's Copy
Sheet keeps its own copy and renews thread and mention ids itself; the
library's internal `duplicateSheet` (unexported before and after, as in HEAD's
API baseline) now renews mention ids too. The subpath export lists in
`tests/consumer/api-baseline.json` are unchanged by this round.

Not run: a format check — the repository has no format script (no Prettier in
`package.json` or `site/package.json`).

### Incident: `prepare` ran `git submodule update` during a dependency add

- When: 2026-10-04, during `pnpm --dir site add -D eslint eslint-plugin-svelte
  svelte-eslint-parser typescript-eslint` (log `r2-add-eslint.log`). The later
  `pnpm --dir site remove svelte-eslint-parser` did not run it.
- What: the root `prepare` script, `git submodule update --init --recursive ||
  true`, ran automatically. This is a git command that may change a checkout,
  which this session had been told not to run; it was not intended.
- Effect checked afterwards: `reference/openpyxl` is at
  `11c962d17c85f78bf95be3b45efc88d1b506e0c5`, the commit HEAD records, with no
  `+`/`-` marker in `git submodule status`. The state before the add was not
  recorded, so whether it moved cannot be shown; it matches HEAD now. The
  worktree's other changes, the index (0 staged) and HEAD (`1378832`) are
  unchanged. Nothing was reverted with git.
