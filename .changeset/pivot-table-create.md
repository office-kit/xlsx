---
'@office-kit/xlsx': minor
---

feat: create and edit PivotTables from a worksheet range

- `addPivotTable(wb, ws, { source, anchor, rows, columns, filters, values, layout, subtotals, rowGrandTotals, columnGrandTotals })` creates a PivotTable and writes its report cells. Row and column items of a date column are written with that column's date format, so they read as dates rather than serial numbers. Values aggregate with sum / count / average / max / min, report filters take one selected item, and the compact, outline and tabular layouts are supported. `refreshPivotTable` recomputes the report after the definition or source data changes, `removePivotTable` deletes it, and `getPivotTableAt`, `getPivotSourceFields`, `getPivotFieldItems`, `getPivotTableOutputRef` and `nextPivotTableName` help build UIs on top.
- `saveWorkbook` writes the pivot cache definition, cache records and pivotTable parts and wires them into `workbook.xml`, the rels and `[Content_Types].xml`. The cache is marked `refreshOnLoad`, so Excel rebuilds the report from the source when the file opens.
- `loadWorkbook` now reads a pivot back into `Worksheet.pivotTables` when the model can hold all of it — every pivot this library writes, and plain Excel pivots. Pivots that use anything the model lacks (hidden items, grouping, calculated fields, custom styles or number formats, …) are still carried through byte for byte, and `listPassthroughPivotTables` describes them by field name. A loaded pivot that is lifted into the model is regenerated on save rather than copied, so its `wb.pivotCaches` entry and its passthrough parts no longer appear after load.
- `renameSheet` now re-points PivotTable sources that name the renamed sheet.
