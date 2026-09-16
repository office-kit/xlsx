---
'@office-kit/xlsx': minor
---

feat: formula value constructors, and the `<calcPr>` setters are now exported

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
