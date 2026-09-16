---
'@office-kit/xlsx': minor
---

feat: `makeFormula` builds a formula as a cell value, and `setFullCalcOnLoad` is exported

Placing a formula took two steps: reach or create a cell, then mutate it with
`setFormula`. `makeFormula(text, { cachedValue })` returns the `CellValue`, so
`setCell(ws, row, col, makeFormula('SUM(B5:I5)'))` is the whole write, and it
composes with the `styleId` argument `setCell` already accepted. `setFormula`
remains the mutate-a-cell-you-hold form.

`setFullCalcOnLoad` existed but was never exported from `@office-kit/xlsx/workbook`,
which left the advice "cache the value or ask Excel to recalculate" unactionable. It
is now public. Excel, LibreOffice and Google Sheets compute an uncached formula on
open, but viewers that never calculate (Quick Look, Outlook and SharePoint previews,
most thumbnailers) render those cells empty, so a generated workbook with formulas
wants both a cached value where the producer can compute one and
`setFullCalcOnLoad(wb, true)` for the rest.
