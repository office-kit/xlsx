---
"@office-kit/xlsx": patch
---

`translateFormula` and `translateRow` now raise `TranslatorError` when a shift would carry a row reference past row 1048576, instead of emitting a reference no spreadsheet can resolve. `translateFormula('=A1048576', 'A1', { rowDelta: 5 })` returned `=A1048581`. The column direction already raised at both ends, so the bottom of the grid was the one edge a translation could fall off quietly.
