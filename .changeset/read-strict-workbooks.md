---
"@office-kit/xlsx": minor
---

Read common ISO 29500 Strict XLSX files through `loadWorkbook` and `loadWorkbookStream`, including mixed-namespace packages, styles, shared strings, formulas, themes and supported charts. ISO date cells and cached dates become Excel serial numbers. Saving produces Transitional XLSX. Strict content that cannot be converted safely raises an explicit error; see the README for supported date ranges and conversion limits.

Fix buffered ZIP compression of sparse binary parts, such as printer settings, so saving and reloading preserves their bytes.
