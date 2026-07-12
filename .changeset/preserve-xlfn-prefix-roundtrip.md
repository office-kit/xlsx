---
'@office-kit/xlsx': patch
---

Preserve the `_xlfn.` future-function prefix on read so dynamic-array formulas
survive a load → save round-trip. The reader stripped `_xlfn.` / `_xlfn._xlws.`
into the model, and the writer emits formula text verbatim, so a loaded
`_xlfn.SCAN(...)` was written back as bare `SCAN(...)` — an unknown name that
Excel renders as `#NAME?`. Formula text is now kept verbatim (matching openpyxl),
fixing every future function (SCAN, BYCOL, LAMBDA, XLOOKUP, FILTER, SEQUENCE,
LET, ANCHORARRAY, …).
