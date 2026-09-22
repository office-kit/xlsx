---
"@office-kit/xlsx": patch
---

Documentation corrections where a docblock described behaviour the code does not have. No runtime change.

- `getMergedCells`, `listHyperlinks`, `listDataValidations`, `listTables`, `listComments`, `listCustomProperties` and the stylesheet pool accessors (`listFonts`, `listFills`, `listBorders`, `listCellXfs`, `listCellStyleXfs`) were documented as returning a "snapshot". They return a read-only view of the live array, so later mutations are visible through it, and a `removeAll*` replaces the array outright and leaves an earlier return value stale. Copy the result before mutating the sheet while reading it.
- `shiftRange` said its result was "clamped to the OOXML grid". It throws, as `shiftRangeStr` already documented.
- `listDefinedNames` offered `{ scope: undefined }` for workbook-scope names. That spelling does not compile under `exactOptionalPropertyTypes` and would list every name if it did; `{ scope: 'workbook' }` is the one that narrows.
