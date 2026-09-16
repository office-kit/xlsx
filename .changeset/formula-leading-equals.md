---
'@office-kit/xlsx': patch
---

fix: formula text starting with `=` produced a workbook Excel calls damaged

OOXML stores formula text without the leading `=` (ECMA-376 §18.3.1.40), but no
setter stripped it. `setFormula(cell, '=SUM(A1:A3)')` emitted
`<f>=SUM(A1:A3)</f>`, which Excel reports as an unreadable-content error on open.
The natural spelling was the broken one, and nothing in the types or the docs said
so.

`setFormula`, `setArrayFormula`, `setSharedFormula` and `setDataTableFormula` now
strip a single leading `=`, so `'=SUM(A1:A3)'` and `'SUM(A1:A3)'` are
interchangeable.
