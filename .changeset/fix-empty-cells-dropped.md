---
'@office-kit/xlsx': patch
---

fix: cells without a value were dropped on save, leaving defined names dangling (#111)

`<c r="B1"/>` and `<c r="C1" s="0"/>` disappeared from the saved sheet: the
writer skipped any cell whose value was `null` unless it also carried a
non-default style. Dropping a genuinely unreferenced empty cell is harmless,
but an empty cell can still be the target of a `definedName`, a form control's
`fmlaLink`, or a drawing or comment anchor — and those were left pointing at
nothing.

A valueless cell now always emits. `deleteCell` — not `setCell(…, null)` — is
how a caller says the cell is gone; cells that were never in the model are
still not written.
