---
'@office-kit/xlsx': minor
---

fix: formula text starting with `=` produced a workbook Excel calls damaged

OOXML stores formula text without the leading `=` (ECMA-376 §18.3.1.40), but no
setter stripped it. `setFormula(cell, '=SUM(A1:A3)')` emitted
`<f>=SUM(A1:A3)</f>`, which Excel reports as an unreadable-content error on open.
The natural spelling was the broken one, and nothing in the types or the docs said
so.

Formula text is now normalised (leading `=`, plus the whitespace around it) both
where it enters the model and where it is serialised, so a hand-built
`FormulaValue` passed straight to `setCell` cannot produce a damaged file either.
The other elements that carry OOXML formula text get the same treatment:
`<formula1>` / `<formula2>` on a data validation, `<formula>` on a
conditional-formatting rule, and a defined name's value.

Two behaviour changes to watch for on upgrade:

- Loading a workbook whose `<f>` carried a leading `=` (including any file this
  library wrote before this release) now gives `getFormulaText(cell)` as
  `'SUM(A1:A3)'` where it returned `'=SUM(A1:A3)'`, and re-saving writes the
  normalised text.
- `makeDataValidation`, `makeCfRule`, `addDefinedName` and the builders over them
  (`addListValidation`, `addCustomValidation`, `addFormulaRule`, …) store the
  normalised text, so reading `dv.formula1` back returns it without the `=`.

`makeFormula` and `makeArrayFormula` now throw `OpenXmlSchemaError` when the text
normalises to nothing (`''` or `'='`). That case used to emit an empty `<f/>`,
which Excel rejects as well, and loading a file that already contains one now
throws rather than carrying a formula cell with no expression.
