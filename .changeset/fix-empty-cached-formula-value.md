---
'@office-kit/xlsx': patch
---

fix: an empty cached formula value and its `t="str"` were dropped on save (#115)

`<c r="A1" t="str"><f>[1]Extern!$A$1</f><v/></c>` came back as
`<c r="A1"><f>[1]Extern!$A$1</f></c>`: the reader collapsed an empty `<v/>`
and an absent `<v>` into the same "no cached value" state, so the writer had
nothing left to emit the type from.

It matters for formulas that reference another workbook — Excel cannot
recalculate those without opening the other file, so the cached value is the
only thing it has to display. An empty `<v/>` is now read as a cached empty
string and written back with its `t="str"`; a genuinely absent `<v>` still
means "no cached value", and an empty `<v/>` on a numeric cell still carries
no number.
