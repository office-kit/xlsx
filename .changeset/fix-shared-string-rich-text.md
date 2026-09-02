---
'@office-kit/xlsx': minor
---

fix: rich text was flattened to plain text on save (#114)

A shared string built from `<r>` runs came back as a single plain `<t>`: the
loader collapsed rich-text `<si>` entries into their concatenated text before
the worksheet reader saw them, so a cell with a bold first half round-tripped
uniformly unformatted. Inline strings (`<c t="inlineStr"><is>`) lost their runs
the same way on read.

Rich-text `<si>` and `<is>` bodies now go through one CT_Rst parser and become
`{ kind: 'rich-text' }` cell values, so per-run fonts survive. Cells holding
rich text are written into `xl/sharedStrings.xml` as `<si><r>…</r></si>` —
where Excel itself stores them — instead of being inlined per cell, so a
formatted string repeated across cells costs one entry rather than one copy
each. `<r>` runs with no `<rPr>` stay separate runs, since that is how Excel
writes the unformatted half of a rich string.

**Behaviour change:** a cell whose shared string is rich text now reads back
as `{ kind: 'rich-text', runs }` where it used to read back as the concatenated
plain string, and rich-text cells serialise into `xl/sharedStrings.xml` rather
than as `t="inlineStr"`. Code that assumed `getCell(...).value` was a string
for those cells needs to handle the union.
