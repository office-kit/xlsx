---
'@office-kit/xlsx': patch
---

fix: styles.xml lost dxfs, tableStyles, colors and extLst on save (#113)

Everything after `<cellStyles>` in `xl/styles.xml` was dropped when a loaded
workbook was saved again. `<dxfs>` was only written when the pool held
entries, and `<tableStyles>`, `<colors>` and `<extLst>` — where Excel keeps
custom table styles, the MRU colour palette and the x14/x15 slicer and
timeline styles — were never read in the first place. Every workbook Excel
writes carries at least `<dxfs>` and `<tableStyles>`, so this hit essentially
any real file.

`Stylesheet` now carries the unmodeled tail (`stylesXmlTail`) verbatim and the
writer re-emits it after `<dxfs>`, where `CT_Stylesheet` (ECMA-376 §18.8.39)
puts it. `<dxfs count="0"/>` is now always written, matching Excel — a
conditional-formatting rule's `dxfId` is an index into that list.
