---
'@office-kit/xlsx': patch
---

fix: workbook.xml lost mc:Ignorable and its namespace declarations on save (#116)

Saving a workbook whose `xl/workbook.xml` carries the markup-compatibility
block Excel writes into essentially every file stripped `mc:Ignorable` and the
namespace declarations it names off the root, renamed the child prefixes
(`x15ac:` → a generated `ns0:`, `xr:` → `x16:`), dropped the `xr2:uid` on
`<workbookView>`, and moved `<workbookPr>` behind the `mc:AlternateContent`
block. Excel reported the result as corrupt. This is the workbook-level sibling
of #105, which fixed the same rewrite on the worksheet.

The root's namespace declarations and `mc:Ignorable` are now carried through
verbatim, and every captured child is written with the prefixes the root
declares — `mc:Ignorable` names them by prefix, so the two have to agree.
`<workbookView>` keeps namespaced attributes it doesn't model (`xr2:uid`),
and the unmodeled head is emitted after `<workbookPr>`, where Excel puts it.
