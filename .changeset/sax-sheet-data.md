---
"@office-kit/xlsx": patch
---

`loadWorkbook` reads a large sheet about 2.4x faster and with a third of the
transient heap. On a 50 000-row, six-column sheet (1 282 KiB archive) the load
goes from 2 024 ms to 852 ms, and peak RSS above the pre-load baseline from
803 MB to 317 MB. A 5 000-row sheet goes from 205 ms to 105 ms.

`<sheetData>` used to be read from a node tree, so every `<c>` and every `<v>`
became an object that existed only long enough to produce one cell. It is now
cut out of the part and walked with a synchronous SAX pass; the rest of the
worksheet keeps the node tree it had.

One behaviour change comes with it: malformed worksheet XML inside
`<sheetData>` is now refused. An unclosed `<row>`, a stray `</c>` or a
`<sheetData>` that is never closed used to be accepted silently, with whatever
cells the tolerant parse happened to recover; they now throw an
`OpenXmlSchemaError`. Well-formed parts, including those written with a
namespace prefix, are unaffected.
