---
'@office-kit/xlsx': patch
---

fix: an invalid hyperlink target is now rejected instead of producing a file Excel refuses (#117)

`setHyperlink` accepted any string as `target` and the writer emitted it
verbatim into the worksheet rels, so `{ target: 'https:// www.example.com' }`
saved without complaint and Excel then reported "We found a problem with some
content in …" and dropped content on repair. One bad row poisoned an entire
export.

The rels `Target` attribute is `xsd:anyURI`. Authoring a hyperlink now throws
an `OpenXmlSchemaError` when the target is empty, contains whitespace or a
control character, or has a non-ASCII host (percent-encoding does not rescue
that one — the host has to be punycode). The check runs in `makeHyperlink`, so
`setHyperlink`, `addUrlHyperlink` and `addMailtoHyperlink` are all covered.

Reading is deliberately unaffected: a workbook whose rels already carry a
broken target still loads, so it can be inspected and repaired.
