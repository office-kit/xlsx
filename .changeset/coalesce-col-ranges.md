---
'@office-kit/xlsx': patch
---

perf(worksheet): coalesce adjacent columns with identical properties into a single `<col>` range when serializing `<cols>`, on both the modelled and the write-only writer. A uniform band (for example 2,451 columns at `width="0.3"`) now emits one `<col min max>` element instead of thousands, shrinking the worksheet part and speeding up reloads. Entries are merged only at serialization time — the `columnDimensions` map, the public API, and files with differing properties or gaps are unchanged, and output stays functionally equivalent in Excel.
