---
'@office-kit/xlsx': patch
---

fix: workbooks with form controls (and other Office 2010+ extensions) came back with `mc:Choice Requires="a14"` pointing at a prefix that had been renamed to `ns0`, so Excel offered to repair the file. Excel's extension namespaces now keep their usual prefixes (`a14`, `x14ac`, `xr`, `c16`, …).
