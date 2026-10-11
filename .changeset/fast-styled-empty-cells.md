---
'@office-kit/xlsx': patch
---

perf: `loadWorkbookStream`'s `iterRows` / `iterValues` no longer slow down on wide sheets full of styled empty cells (`<c r="X1" s="n"/>`), the shape many business exports have. The row walk now drives the XML parser directly instead of handing each SAX event through an async iterator, which makes it several times faster on such sheets; cell references are also parsed without a regex, which speeds up `loadWorkbook` too (#302)
