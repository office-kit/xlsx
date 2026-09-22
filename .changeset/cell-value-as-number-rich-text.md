---
"@office-kit/xlsx": patch
---

`cellValueAsNumber` now reads a rich-text cell, which its documentation has always described ("rich-text concats then parses"). The branch was missing, so a number typed into a cell carrying per-run formatting returned `undefined` while `cellValueAsString` on the same cell returned `"42"`.
