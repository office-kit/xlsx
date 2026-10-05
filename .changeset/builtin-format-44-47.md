---
'@office-kit/xlsx': patch
---

fix: built-in number formats 44 (Accounting) and 47 now carry the codes Excel renders. `getCellDisplayText` showed Accounting cells as a bare number (`1234.5` instead of ` $1,234.50 `), and id 47 as `0101.4` instead of `01:01.4`.
