---
'@office-kit/xlsx': patch
---

fix: `getCellDisplayText` keeps the decimal point when a format shows no fraction digits, as Excel does (`0.` shows `2.`, `#.##` shows `5.` for 5)
