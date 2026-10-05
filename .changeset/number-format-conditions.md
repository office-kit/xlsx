---
'@office-kit/xlsx': patch
---

fix: `getCellDisplayText` ignored comparison sections (`[<1000]0;[<1000000]0.0,"K";0.0,,"M"` showed `12345` instead of `12.3K`), and showed serials 0 and 60 as 12/31/1899 and 2/28/1900 instead of Excel's 1/0/1900 and 2/29/1900.
