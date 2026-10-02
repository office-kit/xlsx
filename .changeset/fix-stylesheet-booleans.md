---
"@office-kit/xlsx": patch
---

Fix stylesheet fonts incorrectly enabling bold, italic, strike-through and other toggles when their XML value is false. Preserve explicit false font values, including differential fonts, and XF flags such as `applyFont="0"` when saving workbooks.
