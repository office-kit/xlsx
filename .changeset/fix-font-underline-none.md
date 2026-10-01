---
"@office-kit/xlsx": patch
---

Fix loading workbooks whose stylesheet fonts contain `<u val="none"/>`. Preserve explicit no-underline values in cell and differential fonts when saving, and render them without an underline in `fontToCss`.
