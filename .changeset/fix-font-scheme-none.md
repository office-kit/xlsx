---
"@office-kit/xlsx": minor
---

fix: `loadWorkbook` threw `OpenXmlSchemaError: expected one of [major, minor], got "none"` on a stylesheet font with `<scheme val="none"/>`, which Excel writes for fonts that opt out of the theme (#306). `FontScheme` now includes `'none'`, and the value is kept on save, in cell, differential and rich-text fonts alike.
