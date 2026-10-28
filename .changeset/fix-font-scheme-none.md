---
"@office-kit/xlsx": minor
---

fix: `loadWorkbook` threw `OpenXmlSchemaError: expected one of [major, minor], got "none"` on a stylesheet font with `<scheme val="none"/>`, which Excel writes for fonts that opt out of the theme (#306). `FontScheme` now includes `'none'`, and the value is kept on save in cell, differential and rich-text fonts alike. A rich-text run with a scheme outside `none` / `major` / `minor` now throws `OpenXmlSchemaError` on load, as the same value in `styles.xml` already did.

Breaking (types only): code that switches exhaustively over `FontScheme`, or builds a `Record<FontScheme, …>`, needs a `'none'` entry.
