---
"@office-kit/xlsx": patch
---

fix: `loadWorkbook` returned numeric character references such as `&#20219;` as literal text instead of decoding them (#131). Decimal and hexadecimal references are now decoded in text and attributes, including the inline and shared strings openpyxl writes. A reference to a character XML does not allow (for example `&#0;`) now fails the load with `OpenXmlSchemaError` instead of being kept as literal text.
