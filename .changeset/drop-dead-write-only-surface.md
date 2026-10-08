---
"@office-kit/xlsx": minor
---

Remove `WriteOnlyOptions.estimatedMaxRow`. It was documented as "Reserved, currently ignored", so passing it never did anything. Delete the property from any `createWriteOnlyWorkbook` options object; nothing else changes.
