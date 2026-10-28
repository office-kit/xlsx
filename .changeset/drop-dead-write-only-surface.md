---
"@office-kit/xlsx": minor
---

Breaking (types only): remove `WriteOnlyOptions.estimatedMaxRow`. It was documented as "Reserved, currently ignored", so passing it never did anything. TypeScript callers that pass it get an excess-property error; delete the property from the `createWriteOnlyWorkbook` options object. Runtime behaviour is unchanged.
