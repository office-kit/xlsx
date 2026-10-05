---
'@office-kit/xlsx': patch
---

fix: `addDefinedName` and `addTable` / `addExcelTable` accepted names Excel rejects (`T1`, `R2C3`, names with spaces), producing a file Excel offered to repair. They now throw `OpenXmlSchemaError`.
