---
'@office-kit/xlsx': patch
---

fix: `saveWorkbook` / `workbookToBytes` now throw `OpenXmlSchemaError` when a conditional-formatting rule's `dxfId` points past the stylesheet's differential formats. Such a file was written before, and Excel silently refused to open it. Add the format with `addDxf(wb.styles, ...)` first.
