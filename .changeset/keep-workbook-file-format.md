---
'@office-kit/xlsx': minor
---

fix: a `.xlsm` with no macros, or a template, is saved in its own format again. Before, any workbook without a `vbaProject` was saved with the plain `.xlsx` workbook content type, and Excel refused to open the result under its `.xlsm` name ("file format or file extension is not valid"). The new `Workbook.fileFormat` (`'xlsx' | 'xlsm' | 'xltx' | 'xltm'`) is set by `loadWorkbook` and decides the content type on save; a `vbaProject` still makes the file macro-enabled.
