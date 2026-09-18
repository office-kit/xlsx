---
"@office-kit/xlsx": patch
---

fix: a failed load now says what the bytes were, and documents which error classes mean "not an xlsx"

Uploads get validated by file extension, so a CSV saved as `.xlsx` or a
renamed legacy `.xls` reaches `loadWorkbook` and used to be told only
`openZip: archive is not a valid zip`. Where a magic number identifies the
input, the message now names it: plain text or a CSV, a UTF-8 or UTF-16
byte-order mark, a PDF, a zip with no central directory (a truncated or
partially uploaded file). An empty or short file reports its own length
against the 22-byte minimum. The OLE compound-document message no longer
asserts the file is encrypted, since a legacy `.xls` arrives in the same
container.

`loadWorkbook`, the `OpenXmlError` classes and the README now also state the
error contract: which class means "reject the upload and tell the user" as
against "retry", and that the class is stable while the message text is not.
