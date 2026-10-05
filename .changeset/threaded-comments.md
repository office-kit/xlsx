---
'@office-kit/xlsx': minor
---

feat: read and write Excel threaded comments

- `Worksheet.threadedComments` models Excel 365 threaded comments (`xl/threadedComments/`): cell, author, creation time, text, replies via `parentId`, the resolved (`done`) state and @mentions. `makeThreadedComment` from `@office-kit/xlsx/worksheet` builds one with a fresh id.
- `Workbook.persons` models the comment authors (`xl/persons/person.xml`); `makePerson` from `@office-kit/xlsx/workbook` builds one.
- `Workbook.persons` and `Worksheet.threadedComments` are optional, so `Workbook` / `Worksheet` object literals written for earlier versions still compile. `createWorkbook` and `addWorksheet` set them to `[]`; when they are absent the workbook has no threaded comments.
- On save, each thread also gets the legacy "[Threaded comment]" note and VML shape Excel writes, so older readers still show it and Excel opens the file without repair. On load, those placeholders are recognised and no longer show up in `Worksheet.legacyComments`.
- `Workbook.authors` is deprecated: it was never read or written, and threaded-comment authors are now in `Workbook.persons`. It is still there, so existing code keeps compiling.
- Threaded-comment and person parts are now read into `Worksheet.threadedComments` and `Workbook.persons` instead of `Workbook.passthrough`. They are still written back on save; code that looked them up in `passthrough` no longer finds them there.
- `duplicateSheet` gives the copied threads and their @mentions new ids, since Excel rejects ids repeated across sheets.
