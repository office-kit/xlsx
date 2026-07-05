# Migrating from openpyxl to @office-kit/xlsx

`@office-kit/xlsx` is a TypeScript port of openpyxl. The data model, naming, and
semantics line up closely; the API shape is different by necessity (TypeScript
prefers free functions over methods, and the project banned classes outside
of `Error` subclasses).

This guide walks through the most common openpyxl idioms and their
@office-kit/xlsx equivalents. Pinned to @office-kit/xlsx `0.6.x`.

## Loading and saving

```python
# openpyxl
from openpyxl import load_workbook
wb = load_workbook('input.xlsx')
wb.save('output.xlsx')
```

```ts
// @office-kit/xlsx
import { loadWorkbook, saveWorkbook } from '@office-kit/xlsx/io';
import { fromFile, toFile } from '@office-kit/xlsx/node';
const wb = await loadWorkbook(fromFile('input.xlsx'));
await saveWorkbook(wb, toFile('output.xlsx'));
```

The `XlsxSource` / `XlsxSink` abstractions decouple the I/O from the workbook,
so the same `loadWorkbook` works against `fromBuffer`, `fromFile`,
`fromBlob`, `fromResponse`, `fromStream`, and `fromReadable`. `loadWorkbook`
accepts a `decompressionLimits` option (on by default) to bound the cost of
adversarial archives — leave it on when the source is untrusted.

## Workbook creation

```python
# openpyxl
from openpyxl import Workbook
wb = Workbook()       # creates a default sheet named 'Sheet'
wb.remove(wb.active)  # ...which callers usually delete immediately
ws = wb.create_sheet('Data')
```

```ts
// @office-kit/xlsx
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
const wb = createWorkbook();         // no default sheet
const ws = addWorksheet(wb, 'Data'); // call directly
```

`createWorkbook()` returns an empty workbook with **no** sheets — translating
`wb.remove(wb.active)` literally produces a no-op (or, worse, a guard that
hides a real bug elsewhere). Just call `addWorksheet` directly.

## Cells

| openpyxl                                | @office-kit/xlsx                            |
| --------------------------------------- | ----------------------------------- |
| `ws['A1'] = 42`                         | `setCellByCoord(ws, 'A1', 42)`      |
| `ws.cell(row=1, column=1, value=42)`    | `setCell(ws, 1, 1, 42)`             |
| `ws['A1'].value`                        | `ws.rows.get(1)?.get(1)?.value`     |
| `ws.iter_rows()`                        | `iterRows(ws)`                      |
| `Cell(formula='=A1+B1')`                | `setFormula(cell, 'A1+B1')`         |

Coordinates are 1-based on both sides. Cell values cover the same shapes
openpyxl does:

- numbers (`number`)
- strings (`string`, automatically deduped via the shared-strings table)
- booleans (`boolean`)
- formulas (`{ kind: 'formula', formula, t, ... }` via `setFormula`)
- errors (`{ kind: 'error', code: '#REF!' }` etc., via `makeErrorValue`)
- rich text (`{ kind: 'rich-text', runs }` via `makeRichText` / `makeTextRun`)
- dates (`Date`)
- durations (`{ kind: 'duration', ms }` via `makeDurationValue`)

## Styles

openpyxl exposes `cell.font`, `cell.fill`, etc. as direct mutable properties
that the cell points at via a styleId. @office-kit/xlsx keeps the same model — the
`Stylesheet` is a pool — but exposes it through the `@office-kit/xlsx/styles` bridge:

```ts
import {
  makeBorder,
  makeColor,
  makeFont,
  makePatternFill,
  makeSide,
  setCellBorder,
  setCellFill,
  setCellFont,
  setCellNumberFormat,
} from '@office-kit/xlsx/styles';

setCellFont(wb, cell, makeFont({ name: 'Arial', size: 14, bold: true }));
setCellFill(wb, cell, makePatternFill({
  patternType: 'solid',
  fgColor: makeColor({ rgb: 'FFFFFF00' }),
}));
setCellBorder(wb, cell, makeBorder({ left: makeSide({ style: 'thin' }) }));
setCellNumberFormat(wb, cell, '#,##0.00');
```

The pool dedups; assigning the same `Font` twice yields the same fontId.
Every style primitive has a `make*` constructor — `makeFont`,
`makeBorder({ left: makeSide(...) })`, `makePatternFill`, `makeColor`,
`makeAlignment`, etc.

## Worksheets

| openpyxl                  | @office-kit/xlsx                              |
| ------------------------- | ------------------------------------- |
| `wb.create_sheet('Data')` | `addWorksheet(wb, 'Data')`            |
| `wb.active`               | `getActiveSheet(wb)`                  |
| `wb.sheetnames`           | `sheetNames(wb)`                      |
| `wb['Data']`              | `getSheet(wb, 'Data')`                |
| `del wb['Data']`          | `removeSheet(wb, 'Data')`             |
| `ws.merged_cells.ranges`  | `getMergedCells(ws)`                  |
| `ws.merge_cells('A1:B2')` | `mergeCells(ws, 'A1:B2')`             |
| `ws.freeze_panes = 'B2'`  | `setFreezePanes(ws, 'B2')`            |

Sheet titles are validated at save time against Excel's rules (1–31 chars,
forbidden `: \ / ? * [ ]`, no leading/trailing apostrophe, reserved name
`History`, case-insensitive uniqueness). `addWorksheet` / `renameSheet`
validate eagerly; direct mutation of `ws.title` falls through to the save-time
gate.

## Streaming write (`write_only=True`)

```python
# openpyxl
wb = Workbook(write_only=True)
ws = wb.create_sheet('Data')
for row in source:
    ws.append(row)
wb.save('big.xlsx')
```

```ts
// @office-kit/xlsx
import { createWriteOnlyWorkbook } from '@office-kit/xlsx/streaming';
import { toFile } from '@office-kit/xlsx/node';

const wb = await createWriteOnlyWorkbook(toFile('big.xlsx'));
const ws = await wb.addWorksheet('Data');
ws.setColumnWidth(1, 20); // must precede the first appendRow
for (const row of source) {
  await ws.appendRow(row);
}
await ws.close();
await wb.finalize();
```

Notable difference: `setColumnWidth` must be called before the first
`appendRow` because `<cols>` is part of the worksheet header that flushes on
the first row. openpyxl is more lenient about this.

## Streaming read (`read_only=True`)

```python
# openpyxl
wb = load_workbook('big.xlsx', read_only=True)
ws = wb['Data']
for row in ws.iter_rows():
    ...
wb.close()
```

```ts
// @office-kit/xlsx
import { fromFile } from '@office-kit/xlsx/node';
import { loadWorkbookStream } from '@office-kit/xlsx/streaming';

const wb = await loadWorkbookStream(fromFile('big.xlsx'));
const ws = wb.openWorksheet('Data');
for await (const row of ws.iterRows()) {
  // row is an array of ReadOnlyCell { row, col, value, styleId }
}
await wb.close();
```

`iterRows` accepts `{ minRow, maxRow, minCol, maxCol }` for sub-sheet
iteration; the SAX path stops walking the bytes once it crosses `maxRow`.
`loadWorkbookStream` accepts the same `decompressionLimits` option as
`loadWorkbook`.

## What's preserved verbatim (no model)

@office-kit/xlsx doesn't re-implement every OOXML schema; instead, parts that
aren't worth modelling round-trip byte-for-byte through `wb.passthrough` (and
their content types via `wb.passthroughContentTypes`). The following live
there:

- `xl/vbaProject.bin`, `xl/vbaProjectSignature.bin`
- `xl/pivotCache/`, `xl/pivotTables/`
- `xl/activeX/`, `xl/embeddings/`, `xl/ctrlProps/`, `xl/oleObjects/`
- `xl/printerSettings/`, `xl/queryTables/`
- `xl/slicers/`, `xl/slicerCaches/`
- `xl/externalLinks/`, `xl/richData/`, `xl/threadedComments/`,
  `xl/persons/`
- `xl/timelineCaches/`, `xl/timelines/`
- `xl/workbookCache/` (Power Query metadata), `xl/model/` (Data Model)
- `xl/calcChain.xml`, `xl/connections.xml`, `xl/metadata.xml`,
  `xl/SheetMetadata.xml`
- `customUI/`, `customXml/`
- Control VML drawings (`xl/drawings/*.vml` excluding comment VML)

If openpyxl preserves it, `@office-kit/xlsx` does too — but typically without a typed
editing surface.

## What's not yet supported

- **Editing pivot tables** — passthrough-only round-trip. The underlying
  caches and table definitions survive load → save, but @office-kit/xlsx has no
  typed API for mutating them. Build the pivot once in Excel, then drive
  the data sheet from @office-kit/xlsx.
- **ZIP64 size overflow (per-archive > 4 GiB)** — entry-count overflow is
  patched in (archives can exceed the ZIP32 65 535-entry cap), but
  `OpenXmlNotImplementedError` is thrown if the archive byte size or
  central-directory offset crosses 4 GiB. xlsx in practice stays well under
  that.
- **Encrypted xlsx** — decrypt with [`msoffcrypto-tool`][msoffcrypto] (or
  similar) first; @office-kit/xlsx has no decryption path.

[msoffcrypto]: https://github.com/nolze/msoffcrypto-tool

## Further reading

- `README.md` — feature matrix, bundle budgets, subpath entries.
- [Cheatsheet](https://baseballyama.github.io/@office-kit/xlsx/docs/cheatsheet) —
  task → exact functions to import.
- [Recipes](https://baseballyama.github.io/@office-kit/xlsx/docs/recipes) —
  prose-style worked examples (styling, charts, validation, streaming).
- `SECURITY.md` — `decompressionLimits` defaults and the threat model when
  loading untrusted input.
