// /llms.txt — self-contained reference for AI assistants.
//
// The llmstxt.org proposal suggests a short index. We extend that: the file
// keeps the index shape (H1 title, blockquote summary, link sections at the
// end) but inlines enough subpath-by-subpath API guidance that an agent
// can use every feature of @office-kit/xlsx without having to fetch additional
// pages. For the long-form prose docs an agent can still follow the links
// at the bottom, append `.md` to any docs URL for raw Markdown, or fetch
// `/llms-full.txt` for the whole site concatenated into a single document.

import { docSections } from '$lib/docs-nav';
import type { RequestHandler } from './$types';

export const prerender = true;

const HEADER = `# @office-kit/xlsx

> Read and write Excel \`.xlsx\` workbooks from Node 22+ and modern browsers, with no Python or native runtime dependencies. Includes a streaming writer (10M rows in fixed memory) and a SAX-based row iterator for huge sheets.

This file is written for AI assistants. It is self-contained: an agent can use
every documented feature of @office-kit/xlsx from this page alone. The link index at
the bottom points at canonical docs — append \`.md\` to any docs URL for raw
Markdown (e.g. \`/docs/install.md\`), or fetch \`/llms-full.txt\` for every
page concatenated into a single document.

## Runtime

- Node \`>=22\` (uses built-in Web Streams, Blob, fetch).
- Modern browsers, Bun, Deno, Cloudflare Workers, edge runtimes — anywhere
  with \`fetch\` + Web Streams.
- ESM-only. \`"sideEffects": false\` — fully tree-shakable.
- No runtime dependencies on Python or native binaries. Internally uses
  \`fflate\` (deflate/inflate), \`saxes\` (SAX XML), \`fast-xml-parser\`.

## Install

\`\`\`sh
pnpm add @office-kit/xlsx
# or: npm install @office-kit/xlsx / bun add @office-kit/xlsx
\`\`\`

TypeScript: \`tsconfig.json\` should use \`"moduleResolution": "bundler"\`
(or \`"node16"\` / \`"nodenext"\`) so subpath entries resolve.

## One way to do one thing

@office-kit/xlsx has **no root barrel export**. \`import "@office-kit/xlsx"\` is intentionally
not supported. Every export lives behind a section subpath, and each public
function has exactly one canonical import path — once you know the function
name, you know where to import it from.

| Subpath                | What lives here                                                                                            |
|------------------------|------------------------------------------------------------------------------------------------------------|
| \`@office-kit/xlsx/io\`          | \`loadWorkbook\`, \`saveWorkbook\`, \`workbookToBytes\`; browser-safe \`fromArrayBuffer\` / \`fromBlob\` / \`fromResponse\` / \`fromStream\` / \`toArrayBuffer\` / \`toBlob\`; types \`XlsxSource\`, \`XlsxSink\`, \`LoadOptions\`, \`SaveOptions\` |
| \`@office-kit/xlsx/node\`        | Node-fs bridges: \`fromFile\`, \`fromFileSync\`, \`fromBuffer\`, \`fromReadable\`, \`toFile\`, \`toWritable\`, \`toBuffer\`, \`workbookToBuffer\` |
| \`@office-kit/xlsx/streaming\`   | \`loadWorkbookStream\`, \`createWriteOnlyWorkbook\` + their type surface (\`ReadOnlyWorkbook\`, \`ReadOnlyWorksheet\`, \`ReadOnlyCell\`, \`IterRowsOptions\`, \`LoadWorkbookStreamOptions\`, \`WriteOnlyWorkbook\`, \`WriteOnlyWorksheet\`, \`WriteOnlyOptions\`, \`WriteOnlyRowItem\`, \`WriteOnlyStyle\`) |
| \`@office-kit/xlsx/workbook\`    | \`createWorkbook\`, \`addWorksheet\`, \`addChartsheet\`, \`addDefinedName\`, \`makeWorkbookProtection\`, \`makeWorkbookView\` / \`makeCustomWorkbookView\`, \`makeCalcProperties\`, \`makeWorkbookProperties\`, \`makeFileVersion\`, \`makeFileSharing\`, \`makeFileRecoveryProperties\`, \`makeSmartTagProperties\` / \`makeSmartTagType\`, \`makeFunctionGroup\` / \`makeFunctionGroups\` |
| \`@office-kit/xlsx/worksheet\`   | \`setCell\`, \`setCellByCoord\`, \`getCell\`, \`getCellByCoord\`, \`getCellsInRow\`, \`getCellsInColumn\`, \`getCellsInRange\`, \`appendRow\`, \`appendRows\`, \`iterRows\`, \`iterValues\`, \`getMaxRow\`, \`getMaxCol\`, \`mergeCells\` / \`unmergeCells\` / \`unmergeCellsAt\`, \`makeSheetView\` / \`makeFreezePane\` / \`freezePaneRef\`, \`setHyperlink\` / \`makeHyperlink\` / \`getCellHyperlink\`, \`addExcelTable\` / \`makeTableDefinition\` / \`makeTableColumn\`, \`addAutoFilter\` / \`addAutoFilterColumn\` / \`makeAutoFilter\` / \`makeFilterColumn\`, \`makeDataValidation\` / \`addDataValidation\`, \`makeCfRule\` / \`makeConditionalFormatting\` / \`addConditionalFormatting\`, \`makeLegacyComment\` / \`getCellComment\`, \`makeColumnDimension\` / \`makeRowDimension\`, \`makeSheetProtection\` / \`makeProtectedRange\`, \`makeSortState\` / \`makeSortCondition\`, \`makeFormControl\` / \`makeOleObject\`, \`makeCustomSheetView\`, \`makeCellWatch\` / \`makeIgnoredError\`, \`makeSheetProperties\` |
| \`@office-kit/xlsx/cell\`        | \`makeCell\`, \`getCoordinate\`, \`setCellValue\`, \`bindValue\`, \`setFormula\` / \`setSharedFormula\` / \`setArrayFormula\` / \`setDataTableFormula\`, \`makeErrorValue\` / \`makeDurationValue\`, \`makeRichText\` / \`makeTextRun\` / \`richTextToString\`, predicates (\`isFormulaCell\`, \`isRichTextCell\`, \`isEmptyCell\`, \`isMergedCell\`, \`isErrorCell\`, \`isFormulaValue\`, \`isRichTextValue\`, \`isErrorValue\`, \`isDurationValue\`), and value coercers (\`cellValueAsString\`, \`cellValueAsBoolean\`, \`cellValueAsDate\`, \`cellValueAsNumber\`, \`cellValueAsPrimitive\`); types \`Cell\`, \`CellValue\`, \`MergedCell\`, \`FormulaValue\`, \`RichText\`, \`TextRun\`, \`InlineFont\`, \`ExcelErrorCode\` |
| \`@office-kit/xlsx/styles\`      | Per-cell: \`setBold\`, \`setFontSize\`, \`setFontName\`, \`setFontColor\`, \`setCellFont\`, \`setCellFill\`, \`setCellBorder\` / \`setCellBorderAll\`, \`setCellAlignment\` / \`centerCell\`, \`setCellProtection\`, \`setCellNumberFormat\`, \`setCellStyle\`, \`setCellBackgroundColor\`, \`setCellAsCurrency\`, \`setCellAsPercent\`, \`setCellAsDate\`, \`setCellAsNumber\`. Range-wide \`setRange*\` variants for the same axes. Built-in format constants \`FORMAT_GENERAL\`, \`FORMAT_TEXT\`, \`FORMAT_NUMBER\`, \`FORMAT_NUMBER_00\`, \`FORMAT_PERCENTAGE\`, \`FORMAT_PERCENTAGE_00\`, \`FORMAT_DATE_DATETIME\`, etc. Named-style + DXF (differential format) APIs |
| \`@office-kit/xlsx/chart\`       | Legacy \`c:\` chart kinds — \`makeBarChart\`, \`makeLineChart\`, \`makeAreaChart\`, \`makePieChart\`, \`makeDoughnutChart\`, \`makeScatterChart\`, \`makeRadarChart\`, \`makeBubbleChart\`, \`makeStockChart\`, \`makeSurfaceChart\`, \`makeOfPieChart\`. Series builder \`makeBarSeries\`. Top-level wrapper \`makeChartSpace\`. Modern \`cx:\` chartex kinds — \`makeSunburstChart\`, \`makeTreemapChart\`, \`makeWaterfallChart\`, \`makeHistogramChart\`, \`makeParetoChart\`, \`makeFunnelChart\`, \`makeBoxWhiskerChart\`, \`makeRegionMapChart\` |
| \`@office-kit/xlsx/chartsheet\`  | Standalone chartsheets (\`addChartsheet\` lives on \`@office-kit/xlsx/workbook\`; chartsheet shape and helpers live here) |
| \`@office-kit/xlsx/drawing\`     | \`addImageAt\`, \`addChartAt\`, \`loadImage\`, \`detectImageFormat\`, \`detectImageDimensions\`; \`makeAbsoluteAnchor\`, \`makeOneCellAnchor\`, \`makeTwoCellAnchor\`, \`anchorMarkerFromCellRef\`; \`makeDrawing\`, \`makePictureDrawingItem\`, \`makeChartDrawingItem\`, \`listImagesOnSheet\`, \`listChartsOnSheet\`, \`removeAllImages\`, \`removeAllCharts\`, \`removeAllDrawingItems\`; DML preset geometry / effect helpers |
| \`@office-kit/xlsx/utils\`       | A1 ↔ row/col conversion (\`coordinateFromString\`, \`coordinateToTuple\`, \`tupleToCoordinate\`, \`columnLetterFromIndex\`, \`columnIndexFromLetter\`, \`rangeBoundaries\`, \`boundariesToRangeString\`, \`parseSheetRange\`, \`formatSheetQualifiedRef\`, validators \`isValidCellRef\` / \`isValidRangeRef\` / \`isValidColumnLetter\` / \`isValidColumnNumber\` / \`isValidRowNumber\`); Excel date helpers (\`excelToDate\`, \`dateToExcel\`, \`excelToDuration\`, \`durationToExcel\`, \`toIso8601\`, \`fromIso8601\`); EMU/pixel/point unit conversion; \`escapeCellString\` / \`unescapeCellString\`; \`inferCellType\` + \`ERROR_CODES\`; error hierarchy \`OpenXmlError\`, \`OpenXmlIoError\`, \`OpenXmlSchemaError\`, \`OpenXmlDecompressionBombError\`, \`OpenXmlInvalidWorkbookError\`, \`OpenXmlNotImplementedError\` |
| \`@office-kit/xlsx/packaging\`   | Low-level OPC (Open Packaging Conventions) parts — escape hatch only |
| \`@office-kit/xlsx/xml\`         | Hardened XML reader/writer — escape hatch only |
| \`@office-kit/xlsx/zip\`         | ZIP reader/writer with decompression-bomb defense — escape hatch only |
| \`@office-kit/xlsx/schema\`      | OOXML namespace and ECMA-376 type constants — escape hatch only |

The four "escape hatch" subpaths (\`packaging\`, \`xml\`, \`zip\`, \`schema\`)
are only needed when extending the library. Anything an end-user wants to do
should be reachable through the higher-level subpaths above them.

Bundle budgets, min + brotli:

- \`@office-kit/xlsx/io\` plus the rest of the high-level surface: ≤ 120 KB (currently ~78 KB)
- \`@office-kit/xlsx/streaming\`: ≤ 80 KB (currently ~47 KB)

## Coordinates and cell values

- Cell coordinates are **1-indexed**. \`row=1, col=1\` corresponds to \`A1\`.
  Matches the openpyxl convention.
- \`CellValue\` is a tagged union:
  \`\`\`ts
  type CellValue =
    | number
    | string
    | boolean
    | Date
    | { kind: 'duration'; ms: number }
    | { kind: 'error'; code: ExcelErrorCode }   // '#DIV/0!' | '#N/A' | '#NAME?' | '#NULL!' | '#NUM!' | '#REF!' | '#VALUE!' | '#GETTING_DATA' | '#SPILL!' | '#CALC!'
    | { kind: 'rich-text'; runs: RichText }
    | FormulaValue                              // { kind: 'formula'; formula: string; cachedValue?: number | string | boolean; ... }
    | null;                                     // empty cell
  \`\`\`
- \`setCell(sheet, row, col, value)\` returns the \`Cell\` it created, so you
  can chain it into \`setFormula\`, \`setBold\`, \`setCellNumberFormat\`, etc.
- For A1-style addressing, prefer \`setCellByCoord(ws, 'A1', value)\` /
  \`getCellByCoord(ws, 'A1')\`.
- To convert between A1 and row/col, use the helpers in \`@office-kit/xlsx/utils\`
  (\`coordinateFromString\`, \`coordinateToTuple\`, \`columnLetterFromIndex\`,
  \`columnIndexFromLetter\`).

## Loading a workbook

\`loadWorkbook\` accepts an \`XlsxSource\` from \`@office-kit/xlsx/io\` or
\`@office-kit/xlsx/node\`. Pick the helper that matches your byte source:

\`\`\`ts
import { loadWorkbook } from '@office-kit/xlsx/io';

// Node, from file path (streamed via fs.ReadStream)
import { fromFile } from '@office-kit/xlsx/node';
const wb = await loadWorkbook(fromFile('in.xlsx'));

// Node, from Buffer / Uint8Array
import { fromBuffer } from '@office-kit/xlsx/node';
const wb = await loadWorkbook(fromBuffer(buf));

// Node, from a Readable
import { fromReadable } from '@office-kit/xlsx/node';
const wb = await loadWorkbook(fromReadable(stream));

// Browser, from fetch (streamed — parsing starts before download finishes)
import { fromResponse } from '@office-kit/xlsx/io';
const wb = await loadWorkbook(fromResponse(await fetch('/sheet.xlsx')));

// Browser, from <input type="file"> / drag-and-drop
import { fromBlob } from '@office-kit/xlsx/io';
const wb = await loadWorkbook(fromBlob(file));

// Anywhere, from ArrayBuffer / ReadableStream<Uint8Array>
import { fromArrayBuffer, fromStream } from '@office-kit/xlsx/io';
const wb = await loadWorkbook(fromArrayBuffer(ab));
const wb2 = await loadWorkbook(fromStream(readable));
\`\`\`

\`wb.sheets\` is an array of *discriminated unions* — \`SheetRef\` objects
with a \`kind\` of \`'worksheet'\` or \`'chartsheet'\`. Narrow on \`kind\`
before reading \`.sheet\`:

\`\`\`ts
for (const ref of wb.sheets) {
  if (ref.kind !== 'worksheet') continue;
  const ws = ref.sheet;          // Worksheet
  console.log(ref.name);         // sheet name (also ws.title)
  // ...
}
\`\`\`

Iterate cells with the canonical helpers from \`@office-kit/xlsx/worksheet\`:

\`\`\`ts
import { iterRows, iterValues } from '@office-kit/xlsx/worksheet';

for (const row of iterRows(ws)) {        // Cell[] per row
  for (const cell of row) console.log(cell.row, cell.col, cell.value);
}
const grid = [...iterValues(ws)];        // CellValue[][]
\`\`\`

\`loadWorkbook\` options (\`LoadOptions\`) cover decompression limits, shared
string interning, calc-property handling, and which parts to skip. The
defaults are safe; only override them when you have a reason.

## Saving a workbook

\`\`\`ts
import { saveWorkbook, workbookToBytes, toBlob, toArrayBuffer } from '@office-kit/xlsx/io';
import { toFile, toWritable, workbookToBuffer } from '@office-kit/xlsx/node';

// Node, to file path (streamed via fs.WriteStream)
await saveWorkbook(wb, toFile('out.xlsx'));

// Node, to a Writable
await saveWorkbook(wb, toWritable(stream));

// Node, to a Buffer
const buf: Buffer = await workbookToBuffer(wb);

// Anywhere, to Uint8Array / Blob / ArrayBuffer
const bytes = await workbookToBytes(wb);
const blob  = await toBlob(wb);
const ab    = await toArrayBuffer(wb);

// Triggering a browser download:
const url = URL.createObjectURL(await toBlob(wb));
const a = document.createElement('a');
a.href = url; a.download = 'report.xlsx'; a.click();
URL.revokeObjectURL(url);
\`\`\`

## Building a workbook from scratch

\`\`\`ts
import { workbookToBytes } from '@office-kit/xlsx/io';
import { createWorkbook, addWorksheet, addDefinedName } from '@office-kit/xlsx/workbook';
import { setCell, mergeCells, makeFreezePane, makeSheetView } from '@office-kit/xlsx/worksheet';
import { setFormula } from '@office-kit/xlsx/cell';
import { setBold, centerCell } from '@office-kit/xlsx/styles';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Sales');

// Header
const title = setCell(ws, 1, 1, 'Q2 sales');
setBold(wb, title);
centerCell(wb, title);
mergeCells(ws, 'A1:C1');

// Body
setCell(ws, 2, 1, 'Region');
setCell(ws, 2, 2, 'Revenue');
setCell(ws, 3, 1, 'NA');
setCell(ws, 3, 2, 12_400);
setCell(ws, 4, 1, 'EU');
setCell(ws, 4, 2, 9_800);

setFormula(setCell(ws, 5, 2, null), 'SUM(B3:B4)', { cachedValue: 22_200 });

// Freeze header
ws.views.push(makeSheetView({ pane: makeFreezePane('A3') }));

// Cross-sheet reference
addDefinedName(wb, { name: 'TotalRevenue', value: 'Sales!$B$5' });

const bytes = await workbookToBytes(wb);
\`\`\`

## Editing an existing workbook

\`\`\`ts
import { loadWorkbook, saveWorkbook } from '@office-kit/xlsx/io';
import { fromFile, toFile } from '@office-kit/xlsx/node';
import { setCell } from '@office-kit/xlsx/worksheet';

const wb = await loadWorkbook(fromFile('in.xlsx'));
const ref = wb.sheets[0];
if (ref?.kind === 'worksheet') setCell(ref.sheet, 1, 1, 'updated');
await saveWorkbook(wb, toFile('out.xlsx'));
\`\`\`

## Streaming read (huge sheets, fixed memory)

\`\`\`ts
import { fromFile } from '@office-kit/xlsx/node';
import { loadWorkbookStream } from '@office-kit/xlsx/streaming';

const wb = await loadWorkbookStream(fromFile('big.xlsx'));
const name = wb.sheetNames[0] ?? '';     // workbook part is parsed eagerly
const sheet = wb.openWorksheet(name);

// Walk every row.
for await (const row of sheet.iterRows()) {
  console.log(row.map((c) => c.value));
}

// Or bound the walk. minRow is tag-scanned — a 1M-row skip does not parse
// the skipped rows, so this is fast even on multi-million-row sheets.
for await (const row of sheet.iterRows({ minRow: 1_000_000, maxRow: 1_000_100, minCol: 1, maxCol: 8 })) {
  // ...
}
await wb.close();
\`\`\`

Each yielded row is a \`ReadOnlyCell[]\`. \`cell.value\` is
\`string | number | boolean | null\`; \`cell.coordinate\` is the A1 string.

## Streaming write (huge sheets, fixed memory)

\`\`\`ts
import { toFile } from '@office-kit/xlsx/node';
import { createWriteOnlyWorkbook } from '@office-kit/xlsx/streaming';

const wb = await createWriteOnlyWorkbook(toFile('big.xlsx'));
const ws = await wb.addWorksheet('Data');

// Column widths MUST be set before the first appendRow — once any row is
// written, the worksheet's <cols> block is locked.
ws.setColumnWidth(1, 24);
ws.setColumnWidth(2, 12);

for (let r = 0; r < 10_000_000; r++) {
  await ws.appendRow([r, \`row-\${r}\`, r * Math.PI]);
}

await ws.close();      // REQUIRED — finalizes this sheet's <sheetData>.
await wb.finalize();   // REQUIRED — writes the OPC central directory.
\`\`\`

Streaming-write constraints:

- Column widths and column-level styling must be set **before** the first
  \`appendRow\`. Once any row is written, \`<cols>\` is locked.
- No random-access edits — once a row is appended, you cannot go back and
  change a cell.
- Both \`ws.close()\` and \`wb.finalize()\` are required; skipping either
  leaves a truncated zip.
- Charts, drawings, and Excel tables are not supported in the streaming
  writer. If you need them, build the workbook in memory.

For a 10M-row × 3-column workbook on commodity hardware: ~30s, ~75 MB peak
heap, ~110 MB on disk.

## Styling cells

Per-cell helpers (apply to a \`Cell\` returned by \`setCell\` / \`getCell\`):

\`\`\`ts
import {
  setBold, setFontSize, setFontName, setFontColor,
  setCellBackgroundColor, centerCell, setCellBorderAll,
  setCellAsCurrency, setCellAsPercent, setCellAsDate, setCellAsNumber,
  setCellNumberFormat, setCellStyle,
  FORMAT_DATE_DATETIME, FORMAT_PERCENTAGE, FORMAT_NUMBER_00,
} from '@office-kit/xlsx/styles';

const c = setCell(ws, 1, 1, 'Header');
setBold(wb, c);
setFontSize(wb, c, 14);
setFontName(wb, c, 'Inter');
setFontColor(wb, c, 'FFFFFFFF');
setCellBackgroundColor(wb, c, 'FF3366CC');   // AARRGGBB (leading FF = opaque)
centerCell(wb, c);
setCellBorderAll(wb, c, { style: 'thin' });

// Number formats
setCellAsCurrency(wb, setCell(ws, 2, 1, 1234.5), { symbol: '$' });
setCellAsPercent(wb, setCell(ws, 2, 2, 0.125), 1);
setCellAsDate(wb, setCell(ws, 2, 3, new Date()), 'yyyy-mm-dd');
setCellAsNumber(wb, setCell(ws, 2, 4, 42_000), 0);
setCellNumberFormat(wb, setCell(ws, 2, 5, new Date()), FORMAT_DATE_DATETIME);

// Atomic apply of multiple aspects in one styleId update:
setCellStyle(wb, c, {
  font: { bold: true, size: 12, color: { rgb: 'FFFFFFFF' } },
  fill: { patternType: 'solid', fgColor: { rgb: 'FF3366CC' } },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: { top: { style: 'thin' }, bottom: { style: 'thin' } },
  numberFormat: FORMAT_NUMBER_00,
});
\`\`\`

For range-wide styling, use the matching \`setRange*\` helpers in
\`@office-kit/xlsx/styles\` (e.g. \`setRangeFont\`, \`setRangeAlignment\`,
\`setRangeBorderBox\`, \`setRangeNumberFormat\`).

Colors throughout the library are hex \`AARRGGBB\` strings. Leading \`FF\`
means opaque; \`80\` would be 50% alpha.

## Formulas

\`\`\`ts
import { setFormula, setSharedFormula, setArrayFormula } from '@office-kit/xlsx/cell';
import { setCell } from '@office-kit/xlsx/worksheet';

// Plain formula with optional cached value
setFormula(setCell(ws, 3, 1, null), 'SUM(A1:A2)', { cachedValue: 42 });

// Array formula (CSE — host range describes the spill area)
setArrayFormula(setCell(ws, 1, 3, null), 'TRANSPOSE(A1:A3)', { ref: 'C1:C3' });

// Shared formula (Excel optimization for runs of similar formulas)
setSharedFormula(setCell(ws, 2, 1, null), '=A1*2', { si: 0, master: true, ref: 'A2:A10' });
\`\`\`

\`cachedValue\` is optional. Excel will recalc on open anyway, but providing
it keeps the file viewable in tools that don't recalculate (Quick Look,
preview generators, headless readers).

## Rich text (multi-format cell content)

\`\`\`ts
import { makeRichText, makeTextRun } from '@office-kit/xlsx/cell';
import { setCell } from '@office-kit/xlsx/worksheet';

const rich = makeRichText([
  makeTextRun('Hello, ',  { bold: true }),
  makeTextRun('world',    { italic: true, color: { rgb: 'FFFF0000' } }),
]);
setCell(ws, 1, 1, { kind: 'rich-text', runs: rich });
\`\`\`

## Hyperlinks

\`\`\`ts
import { setHyperlink } from '@office-kit/xlsx/worksheet';

setHyperlink(ws, 'A1', {
  target: 'https://example.com',
  display: 'Open',                    // tooltip / fallback display text
  tooltip: 'Open example.com',
});
\`\`\`

Hyperlinks live separately from cell values. The cell holds the visible
text via \`setCell\`; \`setHyperlink\` wires up the URL underneath.

## Merge and freeze

\`\`\`ts
import { mergeCells, unmergeCells, makeFreezePane, makeSheetView } from '@office-kit/xlsx/worksheet';

mergeCells(ws, 'A1:C1');                       // accepts string or { from, to } range
ws.views.push(makeSheetView({ pane: makeFreezePane('A2') }));   // freeze row 1
ws.views.push(makeSheetView({ pane: makeFreezePane('B1') }));   // freeze col A
ws.views.push(makeSheetView({ pane: makeFreezePane('B2') }));   // freeze both
\`\`\`

## Excel Tables, AutoFilter, Data Validation, Conditional Formatting

\`\`\`ts
import {
  addExcelTable, addAutoFilter,
  makeDataValidation, addDataValidation,
  makeCfRule, makeConditionalFormatting, addConditionalFormatting,
} from '@office-kit/xlsx/worksheet';

// Excel Table — named range with banded styling and a filter dropdown on
// every header. Use the higher-level addExcelTable; for full control,
// makeTableDefinition + ws.tables.set(...).
addExcelTable(wb, ws, {
  name: 'Items',
  ref: 'A1:C4',
  columns: ['SKU', 'Name', 'Price'],
  style: 'TableStyleMedium2',
});

// Just an autoFilter (no banded styling)
addAutoFilter(ws, 'A1:C1');

// Dropdown data validation
addDataValidation(ws, makeDataValidation({
  type: 'list',
  sqref: 'B2:B100',
  formula1: '"Open,In progress,Closed"',
  prompt: 'Pick a status',
  errorTitle: 'Invalid value',
  error: 'Pick one of the listed values.',
}));

// 3-color scale heat-map
addConditionalFormatting(ws, makeConditionalFormatting({
  sqref: 'A1:A10',
  rules: [makeCfRule({
    type: 'colorScale',
    priority: 1,
    formulas: [],
    innerXml:
      '<colorScale>' +
      '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' +
      '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>' +
      '</colorScale>',
  })],
}));
\`\`\`

## Charts and images

\`\`\`ts
import { makeBarChart, makeBarSeries, makeChartSpace } from '@office-kit/xlsx/chart';
import { addChartAt, addImageAt, loadImage } from '@office-kit/xlsx/drawing';
import { readFile } from 'node:fs/promises';

// Bar / column chart anchored to D2.
const chart = makeBarChart({
  barDir: 'col',
  grouping: 'clustered',
  series: [
    makeBarSeries({
      idx: 0,
      tx:  { kind: 'literal', value: 'Revenue' },
      cat: { ref: 'Sales!$A$2:$A$4' },
      val: { ref: 'Sales!$B$2:$B$4' },
    }),
  ],
});
const space = makeChartSpace({
  plotArea: { chart },
  title: 'Revenue by region',
  legend: { position: 'r' },
});
addChartAt(ws, 'D2', { space }, { widthPx: 480, heightPx: 320 });

// Image at a cell — PNG / JPEG / GIF / BMP / WebP / TIFF / SVG / EMF / WMF.
// Format and dimensions are auto-detected from the bytes.
const img = loadImage(await readFile('logo.png'));
addImageAt(ws, 'B2', img, { widthPx: 200, heightPx: 80 });
\`\`\`

Available chart kinds:

- Legacy \`c:\` family (\`@office-kit/xlsx/chart\`): \`makeBarChart\`,
  \`makeLineChart\`, \`makeAreaChart\`, \`makePieChart\`,
  \`makeDoughnutChart\`, \`makeScatterChart\`, \`makeRadarChart\`,
  \`makeBubbleChart\`, \`makeStockChart\`, \`makeSurfaceChart\`,
  \`makeOfPieChart\`. Wrap them in a \`PlotArea\` and pass to
  \`makeChartSpace\`.
- Modern \`cx:\` chartex family (same subpath): \`makeSunburstChart\`,
  \`makeTreemapChart\`, \`makeWaterfallChart\`, \`makeHistogramChart\`,
  \`makeParetoChart\`, \`makeFunnelChart\`, \`makeBoxWhiskerChart\`,
  \`makeRegionMapChart\`. These return a \`CxChartSpace\` directly.

For absolute positioning instead of a one-cell anchor, use
\`makeAbsoluteAnchor\` or \`makeTwoCellAnchor\` from \`@office-kit/xlsx/drawing\`.

## Multiple sheets, defined names, chartsheets

\`\`\`ts
import {
  createWorkbook, addWorksheet, addChartsheet, addDefinedName,
} from '@office-kit/xlsx/workbook';

const wb = createWorkbook();
addWorksheet(wb, 'Q1');
addWorksheet(wb, 'Q2');
const cs = addChartsheet(wb, 'Dashboard');  // chartsheet has no cell grid
addDefinedName(wb, { name: 'Totals', value: 'Q1!$A$1:$A$10,Q2!$A$1:$A$10' });
\`\`\`

## Errors

Every error thrown by @office-kit/xlsx is a subclass of \`OpenXmlError\`
(\`@office-kit/xlsx/utils\`). Catch the specific subclass that matters:

| Error                              | When                                                                                                |
|------------------------------------|-----------------------------------------------------------------------------------------------------|
| \`OpenXmlIoError\`                   | Source / sink failure (bad bytes, broken zip, partial read, write error)                            |
| \`OpenXmlSchemaError\`               | Input does not conform to the ECMA-376 OOXML schema                                                  |
| \`OpenXmlDecompressionBombError\`    | Archive exceeded \`decompressionLimits\` (per-entry size, total size, or compression ratio)        |
| \`OpenXmlInvalidWorkbookError\`      | Workbook structurally invalid (missing parts, broken relationships)                                  |
| \`OpenXmlNotImplementedError\`       | Feature is not yet supported (e.g. ZIP64 write, encrypted decryption)                               |

\`decompressionLimits\` is **on by default** in both \`loadWorkbook\` and
\`loadWorkbookStream\`. Keep it on when reading untrusted input.

Encrypted xlsx files (CFB Compound Documents) are detected and rejected
with a clear error pointing at
[\`msoffcrypto-tool\`](https://github.com/nolze/msoffcrypto-tool); decrypt
externally first, then load the resulting plain xlsx.

## Common pitfalls

- **No root barrel.** \`import { ... } from '@office-kit/xlsx'\` doesn't work. Use a
  subpath import like \`@office-kit/xlsx/io\`, \`@office-kit/xlsx/worksheet\`, etc.
- **1-indexed coordinates.** \`setCell(ws, 1, 1, ...)\` writes \`A1\`, not
  \`B2\`.
- **\`wb.sheets\` is a discriminated union.** Narrow on
  \`kind === 'worksheet'\` before reading \`.sheet\`; chartsheets are also
  in the array but have a different shape.
- **Streaming-write requires \`close\` + \`finalize\`.** Skipping either
  produces a truncated zip.
- **Streaming-write locks \`<cols>\` after the first row.** Configure
  column widths up front.
- **ZIP64 write is not supported.** Workbooks with > 65 535 entries cannot
  be written; you'll get an \`OpenXmlNotImplementedError\`. Read of ZIP64
  archives works.
- **\`.xls\` / \`.xlsb\` / \`.ods\` / \`.csv\` are out of scope.** @office-kit/xlsx
  reads and writes \`.xlsx\` (and \`.xlsm\` macro-enabled workbooks).
  For older formats, reach for SheetJS.
- **Color values are AARRGGBB hex strings.** \`'FF3366CC'\` is opaque blue.
  Forgetting the alpha prefix gives the wrong color.
- **Don't swallow \`OpenXmlError\`.** Internal-logic errors propagate as
  themselves — they're bugs, not validation failures.

## Cheatsheet

| Task | Functions |
|------|-----------|
| Read xlsx file (Node) → Workbook | \`loadWorkbook\` + \`fromFile\` |
| Read xlsx Buffer (Node) → Workbook | \`loadWorkbook\` + \`fromBuffer\` |
| Read xlsx from \`fetch\` (browser) → Workbook | \`loadWorkbook\` + \`fromResponse\` |
| Read xlsx from \`<input type="file">\` | \`loadWorkbook\` + \`fromBlob\` |
| Iterate every cell of a worksheet | \`iterRows\` (or \`iterValues\`) |
| Find populated extents | \`getMaxRow\` + \`getMaxCol\` |
| Stream-read huge sheet | \`loadWorkbookStream\` + \`openWorksheet\` + \`iterRows\` |
| Stream-read rows N..M | \`loadWorkbookStream\` + \`iterRows({ minRow, maxRow })\` |
| Build a workbook → \`Uint8Array\` | \`createWorkbook\` + \`addWorksheet\` + \`setCell\` + \`workbookToBytes\` |
| Build a workbook → file (Node) | \`createWorkbook\` + \`addWorksheet\` + \`setCell\` + \`saveWorkbook\` + \`toFile\` |
| Build a workbook → \`Buffer\` (Node) | \`workbookToBuffer\` |
| Build a workbook → \`Blob\` (browser) | \`toBlob\` |
| Edit one cell of an existing file | \`loadWorkbook\` + \`setCell\` + \`saveWorkbook\` |
| Append rows | \`appendRow\` / \`appendRows\` |
| Stream-write millions of rows | \`createWriteOnlyWorkbook\` + \`appendRow\` + \`ws.close\` + \`wb.finalize\` |
| Bold + size + fill on header | \`setBold\` + \`setFontSize\` + \`setCellBackgroundColor\` |
| Currency / percent | \`setCellAsCurrency\` / \`setCellAsPercent\` |
| Date number format | \`setCellNumberFormat\` + \`FORMAT_DATE_DATETIME\` |
| Formula with cached value | \`setCell\` + \`setFormula\` |
| Hyperlink | \`setHyperlink\` |
| Merge + freeze header | \`mergeCells\` + \`makeFreezePane\` + \`makeSheetView\` |
| Defined name across sheets | \`addDefinedName\` |
| Excel Table | \`addExcelTable\` |
| AutoFilter only | \`addAutoFilter\` |
| Dropdown validation | \`makeDataValidation\` + \`addDataValidation\` |
| 3-color scale heat-map | \`makeCfRule\` + \`makeConditionalFormatting\` + \`addConditionalFormatting\` |
| Image at cell | \`loadImage\` + \`addImageAt\` |
| Clustered column chart | \`makeBarChart\` + \`makeBarSeries\` + \`makeChartSpace\` + \`addChartAt\` |
| A1 ↔ row/col | \`coordinateFromString\` / \`tupleToCoordinate\` (\`@office-kit/xlsx/utils\`) |
| Convert Date ↔ Excel serial | \`dateToExcel\` / \`excelToDate\` (\`@office-kit/xlsx/utils\`) |
`;

function buildBody(): string {
  // Relative paths so the index works under any base path (custom domain,
  // user page, or project page like /@office-kit/xlsx/). `docs/install.md` from
  // /llms.txt resolves to <base>/docs/install.md regardless.
  const sections = docSections
    .map((section) => {
      const lines = section.links.map(
        (l) => `- [${l.title}](.${l.href}.md): ${l.description}`,
      );
      return `## ${section.title}\n\n${lines.join('\n')}`;
    })
    .join('\n\n');

  return `${HEADER}\n${sections}\n\n## Source\n\n- [GitHub repository](https://github.com/office-kit/xlsx)\n- [npm package](https://www.npmjs.com/package/@office-kit/xlsx)\n`;
}

export const GET: RequestHandler = () => {
  return new Response(buildBody(), {
    headers: {
      'Content-Type': 'text/plain; charset=utf-8',
      'Cache-Control': 'public, max-age=300',
    },
  });
};
