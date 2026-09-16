// Recipe registry. Each entry pairs a piece of human-readable framing
// with the literal source of a real .ts file in this directory; the
// file is type-checked against `@office-kit/xlsx` on every build, so the
// snippet shown to readers can never drift from the live API.

import openAndIterate from './open-and-iterate.ts?raw';
import buildFromScratch from './build-from-scratch.ts?raw';
import styleCells from './style-cells.ts?raw';
import reportStyles from './report-styles.ts?raw';
import numberFormats from './number-formats.ts?raw';
import formulas from './formulas.ts?raw';
import addBarChart from './add-bar-chart.ts?raw';
import insertImage from './insert-image.ts?raw';
import tablesWithFilter from './tables-with-filter.ts?raw';
import inputColumn from './input-column.ts?raw';
import dropdownValidation from './dropdown-validation.ts?raw';
import conditionalColorScale from './conditional-color-scale.ts?raw';
import hyperlinks from './hyperlinks.ts?raw';
import mergeAndFreeze from './merge-and-freeze.ts?raw';
import multiSheet from './multi-sheet.ts?raw';
import browserFileInput from './browser-file-input.ts?raw';
import assertGeneratedWorkbook from './assert-generated-workbook.ts?raw';
import deterministicBytes from './deterministic-bytes.ts?raw';

import basicReadWrite from '../basic-read-write.ts?raw';
import nodeFs from '../node-fs.ts?raw';
import browserFetch from '../browser-fetch.ts?raw';
import streamingWrite from '../streaming-write.ts?raw';
import streamingRead from '../streaming-read.ts?raw';

export type Recipe = {
  /** URL-safe slug, used as the anchor on /docs/recipes. */
  slug: string;
  title: string;
  /** One-line teaser shown above the code block. */
  teaser: string;
  /** Repo path (used as caption above the snippet). */
  path: string;
  /** Verbatim source. */
  source: string;
  /** Optional bullet points to add context after the snippet. */
  notes?: string[];
  /** Names of related public exports — link to /api anchors when available. */
  relatedApi?: string[];
};

export const recipeGroups: Array<{ title: string; recipes: Recipe[] }> = [
  {
    title: 'Basics',
    recipes: [
      {
        slug: 'open-and-read',
        title: 'Open a workbook and read every cell',
        teaser:
          'Load an existing xlsx, narrow the first sheet to a worksheet, and walk every cell.',
        path: 'site/src/lib/examples/recipes/open-and-iterate.ts',
        source: openAndIterate,
        notes: [
          'wb.sheets is a discriminated union — narrow on `kind === "worksheet"` to reach the Worksheet shape (chartsheets have a different surface).',
          'For huge sheets, prefer `loadWorkbookStream` + `iterRows` instead — see the streaming recipe below.',
        ],
        relatedApi: ['loadWorkbook', 'fromFile', 'Worksheet'],
      },
      {
        slug: 'edit-and-save',
        title: 'Edit a single cell and save',
        teaser:
          'The canonical round-trip: load → mutate → write back. Same as the Quick start.',
        path: 'site/src/lib/examples/basic-read-write.ts',
        source: basicReadWrite,
        relatedApi: ['loadWorkbook', 'setCell', 'workbookToBytes', 'fromBuffer'],
      },
      {
        slug: 'build-from-scratch',
        title: 'Build a workbook from scratch',
        teaser:
          'No input file — start with `createWorkbook`, add a sheet, write cells, save.',
        path: 'site/src/lib/examples/recipes/build-from-scratch.ts',
        source: buildFromScratch,
        relatedApi: ['createWorkbook', 'addWorksheet', 'setCell', 'saveWorkbook'],
      },
      {
        slug: 'multi-sheet',
        title: 'Multiple sheets + named ranges',
        teaser:
          'Add several worksheets, define names that span them, and reference them in a formula.',
        path: 'site/src/lib/examples/recipes/multi-sheet.ts',
        source: multiSheet,
        relatedApi: ['addWorksheet', 'addDefinedName', 'makeFormula'],
      },
      {
        slug: 'node-fs-helpers',
        title: 'Direct fs helpers (Node)',
        teaser:
          '`fromFile` / `toFile` skip the manual `readFile` / `writeFile` glue.',
        path: 'site/src/lib/examples/node-fs.ts',
        source: nodeFs,
        relatedApi: ['loadWorkbook', 'saveWorkbook', 'fromFile', 'toFile'],
      },
    ],
  },
  {
    title: 'Cells & values',
    recipes: [
      {
        slug: 'styling',
        title: 'Style a header cell',
        teaser:
          'Bold, font size, fill color, center alignment, and a thin border in five lines.',
        path: 'site/src/lib/examples/recipes/style-cells.ts',
        source: styleCells,
        notes: [
          'These helpers are *cell-level* shortcuts. For range-wide changes, look at `setRangeFont`, `setRangeAlignment`, `setRangeBorderBox`, etc.',
          'Background colors are hex `AARRGGBB` strings — leading `FF` is opaque alpha.',
          '`setBold` and friends merge into the existing font. `setCellFont` replaces it whole, which drops the workbook default (Calibri 11) unless the `Font` you pass is complete. Use `patchCellFont` to change several fields at once.',
          'Styling a whole report cell by cell adds up. `registerCellStyle` (next recipe) builds each look once and hands you an id the write itself carries.',
        ],
        relatedApi: [
          'setBold',
          'setFontSize',
          'setCellBackgroundColor',
          'centerCell',
          'setCellBorderAll',
        ],
      },
      {
        slug: 'report-styles',
        title: 'Style a whole report by style id',
        teaser:
          '`registerCellStyle` returns a `styleId`; `setCell` and `appendRow` take it, so formatting arrives with the value instead of in a second pass.',
        path: 'site/src/lib/examples/recipes/report-styles.ts',
        source: reportStyles,
        notes: [
          'A `styleId` is a complete style, not a patch: an axis you leave out of the spec renders as the workbook default even if the target cell had something there.',
          'Equal specs dedup to one xf, so reusing three ids across a thousand rows costs three records.',
          '`appendRow`\'s `styleIds` are positional. A column with an id is written even when its value is empty, which is how a bordered-but-blank input column survives the append.',
        ],
        relatedApi: ['registerCellStyle', 'setCell', 'appendRow', 'patchCellFont'],
      },
      {
        slug: 'number-formats',
        title: 'Number formats: currency, percent, dates',
        teaser:
          '`setCellAsCurrency` and `setCellAsPercent` are one-shot; everything else goes through `setCellNumberFormat` + a built-in or custom format code.',
        path: 'site/src/lib/examples/recipes/number-formats.ts',
        source: numberFormats,
        relatedApi: [
          'setCellAsCurrency',
          'setCellAsPercent',
          'setCellNumberFormat',
          'FORMAT_DATE_DATETIME',
          'FORMAT_PERCENTAGE',
        ],
      },
      {
        slug: 'formulas',
        title: 'Formulas in a generated workbook',
        teaser:
          'Cache the values you can compute, and set `fullCalcOnLoad` for the ones you cannot.',
        path: 'site/src/lib/examples/recipes/formulas.ts',
        source: formulas,
        notes: [
          'Viewers fall into three camps. Excel, LibreOffice and Google Sheets compute a formula with no cached `<v>` on open. Quick Look, Outlook and SharePoint previews, and most thumbnailers never calculate and render the cell empty. Tools that read the XML directly see whatever you wrote. Supply `cachedValue` wherever the producer can compute it, and call `setFullCalcOnLoad(wb, true)` on any generated workbook with formulas so the ones you left uncached get filled in on first open.',
          '`makeFormula` builds the value for a `setCell` write, so a formula plus its number format is one call when the style comes from `registerCellStyle`. `setFormula` is the same thing applied to a cell you already hold.',
          'A leading `=` is stripped, so `\'=SUM(A1:A3)\'` and `\'SUM(A1:A3)\'` are interchangeable. OOXML stores `<f>` without it.',
          'For shared and array formulas, use `setSharedFormula` / `setArrayFormula` on a Cell from `ensureCell`.',
        ],
        relatedApi: [
          'makeFormula',
          'setFormula',
          'setFullCalcOnLoad',
          'setArrayFormula',
          'setSharedFormula',
        ],
      },
      {
        slug: 'merge-and-freeze',
        title: 'Merge cells + freeze the header row',
        teaser:
          'Merge a title across columns, then freeze row 1 so it stays put while scrolling.',
        path: 'site/src/lib/examples/recipes/merge-and-freeze.ts',
        source: mergeAndFreeze,
        relatedApi: ['mergeCells', 'makeFreezePane', 'makeSheetView'],
      },
      {
        slug: 'hyperlinks',
        title: 'Make a cell clickable',
        teaser: 'Hyperlinks live separately from cell values — set the text, attach the URL.',
        path: 'site/src/lib/examples/recipes/hyperlinks.ts',
        source: hyperlinks,
        relatedApi: ['setHyperlink'],
      },
    ],
  },
  {
    title: 'Tables, validation, conditional formatting',
    recipes: [
      {
        slug: 'tables',
        title: 'Promote a range to an Excel Table',
        teaser:
          'Excel Tables get banded styling, a built-in autoFilter on every header, and a name you can reference in formulas.',
        path: 'site/src/lib/examples/recipes/tables-with-filter.ts',
        source: tablesWithFilter,
        notes: [
          'Pass `style` for one-arg style selection or `styleInfo` for full control over banded rows / columns.',
          'Write the header row first: `addExcelTable` checks that each header cell already holds its column name, and that the column count matches the range width. Excel repairs a file where they disagree by dropping the table. Pass `headerRowCount: 0` for a genuinely header-less table.',
          'For just a filter without table styling, use `setAutoFilter(ws, makeAutoFilter({ ref: "A1:C4" }))`. A Table also gives you banded rows, structured references and a stable name for pivots, with no per-cell border work.',
        ],
        relatedApi: ['addExcelTable', 'setAutoFilter', 'makeAutoFilter'],
      },
      {
        slug: 'dropdown-validation',
        title: 'Dropdown data validation',
        teaser:
          'Restrict a range to a list of allowed values. Excel renders a dropdown arrow on each cell.',
        path: 'site/src/lib/examples/recipes/dropdown-validation.ts',
        source: dropdownValidation,
        notes: [
          'Pass a sheet-relative formula (`=Sheet1!$A$1:$A$10`) instead of a literal array if the choices come from another range.',
        ],
        relatedApi: ['makeDataValidation', 'addDataValidation'],
      },
      {
        slug: 'input-column',
        title: 'A column the recipient fills in',
        teaser:
          'Excel\'s built-in "Input" style marks a column as editable; a decimal validation keeps what they type usable.',
        path: 'site/src/lib/examples/recipes/input-column.ts',
        source: inputColumn,
        notes: [
          '`ensureCell` reaches a blank cell without writing over it, so the style lands on an empty cell rather than one you just cleared.',
          'Pair this with `setRangeProtection(wb, ws, "C2:C3", { locked: false })` and a sheet protection if the rest of the sheet should be read-only.',
        ],
        relatedApi: ['applyBuiltinStyle', 'ensureCell', 'makeDataValidation', 'addDataValidation'],
      },
      {
        slug: 'color-scale',
        title: 'Heat-map with a 3-color scale',
        teaser:
          'Build a `colorScale` rule with `makeCfRule` + inner XML and attach it via `addConditionalFormatting`.',
        path: 'site/src/lib/examples/recipes/conditional-color-scale.ts',
        source: conditionalColorScale,
        relatedApi: ['makeCfRule', 'makeConditionalFormatting', 'addConditionalFormatting'],
      },
    ],
  },
  {
    title: 'Charts & images',
    recipes: [
      {
        slug: 'bar-chart',
        title: 'Add a clustered column chart',
        teaser:
          'Wire a `BarChart` to a data range and anchor it to a cell with `addChartAt`.',
        path: 'site/src/lib/examples/recipes/add-bar-chart.ts',
        source: addBarChart,
        notes: [
          'Same pattern works for `makeLineChart`, `makePieChart`, `makeScatterChart` and friends — wrap them in a `PlotArea` and pass to `makeChartSpace`.',
          'For modern chart kinds (Sunburst, Treemap, Waterfall, Histogram, Pareto, Funnel, BoxWhisker, RegionMap), use the `makeSunburstChart` / `makeTreemapChart` / ... helpers from the chartex family — they emit `cx:` chart space.',
        ],
        relatedApi: [
          'makeBarChart',
          'makeBarSeries',
          'makeChartSpace',
          'addChartAt',
          'makeSunburstChart',
        ],
      },
      {
        slug: 'insert-image',
        title: 'Insert an image at a cell',
        teaser:
          'Drop a PNG / JPEG / GIF / BMP / WebP / TIFF / SVG anchored to a cell — format and dimensions are auto-detected.',
        path: 'site/src/lib/examples/recipes/insert-image.ts',
        source: insertImage,
        relatedApi: ['loadImage', 'addImageAt', 'makeOneCellAnchor'],
      },
    ],
  },
  {
    title: 'Generating files you have to trust',
    recipes: [
      {
        slug: 'assert-generated-workbook',
        title: 'Assert on a workbook you just generated',
        teaser:
          'Load the bytes back and read them with the same API you wrote them with.',
        path: 'site/src/lib/examples/recipes/assert-generated-workbook.ts',
        source: assertGeneratedWorkbook,
        notes: [
          '`fromArrayBuffer` accepts a `Uint8Array` as well as an `ArrayBuffer`, so a renderer\'s output goes straight into `loadWorkbook` with no copy and no temp file.',
          '`getSheet(wb, title)` narrows past the worksheet / chartsheet union, so there is no `kind === "worksheet"` check to write.',
          '`addWorksheet` already validates the title (31-character limit, `[]:*?/\\` and the reserved name `History`), so a test of your own for those is testing this library.',
        ],
        relatedApi: ['loadWorkbook', 'fromArrayBuffer', 'getSheet', 'getRangeValues', 'iterCells'],
      },
      {
        slug: 'deterministic-bytes',
        title: 'Byte-identical output for identical input',
        teaser:
          'Pin `mtime` and the core properties, and the same payload always renders the same bytes.',
        path: 'site/src/lib/examples/recipes/deterministic-bytes.ts',
        source: deterministicBytes,
        notes: [
          'ZIP has no "no timestamp" encoding: each entry carries a DOS mtime, and without `mtime` it comes from the wall clock. That alone makes two renders of the same payload differ.',
          '`createWriteOnlyWorkbook` takes the same option, as does `compressionLevel` (0 stores, 9 is smallest) on both paths.',
          'Core properties are the other moving part. Set `created` / `modified` from your payload, not from `new Date()`.',
        ],
        relatedApi: ['workbookToBytes', 'saveWorkbook', 'createWriteOnlyWorkbook'],
      },
    ],
  },
  {
    title: 'Streaming (huge sheets)',
    recipes: [
      {
        slug: 'streaming-write',
        title: 'Write 10M rows in a fixed memory budget',
        teaser:
          '`createWriteOnlyWorkbook` deflates each row as it arrives — heap stays under 100 MB even for 10M-row sheets.',
        path: 'site/src/lib/examples/streaming-write.ts',
        source: streamingWrite,
        notes: [
          '`setColumnWidth` must run *before* the first `appendRow` — once any row is written, `<cols>` is locked.',
          '`ws.close()` and `wb.finalize()` are required — that\'s when the central directory is written.',
        ],
        relatedApi: ['createWriteOnlyWorkbook', 'toFile'],
      },
      {
        slug: 'streaming-read',
        title: 'Iterate a huge sheet without loading it',
        teaser:
          '`loadWorkbookStream` + `iterRows` walks the file once and yields rows as they\'re parsed.',
        path: 'site/src/lib/examples/streaming-read.ts',
        source: streamingRead,
        notes: [
          'Bound the walk with `iterRows({ minRow, maxRow, minCol, maxCol })` — the parser skips ahead via tag-scan.',
        ],
        relatedApi: ['loadWorkbookStream', 'fromFile'],
      },
    ],
  },
  {
    title: 'Browser',
    recipes: [
      {
        slug: 'browser-fetch',
        title: 'Browser: read xlsx from a fetch response',
        teaser:
          '`fromResponse` is streaming, so the workbook starts parsing while bytes are still arriving.',
        path: 'site/src/lib/examples/browser-fetch.ts',
        source: browserFetch,
        relatedApi: ['fromResponse', 'loadWorkbook'],
      },
      {
        slug: 'browser-file-input',
        title: 'Browser: read xlsx from <input type="file">',
        teaser: '`fromBlob` consumes the File the user just picked, no full buffer.',
        path: 'site/src/lib/examples/recipes/browser-file-input.ts',
        source: browserFileInput,
        relatedApi: ['fromBlob', 'loadWorkbook'],
      },
    ],
  },
];

export const allRecipes: Recipe[] = recipeGroups.flatMap((g) => g.recipes);
