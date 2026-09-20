// The REPL's starter code. The page loads this file as text and shows
// everything below the `repl:start` marker; the imports above it exist so
// svelte-check compiles the snippet against the live library, and an API rename
// breaks the build instead of the REPL.
//
// The REPL runs the text as JavaScript with these names already in scope, so
// below the marker write only what is both valid JavaScript and valid strict
// TypeScript: no annotations, no `as`, and no function whose parameters would
// need a type.

// The spread in `orders.map` copies six small objects once; the rule is about
// hot paths, and the alternative reads worse in code people learn from.
/* oxlint-disable oxc/no-map-spread */

import { makeFormula } from '@office-kit/xlsx/cell';
import {
  makeAlignment,
  makeBorder,
  makeColor,
  makeFont,
  makePatternFill,
  makeSide,
  registerCellStyle,
} from '@office-kit/xlsx/styles';
import { addWorksheet, type Workbook } from '@office-kit/xlsx/workbook';
import { appendRow, mergeCells, setColumnWidths, setFreezePanes } from '@office-kit/xlsx/worksheet';

declare const wb: Workbook;

// repl:start
// Everything exported from @office-kit/xlsx/cell, /styles, /workbook,
// /worksheet and /utils is in scope, so there is nothing to import here.
// In your own code, import each function from its subpath.
// `wb` is an empty workbook from createWorkbook().

// The library never calculates, so work out each formula's result (`net`) and
// cache it. Viewers that do not calculate then show a number, not a blank.
const orders = [
  { date: '2026-07-03', customer: 'Aoki Foods', units: 120, price: 18.5, discount: 0 },
  { date: '2026-07-09', customer: 'Birch & Co', units: 45, price: 64, discount: 0.1 },
  { date: '2026-07-17', customer: 'Aoki Foods', units: 300, price: 17.25, discount: 0.05 },
  { date: '2026-08-02', customer: 'Cedar Labs', units: 8, price: 1250, discount: 0 },
  { date: '2026-08-21', customer: 'Birch & Co', units: 60, price: 64, discount: 0.15 },
  { date: '2026-09-05', customer: 'Cedar Labs', units: 14, price: 1250, discount: 0.08 },
].map((o) => ({ ...o, net: o.units * o.price * (1 - o.discount) }));

const green = makeColor({ rgb: 'FF168A4F' });
const white = makeColor({ rgb: 'FFFFFFFF' });
const mint = makeColor({ rgb: 'FFE6F4EC' });
const wash = makePatternFill({ patternType: 'solid', fgColor: mint });

// Register each look once. A cell then carries a style id, not a style.
const title = registerCellStyle(wb, { font: makeFont({ bold: true, size: 14, color: green }) });
const header = registerCellStyle(wb, {
  font: makeFont({ bold: true, color: white }),
  fill: makePatternFill({ patternType: 'solid', fgColor: green }),
  alignment: makeAlignment({ horizontal: 'center' }),
});
const formats = ['yyyy-mm-dd', 'General', '#,##0', '"$"#,##0.00', '0%', '"$"#,##0.00'];
const plain = formats.map((numberFormat) => registerCellStyle(wb, { numberFormat }));
const striped = formats.map((numberFormat) => registerCellStyle(wb, { numberFormat, fill: wash }));
const totals = formats.map((numberFormat) =>
  registerCellStyle(wb, {
    numberFormat,
    font: makeFont({ bold: true }),
    border: makeBorder({ top: makeSide({ style: 'medium', color: green }) }),
  }),
);

const ws = addWorksheet(wb, 'Orders');
appendRow(ws, ['Orders, third quarter 2026'], { styleIds: [title] });
mergeCells(ws, 'A1:F1');
appendRow(ws, ['Date', 'Customer', 'Units', 'Unit price', 'Discount', 'Net'], {
  styleIds: formats.map(() => header),
});
setFreezePanes(ws, 'A3');

const FIRST = 3;
const last = FIRST + orders.length - 1;
orders.forEach((o, i) => {
  const row = FIRST + i;
  const formula = makeFormula(`C${row}*D${row}*(1-E${row})`, { cachedValue: o.net });
  // A date-only ISO string parses as UTC, so the cell holds the same day everywhere.
  appendRow(ws, [new Date(o.date), o.customer, o.units, o.price, o.discount, formula], {
    styleIds: i % 2 === 1 ? striped : plain,
  });
});
appendRow(
  ws,
  [
    'Total',
    undefined,
    makeFormula(`SUM(C${FIRST}:C${last})`, { cachedValue: orders.reduce((n, o) => n + o.units, 0) }),
    undefined,
    undefined,
    makeFormula(`SUM(F${FIRST}:F${last})`, { cachedValue: orders.reduce((n, o) => n + o.net, 0) }),
  ],
  { styleIds: totals },
);
setColumnWidths(ws, [12, 16, 9, 12, 10, 14]);

// A second sheet that reads the first one through cross-sheet formulas.
const summary = addWorksheet(wb, 'By customer');
appendRow(summary, ['Customer', 'Orders', 'Net'], { styleIds: [header, header, header] });
[...new Set(orders.map((o) => o.customer))].forEach((customer, i) => {
  const row = i + 2;
  const mine = orders.filter((o) => o.customer === customer);
  appendRow(
    summary,
    [
      customer,
      makeFormula(`COUNTIF(Orders!B:B,A${row})`, { cachedValue: mine.length }),
      makeFormula(`SUMIF(Orders!B:B,A${row},Orders!F:F)`, { cachedValue: mine.reduce((n, o) => n + o.net, 0) }),
    ],
    { styleIds: [plain[1], plain[2], plain[5]] },
  );
});
setColumnWidths(summary, [16, 9, 14]);
