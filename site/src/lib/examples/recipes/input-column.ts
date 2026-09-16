// A column the recipient is meant to fill in: Excel's "Input" style marks it
// as editable, and a decimal validation keeps what they type usable.

import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import { applyBuiltinStyle } from '@office-kit/xlsx/styles';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import {
  addDataValidation,
  appendRows,
  ensureCell,
  makeDataValidation,
  setCell,
} from '@office-kit/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Quote');

appendRows(ws, [
  ['Language', 'Words', 'Rate per word'],
  ['de', 71_579],
  ['fr', 12_004],
]);

// Column C is left empty on purpose. ensureCell reaches the cell without
// writing over anything, so the style lands on a genuinely blank cell.
setCell(ws, 1, 3, 'Rate per word');
for (let row = 2; row <= 3; row++) {
  applyBuiltinStyle(wb, ensureCell(ws, row, 3), 'Input');
}

addDataValidation(
  ws,
  makeDataValidation({
    type: 'decimal',
    operator: 'between',
    sqref: 'C2:C3',
    formula1: '0',
    formula2: '10',
    prompt: 'Rate per word, 0 to 10',
    errorTitle: 'Out of range',
    error: 'Enter a rate between 0 and 10.',
  }),
);

await saveWorkbook(wb, toFile('quote.xlsx'));
