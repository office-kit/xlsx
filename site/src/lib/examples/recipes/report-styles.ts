// Style a report by registering each look once, then writing cells with the
// id. There is no second pass over the rows, so nothing can overwrite what the
// first pass wrote.

import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import {
  makeAlignment,
  makeBorder,
  makeColor,
  makeFont,
  makePatternFill,
  makeSide,
  patchCellFont,
  registerCellStyle,
} from '@office-kit/xlsx/styles';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { appendRow, setCell, setColumnWidths, setFreezePanes } from '@office-kit/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Leverage');

const thin = makeSide({ style: 'thin' });
const box = makeBorder({ left: thin, right: thin, top: thin, bottom: thin });

const HEADER = registerCellStyle(wb, {
  font: makeFont({ name: 'Calibri', size: 11, bold: true }),
  fill: makePatternFill({ patternType: 'solid', fgColor: makeColor({ rgb: 'FFEFEFEF' }) }),
  border: box,
  alignment: makeAlignment({ wrapText: true, vertical: 'center' }),
});
const TEXT = registerCellStyle(wb, { border: box });
const INT = registerCellStyle(wb, { border: box, numberFormat: '#,##0' });

// A title needs one field changed, not a whole font: patchCellFont keeps the
// workbook default (Calibri 11) for everything it doesn't mention.
patchCellFont(wb, setCell(ws, 1, 1, 'Translation memory leverage'), { bold: true, size: 13 });

const headers = ['Language', 'Total words', 'Leveraged', 'New'];
appendRow(ws, headers, { styleIds: headers.map(() => HEADER) });
setFreezePanes(ws, { rows: 2, cols: 0 });

const rows: ReadonlyArray<readonly [string, number, number, number]> = [
  ['de', 71_579, 52_310, 19_269],
  ['fr', 12_004, 9_880, 2_124],
];
const dataStyles = [TEXT, INT, INT, INT];
for (const row of rows) appendRow(ws, [...row], { styleIds: dataStyles });

setColumnWidths(ws, [18, 14, 14, 14]);

await saveWorkbook(wb, toFile('leverage.xlsx'));
