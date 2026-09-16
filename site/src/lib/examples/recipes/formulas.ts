// Formulas in a generated workbook: cache the value you already know, and
// ask Excel to recalculate the ones you don't.

import { makeFormula } from '@office-kit/xlsx/cell';
import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import { setCellNumberFormat } from '@office-kit/xlsx/styles';
import { addWorksheet, createWorkbook, setFullCalcOnLoad } from '@office-kit/xlsx/workbook';
import { setCell } from '@office-kit/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Sheet1');

const units = [12, 18, 30];
units.forEach((n, i) => setCell(ws, i + 1, 1, n));

// The producer can add these up, so cache the result: viewers that never
// calculate show the number instead of a blank cell.
const total = units.reduce((a, b) => a + b, 0);
setCellNumberFormat(wb, setCell(ws, 4, 1, makeFormula('SUM(A1:A3)', { cachedValue: total })), '#,##0');

// Where the value depends on data this renderer doesn't have, leave it
// uncached and mark the workbook for a full recalculation on open.
setCell(ws, 5, 1, makeFormula('SUMIFS(Other!B:B, Other!A:A, A1)'));
setFullCalcOnLoad(wb, true);

await saveWorkbook(wb, toFile('with-formulas.xlsx'));
