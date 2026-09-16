// Assert on the bytes a renderer produced, by loading them back.
//
// `fromArrayBuffer` takes the Uint8Array directly, so there is no Buffer or
// temp file between the renderer and the assertions.

import { getCoordinate, getFormulaText, makeFormula } from '@office-kit/xlsx/cell';
import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
import { addWorksheet, createWorkbook, getSheet, sheetNames } from '@office-kit/xlsx/workbook';
import {
  appendRows,
  getAutoFilter,
  getRangeValues,
  iterCells,
  makeAutoFilter,
  setAutoFilter,
  setCell,
} from '@office-kit/xlsx/worksheet';

const render = async (): Promise<Uint8Array> => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Leverage');
  appendRows(ws, [
    ['Language', 'Words'],
    ['de', 71_579],
    ['fr', 12_004],
  ]);
  setCell(ws, 4, 2, makeFormula('SUM(B2:B3)', { cachedValue: 83_583 }));
  setAutoFilter(ws, makeAutoFilter({ ref: 'A1:B3' }));
  return workbookToBytes(wb);
};

const wb = await loadWorkbook(fromArrayBuffer(await render()));

// getSheet narrows past the worksheet / chartsheet union for you.
const ws = getSheet(wb, 'Leverage');
if (!ws) throw new Error(`no Leverage sheet in ${sheetNames(wb).join(', ')}`);

console.log(getRangeValues(ws, 'A1:B3')); // [['Language','Words'], ['de',71579], ['fr',12004]]
console.log(getAutoFilter(ws)?.ref); // 'A1:B3'

// Every formula the renderer placed, keyed by address.
const formulas = new Map<string, string>();
for (const cell of iterCells(ws)) {
  const text = getFormulaText(cell);
  if (text !== undefined) formulas.set(getCoordinate(cell), text);
}
console.log(formulas); // Map { 'B4' => 'SUM(B2:B3)' }
