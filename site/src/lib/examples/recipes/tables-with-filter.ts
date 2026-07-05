// Promote a range to an Excel Table (named range with banded styling and
// a built-in filter dropdown on every header).

import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { addExcelTable, setCell } from '@office-kit/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Inventory');

const headers = ['SKU', 'Name', 'Price'];
headers.forEach((h, i) => setCell(ws, 1, i + 1, h));
const rows: ReadonlyArray<readonly [string, string, number]> = [
  ['A-001', 'Widget', 19.95],
  ['A-002', 'Gadget', 24.5],
  ['A-003', 'Doohickey', 7.25],
];
rows.forEach((row, r) => row.forEach((v, c) => setCell(ws, r + 2, c + 1, v)));

addExcelTable(wb, ws, {
  name: 'Inventory',
  ref: 'A1:C4',
  columns: headers,
  style: 'TableStyleMedium2',
});

await saveWorkbook(wb, toFile('inventory.xlsx'));
