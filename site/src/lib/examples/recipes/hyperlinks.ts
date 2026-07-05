// Make a cell clickable. The text is whatever you set on the cell;
// hyperlink wires up the URL underneath.

import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { setCell, setHyperlink } from '@office-kit/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Links');

setCell(ws, 1, 1, 'Project home');
setHyperlink(ws, 'A1', {
  target: 'https://github.com/office-kit/xlsx',
  tooltip: 'View on GitHub',
});

await saveWorkbook(wb, toFile('with-links.xlsx'));
