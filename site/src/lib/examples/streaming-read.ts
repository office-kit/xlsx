// Iterate huge sheets without loading the full workbook. iterRows is a SAX
// pass — it walks the file once and yields each row's cells.

import { fromFile } from '@office-kit/xlsx/node';
import { loadWorkbookStream } from '@office-kit/xlsx/streaming';

const wb = await loadWorkbookStream(fromFile('big.xlsx'));
const sheet = wb.openWorksheet(wb.sheetNames[0] ?? '');
for await (const row of sheet.iterRows({ minRow: 1, maxRow: 100 })) {
  console.log(row.map((c) => c.value));
}
await wb.close();
