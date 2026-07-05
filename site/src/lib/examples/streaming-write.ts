// Stream millions of rows to disk in a fixed memory budget. Each row is
// deflated as it arrives — no intermediate workbook in memory.

import { toFile } from '@office-kit/xlsx/node';
import { createWriteOnlyWorkbook } from '@office-kit/xlsx/streaming';

const sink = toFile('big.xlsx');
const wb = await createWriteOnlyWorkbook(sink);
const ws = await wb.addWorksheet('Data');
ws.setColumnWidth(1, 24); // must precede the first appendRow

for (let r = 0; r < 10_000_000; r++) {
  await ws.appendRow([r, `row-${r}`, r * Math.PI]);
}
await ws.close();
await wb.finalize();
