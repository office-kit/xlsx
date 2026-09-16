// Byte-identical output for identical input.
//
// Two things in an xlsx move on their own: the per-entry ZIP timestamp and the
// core properties. Pin both and the same payload always renders the same bytes,
// which is what a golden-file test or a content-addressed cache needs.

import { workbookToBytes } from '@office-kit/xlsx/io';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { appendRows } from '@office-kit/xlsx/worksheet';

interface Payload {
  readonly generatedAt: string;
  readonly rows: ReadonlyArray<readonly [string, number]>;
}

const render = async (payload: Payload): Promise<Uint8Array> => {
  const wb = createWorkbook();
  // Stamps come from the payload, never from the clock.
  wb.properties = {
    creator: 'leverage-report',
    created: payload.generatedAt,
    modified: payload.generatedAt,
  };
  const ws = addWorksheet(wb, 'Leverage');
  appendRows(ws, [['Language', 'Words'], ...payload.rows.map((r) => [...r])]);
  return workbookToBytes(wb, { mtime: new Date(payload.generatedAt) });
};

const payload: Payload = {
  generatedAt: '2026-01-02T03:04:00.000Z',
  rows: [
    ['de', 71_579],
    ['fr', 12_004],
  ],
};

const first = await render(payload);
const second = await render(payload);
console.log(Buffer.compare(Buffer.from(first), Buffer.from(second)) === 0); // true
