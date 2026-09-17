// `loadWorkbookStream` and `loadWorkbook` have to answer the same bytes the
// same way. On the post-2018 error tokens and on `t="d"` they did the opposite
// of each other: the streaming reader returned `null` or a NaN, so the cell
// read as empty, while the DOM reader threw.

import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream, type ReadOnlyWorksheet } from '../../src/streaming/read-only.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';

const SHEET_PART = 'xl/worksheets/sheet1.xml';

/** Four numeric rows, with row 2's `<c>` replaced by `cellXml`. */
const savedWithCell = async (cellXml: string): Promise<Uint8Array> => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Data');
  for (let row = 1; row <= 4; row++) setCell(ws, row, 1, row * 10);
  const entries = unzipSync(await workbookToBytes(wb));
  const part = entries[SHEET_PART];
  if (part === undefined) throw new Error(`save produced no ${SHEET_PART}`);
  const patched = new TextDecoder().decode(part).replace('<c r="A2"><v>20</v></c>', cellXml);
  entries[SHEET_PART] = new TextEncoder().encode(patched);
  return zipSync(entries);
};

const drain = async (ws: ReadOnlyWorksheet): Promise<unknown[][]> => {
  const rows: unknown[][] = [];
  for await (const row of ws.iterValues()) rows.push(row);
  return rows;
};

const streamValues = async (cellXml: string): Promise<unknown[][]> => {
  const wb = await loadWorkbookStream(fromBuffer(await savedWithCell(cellXml)));
  try {
    return await drain(wb.openWorksheet('Data'));
  } finally {
    await wb.close();
  }
};

const streamError = async (cellXml: string): Promise<Error> => {
  const wb = await loadWorkbookStream(fromBuffer(await savedWithCell(cellXml)));
  try {
    await drain(wb.openWorksheet('Data'));
  } catch (err) {
    return err as Error;
  } finally {
    await wb.close();
  }
  throw new Error('expected the stream to throw');
};

describe('loadWorkbookStream on a post-2018 error token', () => {
  it('yields the error value instead of an empty cell', async () => {
    expect(await streamValues('<c r="A2" t="e"><v>#SPILL!</v></c>')).toEqual([
      [10],
      [{ kind: 'error', code: '#SPILL!' }],
      [30],
      [40],
    ]);
  });

  it('keeps a token the library does not model', async () => {
    expect(await streamValues('<c r="A2" t="e"><v>#NOTYET!</v></c>')).toEqual([
      [10],
      [{ kind: 'error', code: '#NOTYET!' }],
      [30],
      [40],
    ]);
  });

  it('agrees with loadWorkbook that a non-token t="e" payload is an error', async () => {
    const bytes = await savedWithCell('<c r="A2" t="e"><v>oops</v></c>');
    await expect(loadWorkbook(fromBuffer(bytes))).rejects.toThrow(
      'worksheet: <v>oops</v> at Data!A2 is not an Excel error token',
    );
    const err = await streamError('<c r="A2" t="e"><v>oops</v></c>');
    expect(err).toBeInstanceOf(OpenXmlSchemaError);
    expect(err.message).toBe('worksheet: <v>oops</v> at Data!A2 is not an Excel error token');
  });
});

describe('loadWorkbookStream on t="d"', () => {
  it('reads the date instead of parsing the text as a number', async () => {
    expect(await streamValues('<c r="A2" t="d"><v>2024-03-14T00:00:00</v></c>')).toEqual([
      [10],
      [new Date(Date.UTC(2024, 2, 14))],
      [30],
      [40],
    ]);
  });

  it('reads the same Date loadWorkbook reads', async () => {
    const bytes = await savedWithCell('<c r="A2" t="d"><v>2024-03-14T12:30:45Z</v></c>');
    const wb = await loadWorkbook(fromBuffer(bytes));
    const ws = getSheet(wb, 'Data');
    if (ws === undefined) throw new Error('load produced no worksheet');
    expect(getCell(ws, 2, 1)?.value).toEqual(new Date(Date.UTC(2024, 2, 14, 12, 30, 45)));
    expect(await streamValues('<c r="A2" t="d"><v>2024-03-14T12:30:45Z</v></c>')).toEqual([
      [10],
      [new Date(Date.UTC(2024, 2, 14, 12, 30, 45))],
      [30],
      [40],
    ]);
  });

  it('throws on a date that does not exist, as loadWorkbook does', async () => {
    const err = await streamError('<c r="A2" t="d"><v>2024-02-30</v></c>');
    expect(err.message).toBe('worksheet: <v>2024-02-30</v> at Data!A2 is not an ISO 8601 date, time or duration');
  });
});

describe('loadWorkbookStream on a cell type outside ST_CellType', () => {
  it('throws instead of coercing the text with parseFloat', async () => {
    const bytes = await savedWithCell('<c r="A2" t="q"><v>oops</v></c>');
    await expect(loadWorkbook(fromBuffer(bytes))).rejects.toThrow('worksheet: unknown cell type t="q"');
    const err = await streamError('<c r="A2" t="q"><v>oops</v></c>');
    expect(err).toBeInstanceOf(OpenXmlSchemaError);
    expect(err.message).toBe('worksheet: unknown cell type t="q" at Data!A2');
  });
});
