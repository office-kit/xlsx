// Tests for SaveOptions.mtime / compressionLevel.
//
// ZIP has no "no timestamp" encoding, so fflate stamps the wall clock into each
// entry by default and two saves of the same workbook differ in bytes. Pinning
// mtime is what makes golden-file tests and content-addressed caching possible.

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { saveWorkbook, workbookToBytes } from '../../src/io/save.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import { appendRows, getCell } from '../../src/worksheet/worksheet.js';

const STAMP = new Date(Date.UTC(2026, 0, 2, 3, 4, 0));
const STAMP_ISO = STAMP.toISOString();

const buildReport = () => {
  const wb = createWorkbook();
  // Caller-controlled stamps: the wall clock has no business in a report built
  // from a payload.
  wb.properties = { creator: 'renderer', created: STAMP_ISO, modified: STAMP_ISO };
  const ws = addWorksheet(wb, 'Leverage');
  appendRows(ws, [
    ['Language', 'Words'],
    ['de', 71_579],
    ['fr', 12_004],
  ]);
  return wb;
};

describe('deterministic output', () => {
  it('a pinned mtime makes two saves byte-identical', async () => {
    const first = await workbookToBytes(buildReport(), { mtime: STAMP });
    const second = await workbookToBytes(buildReport(), { mtime: STAMP });
    expect(second).toEqual(first);
  });

  it('without a pinned mtime the bytes carry the wall clock', async () => {
    const pinned = await workbookToBytes(buildReport(), { mtime: STAMP });
    const unpinned = await workbookToBytes(buildReport());
    expect(unpinned).not.toEqual(pinned);
  });

  it('the pinned archive still loads', async () => {
    const bytes = await workbookToBytes(buildReport(), { mtime: STAMP });
    const wb = await loadWorkbook(fromBuffer(bytes));
    const ws = getSheet(wb, 'Leverage');
    expect(ws).toBeDefined();
    if (!ws) return;
    expect(getCell(ws, 2, 2)?.value).toBe(71_579);
    expect(wb.properties?.creator).toBe('renderer');
  });

  it('compressionLevel reaches the deflate stream', async () => {
    const stored = await workbookToBytes(buildReport(), { mtime: STAMP, compressionLevel: 0 });
    const squeezed = await workbookToBytes(buildReport(), { mtime: STAMP, compressionLevel: 9 });
    expect(stored.byteLength).toBeGreaterThan(squeezed.byteLength);
    // Level 0 still has to round-trip.
    const wb = await loadWorkbook(fromBuffer(stored));
    expect(getSheet(wb, 'Leverage')).toBeDefined();
  });

  it('saveWorkbook threads the options through to the sink', async () => {
    const sinkA = toBuffer();
    const sinkB = toBuffer();
    await saveWorkbook(buildReport(), sinkA, { mtime: STAMP });
    await saveWorkbook(buildReport(), sinkB, { mtime: STAMP });
    expect(sinkB.result()).toEqual(sinkA.result());
  });

  it('the streaming writer takes the same option', async () => {
    const render = async (): Promise<Buffer> => {
      const sink = toBuffer();
      const wb = await createWriteOnlyWorkbook(sink, { mtime: STAMP });
      const ws = await wb.addWorksheet('Data');
      await ws.appendRow(['de', 71_579]);
      await ws.close();
      await wb.finalize();
      return sink.result();
    };
    expect(await render()).toEqual(await render());
  });
});
