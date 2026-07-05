// README smoke tests — runs each documented example end-to-end against
// the actual public API surface. The examples in README.md are the
// project's first impression and the most-likely-to-rot piece of
// documentation; this file pins them so a renamed export, removed
// helper, or shifted return shape fails CI rather than the user's
// first hour with the library.

import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

let scratch: string;
beforeAll(() => {
  scratch = mkdtempSync(join(tmpdir(), 'openxml-readme-'));
});
afterAll(() => {
  rmSync(scratch, { recursive: true, force: true });
});

describe('README — full lib read+edit+write', () => {
  it('matches the documented public API surface for the round-trip example', async () => {
    const io = await import('../../src/io/index');
    const workbook = await import('../../src/workbook/index');
    const worksheet = await import('../../src/worksheet/index');
    const node = await import('../../src/node');
    expect(typeof node.fromBuffer).toBe('function');
    expect(typeof io.loadWorkbook).toBe('function');
    expect(typeof io.workbookToBytes).toBe('function');
    expect(typeof worksheet.setCell).toBe('function');
    expect(typeof workbook.createWorkbook).toBe('function');
    expect(typeof workbook.addWorksheet).toBe('function');

    // Build a synthetic workbook, edit a cell, round-trip.
    const wb = workbook.createWorkbook();
    const ws = workbook.addWorksheet(wb, 'Sheet1');
    worksheet.setCell(ws, 1, 1, 'Hello from @office-kit/xlsx');
    const bytes = await io.workbookToBytes(wb);
    const wb2 = await io.loadWorkbook(node.fromBuffer(bytes));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(ref0.sheet.rows.get(1)?.get(1)?.value).toBe('Hello from @office-kit/xlsx');
  });
});

describe('README — Node fromFile / toFile / saveWorkbook', () => {
  it('exposes the documented Node-only public surface', async () => {
    const io = await import('../../src/io/index');
    const node = await import('../../src/node');
    const workbook = await import('../../src/workbook/index');
    expect(typeof node.fromFile).toBe('function');
    expect(typeof io.loadWorkbook).toBe('function');
    expect(typeof io.saveWorkbook).toBe('function');
    expect(typeof node.toFile).toBe('function');

    const out = join(scratch, 'output.xlsx');
    const wb = workbook.createWorkbook();
    workbook.addWorksheet(wb, 'X');
    await io.saveWorkbook(wb, node.toFile(out));
    const wb2 = await io.loadWorkbook(node.fromFile(out));
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(['X']);
  });
});

describe('README — browser fromResponse', () => {
  it('exposes the documented `import { fromResponse } from @office-kit/xlsx/io` shape', async () => {
    const ioBrowser = await import('../../src/io/index');
    expect(typeof ioBrowser.fromResponse).toBe('function');

    // Build a Response-shaped object backed by a real Uint8Array of
    // a synthetic xlsx and feed it through loadWorkbook.
    const io = await import('../../src/io/index');
    const workbook = await import('../../src/workbook/index');
    const wb = workbook.createWorkbook();
    workbook.addWorksheet(wb, 'FromFetch');
    const bytes = await io.workbookToBytes(wb);

    // Wrap the bytes in a Blob so the Response body typechecks under
    // Node's tightened @types/node BodyInit definition. The cast guides
    // TS past the Uint8Array<ArrayBufferLike> → BlobPart mismatch
    // recently introduced by @types/node (the runtime accepts it).
    const response = new Response(new Blob([bytes as unknown as BlobPart]));
    const src = ioBrowser.fromResponse(response);
    const wb2 = await io.loadWorkbook(src);
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(['FromFetch']);
  });
});

describe('README — streaming write (createWriteOnlyWorkbook)', () => {
  it('matches the documented appendRow + finalize flow', async () => {
    const streaming = await import('../../src/streaming/index');
    const io = await import('../../src/io/index');
    const node = await import('../../src/node');
    expect(typeof streaming.createWriteOnlyWorkbook).toBe('function');

    const out = join(scratch, 'big.xlsx');
    const sink = node.toFile(out);
    const wb = await streaming.createWriteOnlyWorkbook(sink);
    const ws = await wb.addWorksheet('Data');
    ws.setColumnWidth(1, 24); // must precede the first appendRow
    for (let r = 0; r < 100; r++) {
      await ws.appendRow([r, `row-${r}`, r * 0.5]);
    }
    await ws.close();
    await wb.finalize();

    const wb2 = await io.loadWorkbook(node.fromFile(out));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(ref0.sheet.rows.size).toBe(100);
  });
});

describe('README — streaming read (loadWorkbookStream + iterRows)', () => {
  it('matches the documented openWorksheet + iterRows + close flow', async () => {
    const streaming = await import('../../src/streaming/index');
    const io = await import('../../src/io/index');
    const node = await import('../../src/node');
    const workbook = await import('../../src/workbook/index');
    const worksheet = await import('../../src/worksheet/index');

    expect(typeof streaming.loadWorkbookStream).toBe('function');

    // Build a workbook to stream through.
    const wb = workbook.createWorkbook();
    const ws = workbook.addWorksheet(wb, 'Big');
    for (let r = 1; r <= 50; r++) worksheet.setCell(ws, r, 1, r);
    const out = join(scratch, 'stream.xlsx');
    await io.saveWorkbook(wb, node.toFile(out));

    const wbS = await streaming.loadWorkbookStream(node.fromFile(out));
    const sheet = wbS.openWorksheet(wbS.sheetNames[0] ?? '');
    const seen: number[] = [];
    for await (const row of sheet.iterRows({ minRow: 1, maxRow: 5 })) {
      const v = row[0]?.value;
      if (typeof v === 'number') seen.push(v);
    }
    expect(seen).toEqual([1, 2, 3, 4, 5]);
    await wbS.close();
  });
});
