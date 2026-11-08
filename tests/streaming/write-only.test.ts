// Phase 4 §3 write-only streaming acceptance.

import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { getColumnDimension } from '../../src/worksheet/index.js';
import { iterRows } from '../../src/worksheet/worksheet.js';

const collect = async (
  fn: (cb: (sink: ReturnType<typeof toBuffer>) => Promise<void>) => Promise<void>,
): Promise<Uint8Array> => {
  const sink = toBuffer();
  await fn(async () => {});
  return sink.result();
};

const writeWorkbook = async (
  build: (
    wb: Awaited<ReturnType<typeof createWriteOnlyWorkbook>>,
  ) => Promise<void>,
): Promise<Uint8Array> => {
  const sink = toBuffer();
  const wb = await createWriteOnlyWorkbook(sink);
  await build(wb);
  await wb.finalize();
  return sink.result();
};

describe('createWriteOnlyWorkbook — basic round-trip', () => {
  it('emits a valid xlsx that loadWorkbook can re-parse', async () => {
    const bytes = await writeWorkbook(async (wb) => {
      const ws = await wb.addWorksheet('Sheet1');
      await ws.appendRow([1, 2, 3]);
      await ws.appendRow(['a', 'b', 'c']);
      await ws.appendRow([true, false, null]);
      await ws.close();
    });
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.length).toBe(1);
    const ws2 = wb2.sheets[0];
    if (!ws2 || ws2.kind !== 'worksheet') throw new Error('expected worksheet');
    const rows: unknown[][] = [];
    for (const cells of iterRows(ws2.sheet)) rows.push(cells.map((c) => c?.value ?? null));
    expect(rows).toEqual([
      [1, 2, 3],
      ['a', 'b', 'c'],
      // Row 3 col 3 is null and write-only skips emitting null cells, so the
      // rectangular iter pads it with null to match the bounding-box width.
      [true, false, null],
    ]);
  });

  it('preserves multiple worksheets in declaration order', async () => {
    const bytes = await writeWorkbook(async (wb) => {
      const a = await wb.addWorksheet('Alpha');
      await a.appendRow([1]);
      await a.close();
      const b = await wb.addWorksheet('Beta');
      await b.appendRow([2]);
      await b.close();
    });
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.sheets.map((s) => s.sheet.title)).toEqual(['Alpha', 'Beta']);
  });
});

describe('createWriteOnlyWorkbook — sequencing constraints', () => {
  it('rejects addWorksheet while the previous sheet is still open', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    const a = await wb.addWorksheet('A');
    await a.appendRow([1]);
    await expect(wb.addWorksheet('B')).rejects.toThrowError(/previous worksheet still open/);
    await a.close();
    const b = await wb.addWorksheet('B');
    await b.close();
    await wb.finalize();
  });

  it('rejects appendRow on a closed worksheet', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    const a = await wb.addWorksheet('A');
    await a.close();
    await expect(a.appendRow([1])).rejects.toThrowError(/already closed/);
    await wb.finalize();
  });

  it('rejects finalize while a worksheet is still open', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    await wb.addWorksheet('A');
    await expect(wb.finalize()).rejects.toThrowError(/still open/);
  });

  it('rejects double-finalize', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    const a = await wb.addWorksheet('A');
    await a.close();
    await wb.finalize();
    await expect(wb.finalize()).rejects.toThrowError(/already finalised/);
  });

  it('rejects duplicate worksheet titles', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    const a = await wb.addWorksheet('Same');
    await a.close();
    await expect(wb.addWorksheet('Same')).rejects.toThrowError(/already in use/);
  });

  it('rejects duplicate worksheet titles case-insensitively', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    const a = await wb.addWorksheet('Data');
    await a.close();
    await expect(wb.addWorksheet('data')).rejects.toThrowError(/already in use/);
    await expect(wb.addWorksheet('DATA')).rejects.toThrowError(/already in use/);
  });

  it('rejects forbidden characters in the title', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    await expect(wb.addWorksheet('Bad/Name')).rejects.toThrowError(/must not contain/);
    await expect(wb.addWorksheet('Bad\\Name')).rejects.toThrowError(/must not contain/);
    await expect(wb.addWorksheet('Bad:Name')).rejects.toThrowError(/must not contain/);
    await expect(wb.addWorksheet('Bad?Name')).rejects.toThrowError(/must not contain/);
    await expect(wb.addWorksheet('Bad*Name')).rejects.toThrowError(/must not contain/);
    await expect(wb.addWorksheet('Bad[1]')).rejects.toThrowError(/must not contain/);
  });

  it('rejects leading or trailing apostrophe in the title', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    await expect(wb.addWorksheet("'Leading")).rejects.toThrowError(/apostrophe/);
    await expect(wb.addWorksheet("Trailing'")).rejects.toThrowError(/apostrophe/);
  });

  it('rejects the reserved name "History" case-insensitively', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    await expect(wb.addWorksheet('History')).rejects.toThrowError(/reserved/);
    await expect(wb.addWorksheet('history')).rejects.toThrowError(/reserved/);
    await expect(wb.addWorksheet('HISTORY')).rejects.toThrowError(/reserved/);
  });
});

describe('createWriteOnlyWorkbook — styles + sharedStrings', () => {
  it('appendRow {value, style} entries dedup into the cellXfs pool', async () => {
    const bytes = await writeWorkbook(async (wb) => {
      const ws = await wb.addWorksheet('Styled');
      await ws.appendRow([
        { value: 'bold', style: { font: { bold: true } } },
        { value: 'plain' },
        { value: 'bold-too', style: { font: { bold: true } } },
      ]);
      await ws.close();
    });
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws = wb2.sheets[0];
    if (!ws || ws.kind !== 'worksheet') throw new Error('expected worksheet');
    const rows: { value: unknown; styleId: number }[][] = [];
    for (const cells of iterRows(ws.sheet)) {
      rows.push(cells.map((c) => (c === undefined ? { value: null, styleId: 0 } : { value: c.value, styleId: c.styleId })));
    }
    // Two cells share the same styleId; the unstyled cell uses 0.
    const styled1 = rows[0]?.[0];
    const styled2 = rows[0]?.[2];
    const plain = rows[0]?.[1];
    if (!styled1 || !styled2 || !plain) throw new Error('row malformed');
    expect(styled1.styleId).toBe(styled2.styleId);
    expect(styled1.styleId).not.toBe(plain.styleId);
  });

  it('repeated string values reuse a single sharedStrings slot', async () => {
    const bytes = await writeWorkbook(async (wb) => {
      const ws = await wb.addWorksheet('Strings');
      await ws.appendRow(['hi', 'hi', 'bye', 'hi']);
      await ws.close();
    });
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    const sst = new TextDecoder().decode(entries['xl/sharedStrings.xml']);
    // Implementation reports unique count both as `count` and
    // `uniqueCount` because that's what Excel tolerates and the
    // existing serializer emits. Either way, "hi" is interned to a
    // single <si> entry.
    expect((sst.match(/<si>/g) ?? []).length).toBe(2);
    expect(sst).toContain('uniqueCount="2"');
  });

  it('omits sharedStrings part when no strings are written', async () => {
    const bytes = await writeWorkbook(async (wb) => {
      const ws = await wb.addWorksheet('Nums');
      await ws.appendRow([1, 2, 3]);
      await ws.close();
    });
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    expect(entries['xl/sharedStrings.xml']).toBeUndefined();
  });
});

describe('createWriteOnlyWorkbook — column widths', () => {
  it('setColumnWidth survives the round-trip', async () => {
    const bytes = await writeWorkbook(async (wb) => {
      const ws = await wb.addWorksheet('Wide');
      ws.setColumnWidth(1, 25.5);
      ws.setColumnWidth(3, 10);
      await ws.appendRow(['a', 'b', 'c']);
      await ws.close();
    });
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws = wb2.sheets[0];
    if (!ws || ws.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(ws.sheet.columnDimensions.get(1)?.width).toBe(25.5);
    expect(ws.sheet.columnDimensions.get(3)?.width).toBe(10);
  });

  it('coalesces contiguous equal widths into ranges', async () => {
    const bytes = await writeWorkbook(async (wb) => {
      const ws = await wb.addWorksheet('Wide');
      ws.setColumnWidth(1, 25.5);
      ws.setColumnWidth(2, 25.5);
      ws.setColumnWidth(4, 10);
      await ws.appendRow(['a', 'b', 'c']);
      await ws.close();
    });
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws = wb2.sheets[0];
    if (!ws || ws.kind !== 'worksheet') throw new Error('expected worksheet');
    const runs = [...ws.sheet.columnDimensions.values()].map((d) => [d.min, d.max, d.width]);
    expect(runs).toEqual([
      [1, 2, 25.5],
      [4, 4, 10],
    ]);
    expect(getColumnDimension(ws.sheet, 2)?.width).toBe(25.5);
  });

  it('keeps a lone zero-width column 1', async () => {
    const bytes = await writeWorkbook(async (wb) => {
      const ws = await wb.addWorksheet('Wide');
      ws.setColumnWidth(1, 0);
      await ws.appendRow(['a', 'b', 'c']);
      await ws.close();
    });
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws = wb2.sheets[0];
    if (!ws || ws.kind !== 'worksheet') throw new Error('expected worksheet');
    expect(getColumnDimension(ws.sheet, 1)?.width).toBe(0);
  });

  it('keeps a zero-width column 1 when a different width follows', async () => {
    const bytes = await writeWorkbook(async (wb) => {
      const ws = await wb.addWorksheet('Wide');
      ws.setColumnWidth(1, 0);
      ws.setColumnWidth(2, 5);
      await ws.appendRow(['a', 'b', 'c']);
      await ws.close();
    });
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws = wb2.sheets[0];
    if (!ws || ws.kind !== 'worksheet') throw new Error('expected worksheet');
    const runs = [...ws.sheet.columnDimensions.values()].map((d) => [d.min, d.max, d.width]);
    expect(runs).toEqual([
      [1, 1, 0],
      [2, 2, 5],
    ]);
  });
});

// Suppress unused import lint.
void collect;
