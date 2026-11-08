// The writer coalesces contiguous <col> entries whose emitted attributes are
// identical into a single range, so a uniform band of columns serializes as
// one <col> element instead of thousands of single-column ones. The
// worksheet's columnDimensions map is never rewritten — merging happens only
// at serialization time.

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook, type Workbook } from '../../src/workbook/workbook.js';
import type { ColumnDimension } from '../../src/worksheet/dimensions.js';
import {
  getColumnDimension,
  hideColumns,
  setCell,
  setColumnWidth,
  setColumnWidths,
} from '../../src/worksheet/index.js';

type Fields = Partial<Omit<ColumnDimension, 'min' | 'max'>>;

const worksheetPart = async (wb: Workbook): Promise<string> => {
  const part = unzipSync(await workbookToBytes(wb))['xl/worksheets/sheet1.xml'];
  if (!part) throw new Error('no worksheet part');
  return new TextDecoder().decode(part);
};

const colElements = (xml: string): string[] => [...xml.matchAll(/<col [^>]*\/>/g)].map((m) => m[0]);

const sheetWithCell = (): { wb: Workbook; ws: ReturnType<typeof addWorksheet> } => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  setCell(ws, 1, 1, 'x');
  return { wb, ws };
};

describe('serializing <cols>', () => {
  it('merges contiguous columns of identical width into one range', async () => {
    const { wb, ws } = sheetWithCell();
    setColumnWidths(ws, [0.3, 0.3, 0.3], 9);
    expect(colElements(await worksheetPart(wb))).toEqual([
      '<col min="9" max="11" width="0.3" customWidth="1"/>',
    ]);
  });

  it('merges a 2451-column band into a single element', async () => {
    const { wb, ws } = sheetWithCell();
    setColumnWidths(ws, new Array<number>(2451).fill(0.3), 9);
    expect(colElements(await worksheetPart(wb))).toEqual([
      '<col min="9" max="2459" width="0.3" customWidth="1"/>',
    ]);
  });

  it('starts a new range where the width changes', async () => {
    const { wb, ws } = sheetWithCell();
    setColumnWidths(ws, [12, 12, 13, 13]);
    expect(colElements(await worksheetPart(wb))).toEqual([
      '<col min="1" max="2" width="12" customWidth="1"/>',
      '<col min="3" max="4" width="13" customWidth="1"/>',
    ]);
  });

  it('leaves a gap between ranges', async () => {
    const { wb, ws } = sheetWithCell();
    setColumnWidth(ws, 1, 12);
    setColumnWidth(ws, 3, 12);
    expect(colElements(await worksheetPart(wb))).toEqual([
      '<col min="1" max="1" width="12" customWidth="1"/>',
      '<col min="3" max="3" width="12" customWidth="1"/>',
    ]);
  });

  it('merges a hidden band into one element at the default width', async () => {
    const { wb, ws } = sheetWithCell();
    hideColumns(ws, 3, 6);
    expect(colElements(await worksheetPart(wb))).toEqual([
      '<col min="3" max="6" width="9.140625" hidden="1"/>',
    ]);
  });

  // Two adjacent columns, identical except for one serialized property.
  const DIFFERENCES: ReadonlyArray<[name: string, left: Fields, right: Fields]> = [
    ['width', { width: 12 }, { width: 13 }],
    ['style', { width: 12 }, { width: 12, style: 5 }],
    ['hidden', { width: 12 }, { width: 12, hidden: true }],
    ['bestFit', { width: 12 }, { width: 12, bestFit: true }],
    ['customWidth', { customWidth: true }, {}],
    ['outlineLevel', { width: 12 }, { width: 12, outlineLevel: 2 }],
    ['collapsed', { width: 12 }, { width: 12, collapsed: true }],
  ];

  it.each(DIFFERENCES)('does not merge when %s differs', async (_name, left, right) => {
    const { wb, ws } = sheetWithCell();
    ws.columnDimensions.set(1, { min: 1, max: 1, ...left });
    ws.columnDimensions.set(2, { min: 2, max: 2, ...right });
    const els = colElements(await worksheetPart(wb));
    expect(els).toHaveLength(2);
    expect(els[0]).toMatch(/^<col min="1" max="1"/);
    expect(els[1]).toMatch(/^<col min="2" max="2"/);
  });

  // Raw fields that differ but serialize to the same attributes are one range.
  const EQUIVALENTS: ReadonlyArray<[name: string, left: Fields, right: Fields]> = [
    ['hidden omitted vs false', { width: 12 }, { width: 12, hidden: false }],
    ['bestFit omitted vs false', { width: 12 }, { width: 12, bestFit: false }],
    ['collapsed omitted vs false', { width: 12 }, { width: 12, collapsed: false }],
    ['width implies customWidth', { width: 12 }, { width: 12, customWidth: true }],
    ['default width spelled out', { customWidth: true }, { width: 9.140625, customWidth: true }],
  ];

  it.each(EQUIVALENTS)('merges when %s emits identical attributes', async (_name, left, right) => {
    const { wb, ws } = sheetWithCell();
    ws.columnDimensions.set(1, { min: 1, max: 1, ...left });
    ws.columnDimensions.set(2, { min: 2, max: 2, ...right });
    const els = colElements(await worksheetPart(wb));
    expect(els).toHaveLength(1);
    expect(els[0]).toMatch(/^<col min="1" max="2"/);
  });
});

describe('coalescing boundaries', () => {
  it('sorts and merges entries written out of column order', async () => {
    const { wb, ws } = sheetWithCell();
    setColumnWidth(ws, 3, 12);
    setColumnWidth(ws, 1, 12);
    setColumnWidth(ws, 2, 12);
    expect(colElements(await worksheetPart(wb))).toEqual([
      '<col min="1" max="3" width="12" customWidth="1"/>',
    ]);
  });

  it('merges authored range entries that abut', async () => {
    const { wb, ws } = sheetWithCell();
    ws.columnDimensions.set(1, { min: 1, max: 5, width: 12, customWidth: true });
    ws.columnDimensions.set(6, { min: 6, max: 10, width: 12, customWidth: true });
    expect(colElements(await worksheetPart(wb))).toEqual([
      '<col min="1" max="10" width="12" customWidth="1"/>',
    ]);
  });

  it('keeps overlapping authored runs separate', async () => {
    const { wb, ws } = sheetWithCell();
    ws.columnDimensions.set(1, { min: 1, max: 10, width: 12, customWidth: true });
    ws.columnDimensions.set(5, { min: 5, max: 15, width: 12, customWidth: true });
    expect(colElements(await worksheetPart(wb))).toEqual([
      '<col min="1" max="10" width="12" customWidth="1"/>',
      '<col min="5" max="15" width="12" customWidth="1"/>',
    ]);
  });

  it('does not rewrite the worksheet map', async () => {
    const { wb, ws } = sheetWithCell();
    setColumnWidths(ws, [12, 12, 13]);
    const before = [...ws.columnDimensions.values()];
    await workbookToBytes(wb);
    const after = [...ws.columnDimensions.values()];
    expect(after).toHaveLength(3);
    expect(after).toStrictEqual(before);
    for (let i = 0; i < before.length; i++) {
      expect(after[i]).toBe(before[i]); // entry objects not replaced either
    }
  });
});

describe('a coalesced file reloaded', () => {
  it('resolves every column in the band to the same width', async () => {
    const { wb, ws } = sheetWithCell();
    setColumnWidths(ws, new Array<number>(2451).fill(0.3), 9);
    const wb2 = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    const loaded = wb2.sheets[0];
    if (!loaded || loaded.kind !== 'worksheet') throw new Error('expected worksheet');
    const entries = [...loaded.sheet.columnDimensions.values()];
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({ min: 9, max: 2459, width: 0.3 });
    expect(getColumnDimension(loaded.sheet, 9)?.width).toBe(0.3);
    expect(getColumnDimension(loaded.sheet, 1500)?.width).toBe(0.3);
    expect(getColumnDimension(loaded.sheet, 2459)?.width).toBe(0.3);
    expect(getColumnDimension(loaded.sheet, 8)).toBeUndefined();
  });
});
