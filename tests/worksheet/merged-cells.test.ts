import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  getCell,
  getMergedCells,
  isMergedCell,
  mergeCells,
  setCell,
  unmergeCells,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('mergeCells / unmergeCells', () => {
  it('starts empty and accepts a string ref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'M');
    expect(getMergedCells(ws).length).toBe(0);
    mergeCells(ws, 'A1:B2');
    expect(getMergedCells(ws).map((r) => `${r.minRow}:${r.minCol}-${r.maxRow}:${r.maxCol}`)).toEqual(['1:1-2:2']);
  });

  it('drops cells in the merged range except the top-left', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'M');
    setCell(ws, 1, 1, 'top-left');
    setCell(ws, 1, 2, 'gone');
    setCell(ws, 2, 1, 'gone-too');
    setCell(ws, 2, 2, 'also-gone');
    mergeCells(ws, 'A1:B2');
    expect(getCell(ws, 1, 1)?.value).toBe('top-left');
    expect(getCell(ws, 1, 2)).toBeUndefined();
    expect(getCell(ws, 2, 1)).toBeUndefined();
    expect(getCell(ws, 2, 2)).toBeUndefined();
  });

  it('rejects an overlapping merge', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'M');
    mergeCells(ws, 'A1:B2');
    expect(() => mergeCells(ws, 'B2:C3')).toThrowError(/overlaps/);
  });

  it('idempotent — re-merging the same range is a no-op', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'M');
    mergeCells(ws, 'A1:B2');
    mergeCells(ws, 'A1:B2');
    expect(getMergedCells(ws).length).toBe(1);
  });

  it('drops cells in a whole-column band except the top-left', () => {
    // The walk picks whichever axis is smaller to enumerate, so a band far
    // larger than the sheet takes the sparse path. Same outcome as a small
    // rectangle, which is what this pins.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'keep');
    setCell(ws, 5, 3, 'inside');
    setCell(ws, 7, 11, 'outside the column band');
    mergeCells(ws, 'A:J');
    expect(getCell(ws, 1, 1)?.value).toBe('keep');
    expect(getCell(ws, 5, 3)).toBeUndefined();
    expect(getCell(ws, 7, 11)?.value).toBe('outside the column band');
  });

  it('prunes a row that the merge empties', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'keep');
    setCell(ws, 4, 2, 'goes');
    mergeCells(ws, 'A1:C9');
    expect(ws.rows.has(4)).toBe(false);
    expect(ws.rows.has(1)).toBe(true);
  });

  it('unmergeCells drops the matching range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'M');
    mergeCells(ws, 'A1:B2');
    mergeCells(ws, 'D1:E5');
    expect(unmergeCells(ws, 'A1:B2')).toBe(true);
    expect(getMergedCells(ws).length).toBe(1);
    expect(unmergeCells(ws, 'A1:B2')).toBe(false);
  });

  it('isMergedCell covers every coord in the range', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'M');
    mergeCells(ws, 'B2:C3');
    expect(isMergedCell(ws, 2, 2)).toBe(true);
    expect(isMergedCell(ws, 3, 3)).toBe(true);
    expect(isMergedCell(ws, 1, 1)).toBe(false);
    expect(isMergedCell(ws, 4, 4)).toBe(false);
  });
});

describe('mergedCells round-trip through saveWorkbook → loadWorkbook', () => {
  it('preserves merged ranges across save / load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Merged');
    setCell(ws, 1, 1, 'header');
    mergeCells(ws, 'A1:C1');
    setCell(ws, 3, 1, 'tally');
    mergeCells(ws, 'A3:B5');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);

    const refs = getMergedCells(ws2).map((r) => `${r.minRow}:${r.minCol}-${r.maxRow}:${r.maxCol}`);
    expect(refs.sort()).toEqual(['1:1-1:3', '3:1-5:2']);
    expect(getCell(ws2, 1, 1)?.value).toBe('header');
    expect(getCell(ws2, 3, 1)?.value).toBe('tally');
  });

  it('omits <mergeCells> when nothing is merged', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'NoMerge');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(getMergedCells(expectSheet(wb2.sheets[0]?.sheet)).length).toBe(0);
  });
});
