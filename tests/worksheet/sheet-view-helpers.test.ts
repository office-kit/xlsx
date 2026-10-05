// Tests for sheet view + tab-color ergonomic helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import type { Worksheet } from '../../src/worksheet/worksheet.js';
import {
  removeSheetTabColor,
  setRightToLeft,
  setSheetTabColor,
  setSheetViewMode,
  setSheetZoom,
  setShowFormulas,
  setShowGridLines,
  setShowRowColHeaders,
  setShowZeros,
} from '../../src/worksheet/worksheet.js';

describe('setSheetTabColor', () => {
  it('hex string assigns rgb color via sheetProperties', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSheetTabColor(ws, 'FF0070C0');
    expect(ws.sheetProperties?.tabColor?.rgb).toBe('FF0070C0');
  });

  it('Color partial assigns theme + tint', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSheetTabColor(ws, { theme: 4, tint: 0.4 });
    expect(ws.sheetProperties?.tabColor?.theme).toBe(4);
    expect(ws.sheetProperties?.tabColor?.tint).toBeCloseTo(0.4);
  });

  it('removeSheetTabColor drops only tabColor and preserves siblings', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSheetTabColor(ws, 'FF112233');
    if (!ws.sheetProperties) throw new Error('sheetProperties was not set');
    ws.sheetProperties.codeName = 'Sheet1';
    removeSheetTabColor(ws);
    expect(ws.sheetProperties.tabColor).toBeUndefined();
    expect(ws.sheetProperties.codeName).toBe('Sheet1');
  });

  it('removeSheetTabColor on a worksheet with no sheetProperties is a no-op', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    removeSheetTabColor(ws);
    expect(ws.sheetProperties).toBeUndefined();
  });
});

describe('sheet view toggles', () => {
  it('setShowGridLines / RowColHeaders / Formulas / Zeros / RightToLeft set primary view fields', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setShowGridLines(ws, false);
    setShowRowColHeaders(ws, false);
    setShowFormulas(ws, true);
    setShowZeros(ws, false);
    setRightToLeft(ws, true);
    const v = ws.views[0];
    expect(v?.showGridLines).toBe(false);
    expect(v?.showRowColHeaders).toBe(false);
    expect(v?.showFormulas).toBe(true);
    expect(v?.showZeros).toBe(false);
    expect(v?.rightToLeft).toBe(true);
  });

  it('lazily creates the primary view if absent', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(ws.views.length).toBe(0);
    setShowGridLines(ws, false);
    expect(ws.views.length).toBe(1);
  });

  it('setSheetZoom sets zoomScale on the primary view', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSheetZoom(ws, 125);
    expect(ws.views[0]?.zoomScale).toBe(125);
  });

  it('setSheetZoom rejects out-of-range or non-integer scale', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(() => setSheetZoom(ws, 9)).toThrow(/in \[10, 400\]/);
    expect(() => setSheetZoom(ws, 401)).toThrow(/in \[10, 400\]/);
    expect(() => setSheetZoom(ws, 100.5)).toThrow(/in \[10, 400\]/);
  });

  it('setSheetViewMode set page-break-preview / page-layout / normal', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSheetViewMode(ws, 'pageBreakPreview');
    expect(ws.views[0]?.view).toBe('pageBreakPreview');
    setSheetViewMode(ws, 'pageLayout');
    expect(ws.views[0]?.view).toBe('pageLayout');
    setSheetViewMode(ws, 'normal');
    expect(ws.views[0]?.view).toBe('normal');
  });
});

describe('sheet view + tab color round-trip', () => {
  it('all toggles + tab color survive save → load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'View');
    setSheetTabColor(ws, 'FFAA3300');
    setShowGridLines(ws, false);
    setShowRowColHeaders(ws, false);
    setSheetZoom(ws, 150);
    setSheetViewMode(ws, 'pageLayout');
    setRightToLeft(ws, true);
    const primary = ws.views[0];
    if (!primary) throw new Error('expected a sheet view');
    primary.showOutlineSymbols = false;

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb2.sheets[0]?.sheet;
    if (!sheet || !('rows' in sheet)) throw new Error('expected worksheet');
    const ws2 = sheet as Worksheet;
    expect(ws2.sheetProperties?.tabColor?.rgb).toBe('FFAA3300');
    expect(ws2.views[0]?.showGridLines).toBe(false);
    expect(ws2.views[0]?.showRowColHeaders).toBe(false);
    expect(ws2.views[0]?.zoomScale).toBe(150);
    expect(ws2.views[0]?.view).toBe('pageLayout');
    expect(ws2.views[0]?.rightToLeft).toBe(true);
    expect(ws2.views[0]?.showOutlineSymbols).toBe(false);
  });
});
