// Tests for setActiveCell / setSelectedRange.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { unzipSync } from 'fflate';
import type { Chartsheet } from '../../src/chartsheet/chartsheet.js';
import { activeSelection } from '../../src/worksheet/views.js';
import {
  setActiveCell,
  setFreezePanes,
  setSelectedRange,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';

describe('setActiveCell', () => {
  it('lazily creates a Selection on the primary view', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setActiveCell(ws, 'B5');
    expect(ws.views[0]?.selections?.[0]?.activeCell).toBe('B5');
    expect(ws.views[0]?.selections?.[0]?.sqref).toBe('B5');
  });

  it('updates activeCell + sqref together when sqref tracked the previous activeCell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setActiveCell(ws, 'A1');
    setActiveCell(ws, 'C3');
    expect(ws.views[0]?.selections?.[0]?.activeCell).toBe('C3');
    expect(ws.views[0]?.selections?.[0]?.sqref).toBe('C3');
  });

  it('preserves an explicitly-set sqref', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSelectedRange(ws, 'A1:D10');
    setActiveCell(ws, 'B2');
    expect(ws.views[0]?.selections?.[0]?.activeCell).toBe('B2');
    expect(ws.views[0]?.selections?.[0]?.sqref).toBe('A1:D10');
  });
});

describe('setSelectedRange', () => {
  it('single range sets sqref + derives activeCell from top-left when missing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSelectedRange(ws, 'B2:D5');
    expect(ws.views[0]?.selections?.[0]?.sqref).toBe('B2:D5');
    expect(ws.views[0]?.selections?.[0]?.activeCell).toBe('B2');
  });

  it('multi-range sqref keeps the first ref as activeCell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setSelectedRange(ws, 'A1 C3:D4');
    expect(ws.views[0]?.selections?.[0]?.activeCell).toBe('A1');
    expect(ws.views[0]?.selections?.[0]?.sqref).toBe('A1 C3:D4');
  });

  it('does not overwrite an existing activeCell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setActiveCell(ws, 'E10');
    setSelectedRange(ws, 'A1:B2');
    expect(ws.views[0]?.selections?.[0]?.activeCell).toBe('E10');
    expect(ws.views[0]?.selections?.[0]?.sqref).toBe('A1:B2');
  });
});

describe('selection round-trip', () => {
  it('selection survives save → load', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sel');
    setSelectedRange(ws, 'A1:C5');
    setActiveCell(ws, 'B3');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb2.sheets[0]?.sheet;
    if (!sheet || !('rows' in sheet)) throw new Error('expected worksheet');
    const ws2 = sheet as Worksheet;
    expect(ws2.views[0]?.selections?.[0]?.activeCell).toBe('B3');
    expect(ws2.views[0]?.selections?.[0]?.sqref).toBe('A1:C5');
  });
});

describe('per-pane selections', () => {
  // As Excel writes a view frozen at B2 with D10 selected.
  const FROZEN = `<sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/><selection pane="bottomRight" activeCell="D10" sqref="D10"/></sheetView>`;

  it('keeps every pane selection, so Excel reopens on the same active cell', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setFreezePanes(ws, 'B2');
    const view = ws.views[0];
    if (!view) throw new Error('no view');
    view.selections = [
      { pane: 'topRight', activeCell: 'B1', sqref: 'B1' },
      { pane: 'bottomLeft', activeCell: 'A2', sqref: 'A2' },
      { pane: 'bottomRight', activeCell: 'D10', sqref: 'D10' },
    ];
    const bytes = await workbookToBytes(wb);
    const sheetXml = new TextDecoder().decode(unzipSync(bytes)['xl/worksheets/sheet1.xml']);
    expect(sheetXml).toContain(FROZEN.slice(FROZEN.indexOf('<selection'), FROZEN.indexOf('</sheetView>')));
    const back = expectSheet((await loadWorkbook(fromBuffer(bytes))).sheets[0]?.sheet).views[0];
    expect(back?.selections).toHaveLength(3);
    expect(back && activeSelection(back)?.activeCell).toBe('D10');
  });

  it('setActiveCell targets the active pane of a frozen view', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setFreezePanes(ws, 'B2');
    setActiveCell(ws, 'D10');
    expect(ws.views[0]?.selections).toEqual([{ pane: 'bottomRight', activeCell: 'D10', sqref: 'D10' }]);
  });
});

function expectSheet(ws: Worksheet | Chartsheet | undefined): Worksheet {
  if (!ws || !('rows' in ws)) throw new Error('expected worksheet');
  return ws;
}
