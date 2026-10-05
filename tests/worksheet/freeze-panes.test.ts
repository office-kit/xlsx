import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { freezePaneRef, makeFreezePane, makeSheetView } from '../../src/worksheet/views.js';
import { getFreezePanes, setCell, setFreezePanes, type Worksheet } from '../../src/worksheet/worksheet.js';

const expectSheet = (ws: Worksheet | import('../../src/chartsheet/chartsheet.js').Chartsheet | undefined): Worksheet => {
  if (!ws) throw new Error('expected sheet');
  if (!('rows' in ws)) throw new Error('expected worksheet, got chartsheet');
  return ws;
};

describe('makeFreezePane', () => {
  it('B2 → 1 row + 1 col, activePane=bottomRight', () => {
    const p = makeFreezePane('B2');
    expect(p.xSplit).toBe(1);
    expect(p.ySplit).toBe(1);
    expect(p.activePane).toBe('bottomRight');
    expect(p.state).toBe('frozen');
    expect(p.topLeftCell).toBe('B2');
  });

  it('A2 → row only (bottomLeft)', () => {
    const p = makeFreezePane('A2');
    expect(p.xSplit).toBeUndefined();
    expect(p.ySplit).toBe(1);
    expect(p.activePane).toBe('bottomLeft');
  });

  it('B1 → col only (topRight)', () => {
    const p = makeFreezePane('B1');
    expect(p.xSplit).toBe(1);
    expect(p.ySplit).toBeUndefined();
    expect(p.activePane).toBe('topRight');
  });

  it('A1 throws (no axis to freeze)', () => {
    expect(() => makeFreezePane('A1')).toThrowError(/A1/);
  });
});

describe('setFreezePanes / getFreezePanes', () => {
  it('lazily creates the primary view', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    expect(ws.views.length).toBe(0);
    setFreezePanes(ws, 'B2');
    expect(ws.views.length).toBe(1);
    expect(ws.views[0]?.pane?.state).toBe('frozen');
    expect(getFreezePanes(ws)).toBe('B2');
  });

  it('passing undefined clears the freeze', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setFreezePanes(ws, 'C3');
    setFreezePanes(ws, undefined);
    expect(ws.views[0]?.pane).toBeUndefined();
    expect(getFreezePanes(ws)).toBeUndefined();
  });
});

describe('freezePaneRef from a hand-built SheetView', () => {
  it('returns the topLeftCell when present', () => {
    const view = makeSheetView({ pane: { state: 'frozen', xSplit: 2, ySplit: 1, topLeftCell: 'C2' } });
    expect(freezePaneRef(view)).toBe('C2');
  });

  it('falls back to (xSplit+1, ySplit+1) when topLeftCell is missing', () => {
    const view = makeSheetView({ pane: { state: 'frozen', xSplit: 2, ySplit: 1 } });
    expect(freezePaneRef(view)).toBe('C2');
  });

  it('returns undefined for non-frozen panes', () => {
    const view = makeSheetView({ pane: { state: 'split', xSplit: 2, ySplit: 1 } });
    expect(freezePaneRef(view)).toBeUndefined();
  });
});

describe('SheetView round-trip through saveWorkbook → loadWorkbook', () => {
  it('preserves freeze panes', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'F');
    setCell(ws, 1, 1, 'header');
    setFreezePanes(ws, 'B2');

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(getFreezePanes(ws2)).toBe('B2');
    expect(ws2.views[0]?.pane?.activePane).toBe('bottomRight');
  });

  it('preserves selection blocks', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    ws.views.push(makeSheetView({ tabSelected: true, selections: [{ activeCell: 'C5', sqref: 'C5' }] }));
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ws2 = expectSheet(wb2.sheets[0]?.sheet);
    expect(ws2.views[0]?.tabSelected).toBe(true);
    expect(ws2.views[0]?.selections?.[0]?.activeCell).toBe('C5');
    expect(ws2.views[0]?.selections?.[0]?.sqref).toBe('C5');
  });

  it('reads sheetView from openpyxl genuine fixture', async () => {
    const { readFileSync } = await import('node:fs');
    const { dirname, resolve } = await import('node:path');
    const { fileURLToPath } = await import('node:url');
    const here = dirname(fileURLToPath(import.meta.url));
    const fixturePath = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine/empty-with-styles.xlsx');
    const bytes = readFileSync(fixturePath);
    const wb = await loadWorkbook(fromBuffer(bytes));
    const ws = expectSheet(wb.sheets[0]?.sheet);
    // empty-with-styles.xlsx has sheetView with tabSelected="1"
    expect(ws.views.length).toBeGreaterThan(0);
    expect(ws.views[0]?.tabSelected).toBe(true);
    expect(ws.views[0]?.selections?.[0]?.activeCell).toBe('A3');
  });
});
