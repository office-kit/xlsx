// Smoke tests for the docs/migrate-from-openpyxl.md examples. The
// migration guide is the first stop for users coming from openpyxl,
// so its examples must compile + run against the actual public API.
// A renamed export or shifted signature now fails CI.

import { describe, expect, it } from 'vitest';

describe('migrate-from-openpyxl — public API smoke', () => {
  it('Loading and saving — fromFile / toFile from @office-kit/xlsx/node, load/save from @office-kit/xlsx/io', async () => {
    const node = await import('../../src/node');
    const io = await import('../../src/io/index');
    expect(typeof node.fromFile).toBe('function');
    expect(typeof node.toFile).toBe('function');
    expect(typeof io.loadWorkbook).toBe('function');
    expect(typeof io.saveWorkbook).toBe('function');
  });

  it('Cells — setCell / setCellByCoord / iterRows / setFormula / makeErrorValue / makeRichText / makeTextRun / makeDurationValue', async () => {
    const worksheet = await import('../../src/worksheet/index');
    const cell = await import('../../src/cell/index');
    expect(typeof worksheet.setCell).toBe('function');
    expect(typeof worksheet.setCellByCoord).toBe('function');
    expect(typeof worksheet.iterRows).toBe('function');
    expect(typeof cell.setFormula).toBe('function');
    expect(typeof cell.makeErrorValue).toBe('function');
    expect(typeof cell.makeRichText).toBe('function');
    expect(typeof cell.makeTextRun).toBe('function');
    expect(typeof cell.makeDurationValue).toBe('function');
  });

  it('Styles — setCellFont / setCellFill / setCellNumberFormat with (wb, cell, …) signature', async () => {
    const styles = await import('../../src/styles/index');
    const workbook = await import('../../src/workbook/index');
    const worksheet = await import('../../src/worksheet/index');
    expect(typeof styles.setCellFont).toBe('function');
    expect(typeof styles.setCellFill).toBe('function');
    expect(typeof styles.setCellNumberFormat).toBe('function');

    // Exercise the canonical migration-guide flow.
    const wb = workbook.createWorkbook();
    const ws = workbook.addWorksheet(wb, 'Style');
    worksheet.setCell(ws, 1, 1, 'styled');
    const cell = ws.rows.get(1)?.get(1);
    if (!cell) throw new Error('cell missing');
    styles.setCellFont(wb, cell, styles.makeFont({ name: 'Arial', size: 14, bold: true }));
    styles.setCellFill(
      wb,
      cell,
      styles.makePatternFill({ patternType: 'solid', fgColor: styles.makeColor({ rgb: 'FFFFFF00' }) }),
    );
    styles.setCellNumberFormat(wb, cell, '#,##0.00');
    expect(cell.styleId).toBeGreaterThan(0);
  });

  it('Worksheets — addWorksheet / sheetNames / getActiveSheet / getSheet / removeSheet / mergeCells / setFreezePanes / getMergedCells', async () => {
    const workbook = await import('../../src/workbook/index');
    const worksheet = await import('../../src/worksheet/index');
    expect(typeof workbook.addWorksheet).toBe('function');
    expect(typeof workbook.sheetNames).toBe('function');
    expect(typeof workbook.getActiveSheet).toBe('function');
    expect(typeof workbook.getSheet).toBe('function');
    expect(typeof workbook.removeSheet).toBe('function');
    expect(typeof worksheet.mergeCells).toBe('function');
    expect(typeof worksheet.setFreezePanes).toBe('function');
    expect(typeof worksheet.getMergedCells).toBe('function');
  });

  it('Streaming write / read — createWriteOnlyWorkbook / loadWorkbookStream', async () => {
    const streaming = await import('../../src/streaming/index');
    expect(typeof streaming.createWriteOnlyWorkbook).toBe('function');
    expect(typeof streaming.loadWorkbookStream).toBe('function');
  });
});
