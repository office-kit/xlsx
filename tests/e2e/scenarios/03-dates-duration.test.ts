// Scenario 03: Date / Duration cells under both Excel epochs.
// Output: 03-dates-windows.xlsx (1900) + 03-dates-mac.xlsx (1904)
//
// What to verify in Excel:
// - 03-dates-windows.xlsx column A renders dates 2024-01-01 etc.
//   (after applying date format to the cells if they show as numbers).
// - 03-dates-mac.xlsx renders the same dates — opening in macOS Excel
//   under 1904 mode the date should match. (Cross-checking on Windows
//   Excel with date1904=true will show a different date — this is the
//   epoch flag's job.)
// - Duration cells in column B render fraction-of-day numbers; with
//   format `[h]:mm:ss` applied they should show e.g. 1:30:00 for 90 min.

import { describe, expect, it } from 'vitest';
import { makeDurationValue } from '../../../src/cell/index.js';
import { addCellXf, addNumFmt, defaultCellXf } from '../../../src/styles/index.js';
import { addWorksheet, createWorkbook } from '../../../src/workbook/index.js';
import { setCell } from '../../../src/worksheet/index.js';
import { writeWorkbook } from '../_helpers.js';

const buildDateBook = (date1904: boolean) => {
  const wb = createWorkbook({ date1904 });
  const ws = addWorksheet(wb, date1904 ? '1904 epoch' : '1900 epoch');

  // Reserve cellXfs[0] for default + allocate two date numFmt xfs.
  addCellXf(wb.styles, defaultCellXf());
  const dateNumFmtId = addNumFmt(wb.styles, 'yyyy-mm-dd');
  const dateXfId = addCellXf(wb.styles, {
    ...defaultCellXf(),
    numFmtId: dateNumFmtId,
    applyNumberFormat: true,
  });
  const durationNumFmtId = addNumFmt(wb.styles, '[h]:mm:ss');
  const durationXfId = addCellXf(wb.styles, {
    ...defaultCellXf(),
    numFmtId: durationNumFmtId,
    applyNumberFormat: true,
  });

  const labels = ['Date 1', 'Date 2', 'Date 3', 'Mid-day', 'Now-ish'];
  const dates = [
    new Date(Date.UTC(2024, 0, 1)),
    new Date(Date.UTC(2024, 5, 15)),
    new Date(Date.UTC(2024, 11, 31)),
    new Date(Date.UTC(2024, 0, 1, 12, 30)),
    new Date(Date.UTC(2026, 4, 5, 9, 0)),
  ];
  for (let i = 0; i < dates.length; i++) {
    setCell(ws, i + 1, 1, labels[i] ?? null);
    setCell(ws, i + 1, 2, dates[i] as Date, dateXfId);
  }

  // Duration column.
  setCell(ws, 1, 3, '90 minutes');
  setCell(ws, 1, 4, makeDurationValue(90 * 60 * 1000), durationXfId);
  setCell(ws, 2, 3, '2 hours 15 min');
  setCell(ws, 2, 4, makeDurationValue((2 * 60 + 15) * 60 * 1000), durationXfId);
  setCell(ws, 3, 3, '36 hours');
  setCell(ws, 3, 4, makeDurationValue(36 * 60 * 60 * 1000), durationXfId);

  return wb;
};

describe('e2e 03 — date / duration cells', () => {
  it('writes 03-dates-windows.xlsx (default 1900 epoch)', async () => {
    const wb = buildDateBook(false);
    const result = await writeWorkbook('03-dates-windows.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });

  it('writes 03-dates-mac.xlsx (1904 epoch)', async () => {
    const wb = buildDateBook(true);
    const result = await writeWorkbook('03-dates-mac.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
