// Regression: future-function (dynamic-array) formulas must keep their
// `_xlfn.` storage prefix across a load → save round-trip.
//
// In OOXML, modern functions (SCAN, LAMBDA, XLOOKUP, FILTER, SEQUENCE, …)
// are stored in `<f>` with an `_xlfn.` prefix — Excel strips it for display
// and re-adds it on save. A stored bare `SCAN(...)` is an unknown name, so
// Excel renders #NAME?. The reader used to strip `_xlfn.` into the model and
// the writer emits formula text verbatim, so a loaded `_xlfn.SCAN` was
// written back bare — corrupting every dynamic-array formula it merely passed
// through. Consistent with openpyxl, which surfaces the prefix verbatim, the
// reader now keeps the formula text as-is.

import { describe, expect, it } from 'vitest';
import { unzipSync, strFromU8 } from 'fflate';
import { setArrayFormula } from '../../src/cell/cell';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { setCell } from '../../src/worksheet/worksheet';

const SCAN_FORMULA = '_xlfn.SCAN(0,B2:F2,_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b))';

const sheet1Xml = (bytes: Uint8Array): string => {
  const part = unzipSync(bytes)['xl/worksheets/sheet1.xml'];
  if (!part) throw new Error('xl/worksheets/sheet1.xml missing from output');
  return strFromU8(part);
};

describe('future-function `_xlfn.` prefix round-trip', () => {
  it('keeps `_xlfn.SCAN` / `_xlfn.LAMBDA` intact through load → save', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    // A dynamic-array spill anchor, exactly as Excel stores it on disk.
    const anchor = setCell(ws, 2, 8, null); // H2
    setArrayFormula(anchor, 'H2:L2', SCAN_FORMULA, { cachedValue: 10 });

    // First serialization ships the prefix verbatim (writer passes text through).
    const bytes1 = await workbookToBytes(wb);
    expect(sheet1Xml(bytes1)).toContain('_xlfn.SCAN');

    // Load it back — this is where the read path used to strip the prefix.
    const wb2 = await loadWorkbook(fromBuffer(bytes1));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    const cell = ref0.sheet.rows.get(2)?.get(8);
    const value = cell?.value;
    if (!value || typeof value !== 'object' || !('kind' in value) || value.kind !== 'formula') {
      throw new Error('H2 is not a formula');
    }
    expect(value.formula).toContain('_xlfn.SCAN');
    expect(value.formula).toContain('_xlfn.LAMBDA');

    // Save again — the prefix must survive so Excel does not render #NAME?.
    const bytes2 = await workbookToBytes(wb2);
    const xml2 = sheet1Xml(bytes2);
    expect(xml2).toContain('_xlfn.SCAN');
    expect(xml2).toContain('_xlfn.LAMBDA');
    // A bare `SCAN(` (no prefix) would be the corruption we are guarding against.
    expect(xml2).not.toMatch(/(?<!\.)\bSCAN\(/);
  });

  // The `_xlfn._xlws.` compound prefix is the worksheet-function form Excel
  // writes for the spilling dynamic-array functions (FILTER, SORT, UNIQUE,
  // …). Stripping only the leading `_xlfn.` would leave a bare `_xlws.FILTER`,
  // and stripping both would leave bare `FILTER` — both #NAME? on reopen. The
  // whole prefix must survive verbatim.
  it('keeps the compound `_xlfn._xlws.` prefix (FILTER/SORT) intact', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const formula = '_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A2:A8,B2:B8>0))';
    const anchor = setCell(ws, 2, 8, null); // H2
    setArrayFormula(anchor, 'H2:H8', formula, { cachedValue: 1 });

    const bytes1 = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes1));
    const ref0 = wb2.sheets[0];
    if (ref0?.kind !== 'worksheet') throw new Error('expected worksheet');
    const value = ref0.sheet.rows.get(2)?.get(8)?.value;
    if (!value || typeof value !== 'object' || !('kind' in value) || value.kind !== 'formula') {
      throw new Error('H2 is not a formula');
    }
    // Model preserves the full compound prefix, not a partially-stripped form.
    expect(value.formula).toBe(formula);

    const xml2 = sheet1Xml(await workbookToBytes(wb2));
    expect(xml2).toContain('_xlfn._xlws.FILTER');
    expect(xml2).toContain('_xlfn._xlws.SORT');
    // No partially- or fully-stripped name may leak through.
    expect(xml2).not.toMatch(/(?<!\.)\b(?:FILTER|SORT)\(/);
  });
});
