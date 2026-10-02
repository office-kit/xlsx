// The `list*` accessors and `getMergedCells` return `ReadonlyArray`, which stops
// a TypeScript caller writing to them but says nothing about whether the array
// is the sheet's own. It is, and their docblocks used to call it a "snapshot",
// which reads as the opposite. These pin what the corrected docs now promise,
// so the wording and the behaviour cannot drift apart again.

import { describe, expect, it } from 'vitest';
import { addDefinedName, listDefinedNames, removeDefinedNames } from '../../src/workbook/defined-names.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { makeCellRange, shiftRange } from '../../src/worksheet/cell-range.js';
import {
  getMergedCells,
  listHyperlinks,
  mergeCells,
  removeAllHyperlinks,
  setHyperlink,
} from '../../src/worksheet/worksheet.js';

describe('the list accessors are live views', () => {
  it('getMergedCells reflects a merge added after the call', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const view = getMergedCells(ws);
    expect(view).toHaveLength(0);
    mergeCells(ws, 'A1:B2');
    expect(view).toHaveLength(1);
  });

  it('a removeAll* leaves an earlier return value stale', () => {
    // The remove helpers assign a fresh array to the worksheet field rather
    // than emptying it, so a view taken before one keeps the old contents.
    // Worth pinning because it is the surprising half of the contract.
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setHyperlink(ws, 'A1', { target: 'https://example.com' });
    const view = listHyperlinks(ws);
    expect(view).toHaveLength(1);
    removeAllHyperlinks(ws);
    expect(view).toHaveLength(1);
    expect(listHyperlinks(ws)).toHaveLength(0);
  });

  it('copying the view detaches it', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const copy = [...getMergedCells(ws)];
    mergeCells(ws, 'A1:B2');
    expect(copy).toHaveLength(0);
  });
});

describe('listDefinedNames scope option', () => {
  it('narrows with `workbook`, and lists everything when omitted', () => {
    // The docblock used to offer `{ scope: undefined }` for workbook-scope. It
    // would have fallen through to the `all` default and returned every name,
    // and under `exactOptionalPropertyTypes` it does not even compile, so the
    // only spelling that narrows is `'workbook'`.
    const wb = createWorkbook();
    addWorksheet(wb, 'S');
    addDefinedName(wb, { name: 'GlobalOne', value: 'S!$A$1' });
    addDefinedName(wb, { name: 'SheetOne', value: 'S!$A$2', scope: 0 });

    expect(listDefinedNames(wb)).toHaveLength(2);
    expect(listDefinedNames(wb, {})).toHaveLength(2);
    expect(listDefinedNames(wb, { scope: 'all' })).toHaveLength(2);
    expect(listDefinedNames(wb, { scope: 'workbook' }).map((d) => d.name)).toEqual(['GlobalOne']);
    expect(listDefinedNames(wb, { scope: 0 }).map((d) => d.name)).toEqual(['SheetOne']);
  });

  it('is a live view when listing all and a fresh array when narrowed', () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'S');
    addDefinedName(wb, { name: 'GlobalOne', value: 'S!$A$1' });
    const all = listDefinedNames(wb);
    const narrowed = listDefinedNames(wb, { scope: 'workbook' });
    addDefinedName(wb, { name: 'GlobalTwo', value: 'S!$A$2' });
    expect(all).toHaveLength(2);
    expect(narrowed).toHaveLength(1);
    removeDefinedNames(wb, () => true);
    expect(listDefinedNames(wb)).toHaveLength(0);
  });
});

describe('shiftRange bounds', () => {
  it('throws rather than clamping when the shift leaves the grid', () => {
    // The docblock claimed the result was "clamped to the OOXML grid". Clamping
    // would silently change which cells the range covers, which is why the
    // implementation throws and `shiftRangeStr` documents that it does.
    const r = makeCellRange(1, 1, 2, 2);
    expect(() => shiftRange(r, -5, 0)).toThrowError(OpenXmlSchemaError);
    expect(() => shiftRange(r, 0, -5)).toThrowError(OpenXmlSchemaError);
    expect(() => shiftRange(makeCellRange(1, 1, 1, 1), 1_048_576, 0)).toThrowError(OpenXmlSchemaError);
    expect(() => shiftRange(makeCellRange(1, 1, 1, 1), 0, 16_384)).toThrowError(OpenXmlSchemaError);
  });

  it('allows a shift that lands on the grid edge', () => {
    expect(shiftRange(makeCellRange(1, 1, 1, 1), 1_048_575, 16_383)).toEqual({
      minRow: 1_048_576,
      minCol: 16_384,
      maxRow: 1_048_576,
      maxCol: 16_384,
    });
  });
});
