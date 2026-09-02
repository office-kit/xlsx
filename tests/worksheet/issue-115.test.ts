// Regression for https://github.com/office-kit/xlsx/issues/115 — a formula
// cell whose cached result is the empty string (`<c t="str"><f>…</f><v/></c>`)
// came back with both the `t="str"` and the `<v/>` gone. The reader collapsed
// an empty `<v/>` and an absent `<v>` into the same "no cached value" state.
//
// It matters for links into another workbook: Excel cannot recalculate those
// without opening the other file, so the cached value is the only thing it has
// to display.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell';
import { makeSharedStrings } from '../../src/workbook/shared-strings';
import { parseWorksheetXml } from '../../src/worksheet/reader';
import { worksheetToBytes } from '../../src/worksheet/writer';
import { getCell, makeWorksheet, setCell } from '../../src/worksheet/worksheet';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const sheetXml = (cells: string): string =>
  `<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData><row r="1">${cells}</row></sheetData></worksheet>`;

describe('issue #115 — an empty cached formula value survives a round-trip', () => {
  it('keeps t="str" and <v/> for a formula whose cached result is empty', () => {
    const ws = parseWorksheetXml(sheetXml('<c r="A1" t="str"><f>[1]Extern!$A$1</f><v/></c>'), 'Tabelle1', {
      sharedStrings: [],
    });
    const value = getCell(ws, 1, 1)?.value;
    expect(value).toMatchObject({ kind: 'formula', formula: '[1]Extern!$A$1', cachedValue: '' });

    const out = new TextDecoder().decode(worksheetToBytes(ws, { sharedStrings: makeSharedStrings() }));
    expect(out).toContain('<c r="A1" t="str"><f>[1]Extern!$A$1</f><v/></c>');
  });

  it('still reports no cached value when <v> is absent altogether', () => {
    const ws = parseWorksheetXml(sheetXml('<c r="A1"><f>SUM(B1:B2)</f></c>'), 'Tabelle1', {
      sharedStrings: [],
    });
    const value = getCell(ws, 1, 1)?.value as { cachedValue?: unknown };
    expect(value.cachedValue).toBeUndefined();

    const out = new TextDecoder().decode(worksheetToBytes(ws, { sharedStrings: makeSharedStrings() }));
    expect(out).toContain('<c r="A1"><f>SUM(B1:B2)</f></c>');
  });

  it('treats an empty <v/> under the numeric default as no cached number', () => {
    const ws = parseWorksheetXml(sheetXml('<c r="A1"><f>SUM(B1:B2)</f><v/></c>'), 'Tabelle1', {
      sharedStrings: [],
    });
    const value = getCell(ws, 1, 1)?.value as { cachedValue?: unknown };
    expect(value.cachedValue).toBeUndefined();
  });

  it('round-trips a non-empty cached string unchanged', () => {
    const ws = makeWorksheet('Tabelle1');
    setFormula(setCell(ws, 1, 1), 'A2&""', { cachedValue: 'Text' });
    const out = new TextDecoder().decode(worksheetToBytes(ws, { sharedStrings: makeSharedStrings() }));
    expect(out).toContain('<c r="A1" t="str"><f>A2&amp;""</f><v>Text</v></c>');
  });
});
