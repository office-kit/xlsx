// Conformance smoke test. Locks in the contract that anything @office-kit/xlsx
// emits is OPC-clean, schema-valid, and semantically consistent.
//
// Each new feature added to the writer should grow this file (or a sibling)
// rather than adding bespoke assertions, so the validator stays the single
// gate for "the bytes we wrote are a real xlsx".

import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell';
import { workbookToBytes } from '../../src/io/save';
import { setCellFont } from '../../src/styles/cell-style';
import { makeFont } from '../../src/styles/fonts';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { mergeCells, setCell, type Worksheet } from '../../src/worksheet/worksheet';
import { validateXlsx } from './validate';

const dump = (issues: { tier: string; part: string; message: string }[]): string =>
  issues.map((i) => `[${i.tier}] ${i.part}: ${i.message}`).join('\n');

describe('conformance: @office-kit/xlsx output validates against ECMA-376', () => {
  it('empty single-sheet workbook', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Only');
    const result = await validateXlsx(await workbookToBytes(wb));
    expect(result.issues, dump(result.issues)).toEqual([]);
    expect(result.ok).toBe(true);
  });

  it('mixed cell values, formulas, multiple sheets', async () => {
    const wb = createWorkbook();
    const a = addWorksheet(wb, 'Alpha');
    setCell(a, 1, 1, 42);
    setCell(a, 1, 2, 'hello');
    setCell(a, 1, 3, true);
    setCell(a, 2, 1, 'with " < > & symbols');
    const f = setCell(a, 2, 2);
    setFormula(f, 'A1+1', { cachedValue: 43 });
    addWorksheet(wb, 'Beta');
    const c = addWorksheet(wb, 'Gamma');
    setCell(c, 5, 5, 100);
    const result = await validateXlsx(await workbookToBytes(wb));
    expect(result.issues, dump(result.issues)).toEqual([]);
  });

  it('styled cell + custom font', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Styled');
    const c = setCell(ws, 1, 1, 'bold');
    setCellFont(wb, c, makeFont({ bold: true, size: 14 }));
    const result = await validateXlsx(await workbookToBytes(wb));
    expect(result.issues, dump(result.issues)).toEqual([]);
  });

  it('merged cells', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Merged') as Worksheet;
    setCell(ws, 1, 1, 'header');
    mergeCells(ws, 'A1:C1');
    const result = await validateXlsx(await workbookToBytes(wb));
    expect(result.issues, dump(result.issues)).toEqual([]);
  });
});

// Negative tests: the validator must actually catch these. If any of these
// stop reporting, that's a regression in the validator itself.
describe('conformance: validator detects deliberately broken xlsx', () => {
  const mutate = async (
    transform: (entries: Record<string, Uint8Array>) => Record<string, Uint8Array>,
  ): Promise<Uint8Array> => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, 'a');
    const original = await workbookToBytes(wb);
    return zipSync(transform(unzipSync(original)));
  };

  it('rejects a worksheet with a coordinate inside the wrong row', async () => {
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const bytes = await mutate((entries) => {
      const sheet = decoder.decode(entries['xl/worksheets/sheet1.xml']);
      // change <c r="A1" …> to <c r="A99" …> while leaving its parent <row r="1">
      const broken = sheet.replace(/r="A1"/, 'r="A99"');
      entries['xl/worksheets/sheet1.xml'] = encoder.encode(broken);
      return entries;
    });
    const result = await validateXlsx(bytes);
    expect(result.ok).toBe(false);
    expect(result.issues.some((i) => i.tier === 'semantic' && /A99/.test(i.message))).toBe(true);
  });

  it('rejects a package missing a part referenced by an Override', async () => {
    const bytes = await mutate((entries) => {
      delete entries['xl/styles.xml'];
      return entries;
    });
    const result = await validateXlsx(bytes);
    expect(result.ok).toBe(false);
    expect(
      result.issues.some(
        (i) => i.tier === 'opc' && /styles\.xml/.test(i.part) && /does not exist/.test(i.message),
      ),
    ).toBe(true);
  });

  it('rejects schema-invalid worksheet XML', async () => {
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const bytes = await mutate((entries) => {
      const sheet = decoder.decode(entries['xl/worksheets/sheet1.xml']);
      // Insert an element that doesn't exist in the SpreadsheetML schema.
      const broken = sheet.replace('<sheetData>', '<sheetData><nonsenseElement/>');
      entries['xl/worksheets/sheet1.xml'] = encoder.encode(broken);
      return entries;
    });
    const result = await validateXlsx(bytes);
    expect(result.ok).toBe(false);
    expect(result.issues.some((i) => i.tier === 'xsd')).toBe(true);
  });

  it('rejects two <sheet> elements that differ only in case', async () => {
    // Hand-craft the broken workbook bytes instead of going through
    // addWorksheet, since the typed API now rejects this case-insensitively at
    // the source. The validator must catch the same break in already-written
    // bytes (e.g. files produced before that fix landed, or by other tools).
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const wb = createWorkbook();
    addWorksheet(wb, 'Data');
    addWorksheet(wb, 'Other');
    const original = await workbookToBytes(wb);
    const entries = unzipSync(original);
    const workbookXml = decoder.decode(entries['xl/workbook.xml']);
    const broken = workbookXml.replace(/name="Other"/, 'name="data"');
    entries['xl/workbook.xml'] = encoder.encode(broken);
    const bytes = zipSync(entries);
    const result = await validateXlsx(bytes, { skipXsd: true });
    expect(result.ok).toBe(false);
    expect(
      result.issues.some(
        (i) => i.tier === 'semantic' && /case-insensitive collision/.test(i.message),
      ),
    ).toBe(true);
  });
});
