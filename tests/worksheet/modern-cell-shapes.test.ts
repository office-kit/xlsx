// Cell shapes a current Excel writes that the reader used to refuse: the error
// tokens added since 2018, and the `t="d"` ISO-8601 date of ISO 29500 strict.
// One such cell anywhere in a part aborted the whole load, so a sheet of good
// rows was lost to a single spilled formula.

import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { isErrorValue } from '../../src/cell/cell.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { ERROR_CODES } from '../../src/utils/inference.js';
import { addWorksheet, createWorkbook, getSheet, type Workbook } from '../../src/workbook/workbook.js';
import { parseWorksheetXml } from '../../src/worksheet/reader.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const SHEET_PART = 'xl/worksheets/sheet1.xml';

const sheet = (body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<worksheet xmlns="${MAIN_NS}"><sheetData>${body}</sheetData></worksheet>`;

const readSheet = (body: string) => parseWorksheetXml(sheet(body), 'Data', { sharedStrings: [] });

/**
 * A workbook saved by the library with the `<c>` of its one cell swapped out.
 * Going through the real save path keeps content types, rels and the workbook
 * part exactly as a legitimate file has them, so the load reaches the cell.
 */
const savedWithCell = async (cellXml: string): Promise<Uint8Array> => {
  const wb = createWorkbook();
  setCell(addWorksheet(wb, 'Data'), 1, 1, 42);
  const entries = unzipSync(await workbookToBytes(wb));
  const part = entries[SHEET_PART];
  if (part === undefined) throw new Error(`save produced no ${SHEET_PART}`);
  const patched = new TextDecoder().decode(part).replace('<c r="A1"><v>42</v></c>', cellXml);
  entries[SHEET_PART] = new TextEncoder().encode(patched);
  return zipSync(entries);
};

/** Load a patched single-cell workbook and hand back its one worksheet. */
const loadPatched = async (cellXml: string) => {
  const wb = await loadWorkbook(fromBuffer(await savedWithCell(cellXml)));
  const ws = getSheet(wb, 'Data');
  if (ws === undefined) throw new Error('load produced no worksheet');
  return { wb, ws };
};

const savedSheetText = async (wb: Workbook): Promise<string> => {
  const part = unzipSync(await workbookToBytes(wb))[SHEET_PART];
  if (part === undefined) throw new Error(`save produced no ${SHEET_PART}`);
  return new TextDecoder().decode(part);
};

const MODERN_CODES = [
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#CONNECT!',
  '#BUSY!',
  '#UNKNOWN!',
  '#PYTHON!',
  '#EXTERNAL!',
] as const;

describe('error tokens Excel added after 2018', () => {
  it('reads each one as an error cell', () => {
    for (const code of MODERN_CODES) {
      const ws = readSheet(`<row r="1"><c r="A1" t="e"><v>${code}</v></c></row>`);
      expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'error', code });
    }
  });

  it('carries #SPILL! through loadWorkbook', async () => {
    const { ws } = await loadPatched('<c r="A1" t="e"><v>#SPILL!</v></c>');
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'error', code: '#SPILL!' });
  });

  it('lists them in ERROR_CODES, so a caller can name one on a write too', () => {
    for (const code of MODERN_CODES) expect(ERROR_CODES.has(code)).toBe(true);
  });

  it('writes them back as t="e" verbatim', async () => {
    const { wb } = await loadPatched('<c r="A1" t="e"><v>#CALC!</v></c>');
    expect(await savedSheetText(wb)).toContain('t="e"><v>#CALC!</v>');
  });
});

describe('an error token the library does not model', () => {
  it('keeps the token rather than dropping the cell', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="e"><v>#NOTYET!</v></c></row>');
    const value = getCell(ws, 1, 1)?.value ?? null;
    expect(isErrorValue(value)).toBe(true);
    expect(value).toEqual({ kind: 'error', code: '#NOTYET!' });
  });

  it('survives a load then save', async () => {
    const { wb } = await loadPatched('<c r="A1" t="e"><v>#NOTYET!</v></c>');
    expect(await savedSheetText(wb)).toContain('t="e"><v>#NOTYET!</v>');
  });

  it('rejects a t="e" payload that is not an error token at all', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="e"><v>oops</v></c></row>')).toThrow(OpenXmlSchemaError);
    expect(() => readSheet('<row r="1"><c r="A1" t="e"><v>oops</v></c></row>')).toThrow(
      'worksheet: <v>oops</v> at Data!A1 is not an Excel error token',
    );
  });

  it('rejects a t="e" with no value', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="e"/></row>')).toThrow(OpenXmlSchemaError);
  });
});

describe('t="d", the ISO 29500 strict date cell', () => {
  it('reads a naive datetime as UTC, the way every other Date in the model is read', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="d"><v>2024-03-14T00:00:00</v></c></row>');
    expect(getCell(ws, 1, 1)?.value).toEqual(new Date(Date.UTC(2024, 2, 14)));
  });

  it('reads the date-only, fractional-second and offset forms', () => {
    const cases: ReadonlyArray<readonly [string, Date]> = [
      ['2024-03-14', new Date(Date.UTC(2024, 2, 14))],
      ['2024-03-14T12:30:45.500', new Date(Date.UTC(2024, 2, 14, 12, 30, 45, 500))],
      ['2024-03-14T12:30:45Z', new Date(Date.UTC(2024, 2, 14, 12, 30, 45))],
      ['2024-03-14T12:30:45+02:00', new Date(Date.UTC(2024, 2, 14, 10, 30, 45))],
      ['2024-03-14T12:30:45-05:30', new Date(Date.UTC(2024, 2, 14, 18, 0, 45))],
    ];
    for (const [text, expected] of cases) {
      const ws = readSheet(`<row r="1"><c r="A1" t="d"><v>${text}</v></c></row>`);
      expect(getCell(ws, 1, 1)?.value).toEqual(expected);
    }
  });

  it('reads the time-only and duration forms as a duration', () => {
    const ws = readSheet(
      '<row r="1"><c r="A1" t="d"><v>14:30:00</v></c><c r="B1" t="d"><v>PT4H30M</v></c></row>',
    );
    expect(getCell(ws, 1, 1)?.value).toEqual({ kind: 'duration', ms: 52_200_000 });
    expect(getCell(ws, 1, 2)?.value).toEqual({ kind: 'duration', ms: 16_200_000 });
  });

  it('reads a blank value as an empty cell, as t="n" and t="b" do', () => {
    const ws = readSheet('<row r="1"><c r="A1" t="d"><v></v></c><c r="B1" t="d"/></row>');
    expect(getCell(ws, 1, 1)?.value).toBeNull();
    expect(getCell(ws, 1, 2)?.value).toBeNull();
  });

  it('rejects a date that does not exist', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><v>2024-02-30</v></c></row>')).toThrow(
      'worksheet: <v>2024-02-30</v> at Data!A1 is not an ISO 8601 date, time or duration',
    );
  });

  it('rejects text', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="d"><v>last tuesday</v></c></row>')).toThrow(
      OpenXmlSchemaError,
    );
  });

  it('loads through loadWorkbook and saves back as a serial, not as t="d"', async () => {
    const { wb, ws } = await loadPatched('<c r="A1" t="d"><v>2024-03-14T00:00:00</v></c>');
    expect(getCell(ws, 1, 1)?.value).toEqual(new Date(Date.UTC(2024, 2, 14)));
    const out = await savedSheetText(wb);
    expect(out).toContain('<c r="A1"><v>45365</v></c>');
    expect(out).not.toContain('t="d"');
  });
});

describe('a cell type outside ST_CellType', () => {
  it('is still refused, naming the type', () => {
    expect(() => readSheet('<row r="1"><c r="A1" t="q"><v>1</v></c></row>')).toThrow(
      'worksheet: unknown cell type t="q"',
    );
  });
});
