import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook, parseSheetEntries } from '../../src/io/load.js';
import { resolveRelTarget } from '../../src/packaging/part-name.js';
import { parseXml } from '../../src/xml/parser.js';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

describe('resolveRelTarget', () => {
  it('treats /-prefixed targets as package-absolute', () => {
    expect(resolveRelTarget('xl/workbook.xml', '/xl/styles.xml')).toBe('xl/styles.xml');
  });

  it('joins relative targets against the source part dir', () => {
    expect(resolveRelTarget('xl/workbook.xml', 'worksheets/sheet1.xml')).toBe('xl/worksheets/sheet1.xml');
  });

  it('handles root-rels source where parent dir is empty', () => {
    expect(resolveRelTarget('', 'xl/workbook.xml')).toBe('xl/workbook.xml');
  });

  it('collapses .. segments', () => {
    expect(resolveRelTarget('xl/worksheets/sheet1.xml', '../theme/theme1.xml')).toBe('xl/theme/theme1.xml');
  });
});

describe('parseSheetEntries', () => {
  it('extracts name / sheetId / r:id from a multi-sheet workbook.xml', () => {
    const workbookXml = readFileSync(resolve(FIXTURES, 'empty.xlsx'));
    // Use the loadWorkbook flow indirectly: parse the workbook.xml inside.
    // Easier — read the raw entry via openZip elsewhere; here we just unit-test
    // parseSheetEntries against a hand-built XML string.
    const xml = `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Alpha" sheetId="1" r:id="rId1"/>
    <sheet name="Beta" sheetId="3" r:id="rId2" state="hidden"/>
    <sheet name="Gamma" sheetId="7" r:id="rId3" state="veryHidden"/>
  </sheets>
</workbook>`;
    const entries = parseSheetEntries(parseXml(xml));
    expect(entries).toEqual([
      { name: 'Alpha', sheetId: 1, rId: 'rId1', state: 'visible' },
      { name: 'Beta', sheetId: 3, rId: 'rId2', state: 'hidden' },
      { name: 'Gamma', sheetId: 7, rId: 'rId3', state: 'veryHidden' },
    ]);
    // workbookXml is loaded only to verify the fixture itself is reachable.
    expect(workbookXml.byteLength).toBeGreaterThan(0);
  });

  it('throws on a sheet missing required attrs', () => {
    expect(() =>
      parseSheetEntries(
        parseXml(
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="A" sheetId="1"/></sheets></workbook>',
        ),
      ),
    ).toThrowError(/r:id/);
  });
});

describe('loadWorkbook — empty.xlsx skeleton', () => {
  it('reads openpyxl genuine/empty.xlsx and produces a 3-sheet Workbook scaffold', async () => {
    const bytes = readFileSync(resolve(FIXTURES, 'empty.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));

    expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(wb.sheets.map((s) => s.sheetId)).toEqual([1, 2, 3]);
    expect(wb.sheets.every((s) => s.state === 'visible')).toBe(true);
    // Cells haven't been read in the minimum-skeleton stage.
    for (const ref of wb.sheets) {
      if (ref.kind !== 'worksheet') throw new Error('expected only worksheets in fixture');
      expect(ref.sheet.rows.size).toBe(0);
    }
  });

  it('reads sheet content cells through the worksheet reader', async () => {
    // Build a synthetic single-sheet xlsx with one number cell.
    const { createZipWriter } = await import('../../src/zip/writer.js');
    const { toBuffer } = await import('../../src/io/node.js');
    const sink = toBuffer();
    const w = createZipWriter(sink);
    const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);
    await w.addEntry(
      '[Content_Types].xml',
      utf8(
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
      ),
    );
    await w.addEntry(
      '_rels/.rels',
      utf8(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      ),
    );
    await w.addEntry(
      'xl/workbook.xml',
      utf8(
        '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ),
    );
    await w.addEntry(
      'xl/_rels/workbook.xml.rels',
      utf8(
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
      ),
    );
    await w.addEntry(
      'xl/worksheets/sheet1.xml',
      utf8(
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>42</v></c><c r="B1" t="b"><v>1</v></c></row><row r="3"><c r="C3"><f>1+2</f><v>3</v></c></row></sheetData></worksheet>',
      ),
    );
    await w.finalize();
    const wb = await loadWorkbook(fromBuffer(sink.result()));
    const ws = wb.sheets[0]?.sheet;
    if (!ws || !('rows' in ws)) throw new Error('expected one worksheet');
    expect(ws.title).toBe('Data');
    const { getCell } = await import('../../src/worksheet/worksheet.js');
    expect(getCell(ws, 1, 1)?.value).toBe(42);
    expect(getCell(ws, 1, 2)?.value).toBe(true);
    const f = getCell(ws, 3, 3)?.value as { kind: string; formula: string; cachedValue: number };
    expect(f.kind).toBe('formula');
    expect(f.formula).toBe('1+2');
    expect(f.cachedValue).toBe(3);
  });

  it('resolves t="s" cells against the sharedStrings table', async () => {
    const bytes = readFileSync(resolve(FIXTURES, 'empty-with-styles.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));
    const ws = wb.sheets[0]?.sheet;
    if (!ws || !('rows' in ws)) throw new Error('expected one worksheet');
    // empty-with-styles.xlsx has A1 as t="s" -> sst[0] = "TEST HERE"
    const { getCell } = await import('../../src/worksheet/worksheet.js');
    expect(getCell(ws, 1, 1)?.value).toBe('TEST HERE');
    // A2..A5 are numeric (date / pi / fraction / scientific) — read as numbers.
    expect(typeof getCell(ws, 2, 1)?.value).toBe('number');
    expect(getCell(ws, 3, 1)?.value).toBeCloseTo(3.14);
  });

  it('loads xl/styles.xml into Workbook.styles', async () => {
    const bytes = readFileSync(resolve(FIXTURES, 'empty-with-styles.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));
    expect(wb.styles.fonts.length).toBe(1);
    expect(wb.styles.fills.length).toBe(2);
    expect(wb.styles.cellXfs.length).toBe(5);
    // The cell A2 has s="2" which points at numFmtId=14 (date format).
    const ws = wb.sheets[0]?.sheet;
    if (!ws || !('rows' in ws)) throw new Error('expected one worksheet');
    const { getCell } = await import('../../src/worksheet/worksheet.js');
    const a2 = getCell(ws, 2, 1);
    expect(a2?.styleId).toBe(2);
    expect(wb.styles.cellXfs[a2?.styleId ?? 0]?.numFmtId).toBe(14);
  });

  it('rejects an archive missing [Content_Types].xml', async () => {
    const { createZipWriter } = await import('../../src/zip/writer.js');
    const { toBuffer } = await import('../../src/io/node.js');
    const sink = toBuffer();
    const w = createZipWriter(sink);
    await w.addEntry('foo.txt', new TextEncoder().encode('bar'));
    await w.finalize();
    await expect(loadWorkbook(fromBuffer(sink.result()))).rejects.toThrow(/Content_Types/);
  });
});
