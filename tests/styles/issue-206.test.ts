import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { assert, describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { fromTree, toTree } from '../../src/schema/serialize.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { getCellFont, setBold, setCellFont } from '../../src/styles/cell-style.js';
import { addDxf, getDxfs } from '../../src/styles/differential.js';
import { fontToCss, makeFont } from '../../src/styles/fonts.js';
import { FontSchema } from '../../src/styles/fonts.schema.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';
import { SHEET_MAIN_NS } from '../../src/xml/namespaces.js';
import { parseXml } from '../../src/xml/parser.js';
import { serializeXml } from '../../src/xml/serializer.js';
import { validateXlsx } from '../conformance/validate.js';

const parseFont = (underline: string) =>
  fromTree(parseXml(`<font xmlns="${SHEET_MAIN_NS}"><name val="Arial"/>${underline}</font>`), FontSchema);

async function fixture(): Promise<Uint8Array> {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Report');
  setCellFont(wb, setCell(ws, 1, 1, 42), makeFont({ name: 'Arial', underline: 'double' }));
  addDxf(wb.styles, { font: makeFont({ underline: 'double' }) });
  const parts = unzipSync(await workbookToBytes(wb));
  const styles = parts['xl/styles.xml'];
  assert(styles, 'expected stylesheet part');
  const xml = strFromU8(styles);
  expect(xml.match(/<u val="double"\/>/g)).toHaveLength(2);
  parts['xl/styles.xml'] = strToU8(xml.replaceAll('<u val="double"/>', '<u val="none"/>'));
  return zipSync(parts);
}

describe('issue #206 — explicit no-underline fonts', () => {
  it.each([
    ['', undefined],
    ['<u/>', 'single'],
    ['<u val="none"/>', 'none'],
    ['<u val="single"/>', 'single'],
    ['<u val="double"/>', 'double'],
    ['<u val="singleAccounting"/>', 'singleAccounting'],
    ['<u val="doubleAccounting"/>', 'doubleAccounting'],
  ] as const)('preserves the meaning of %s', (xml, expected) => {
    const font = parseFont(xml);
    expect(font).toEqual({ name: 'Arial', ...(expected === undefined ? {} : { underline: expected }) });
    const saved = serializeXml(toTree(makeFont(font), FontSchema));
    expect(fromTree(parseXml(saved), FontSchema)).toEqual(font);
  });

  it.each(['wavy', 'None', ''])('still rejects invalid underline value %j', (value) => {
    expect(() => parseFont(`<u val="${value}"/>`)).toThrow(OpenXmlSchemaError);
  });

  it('does not render an underline for none, and preserves strike-through', () => {
    const font = parseFont('<u val="none"/>');
    expect(fontToCss(font)['text-decoration']).toBe('none');
    expect(fontToCss({ ...font, strike: true })['text-decoration']).toBe('line-through');
  });

  it('loads, edits and saves cell and differential fonts without losing none', async () => {
    const bytes = await fixture();
    expect((await validateXlsx(bytes)).issues).toEqual([]);
    const wb = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb.sheets[0];
    assert(sheet?.kind === 'worksheet', 'expected report worksheet');
    const cell = getCell(sheet.sheet, 1, 1);
    assert(cell, 'expected report cell');
    expect(cell.value).toBe(42);
    expect(getCellFont(wb, cell).underline).toBe('none');
    expect(getDxfs(wb.styles)[0]?.font?.underline).toBe('none');
    setBold(wb, cell);
    expect(getCellFont(wb, cell)).toMatchObject({ name: 'Arial', bold: true, underline: 'none' });
    const saved = await workbookToBytes(wb);
    expect((await validateXlsx(saved)).issues).toEqual([]);
    const back = await loadWorkbook(fromBuffer(saved));
    expect(back.styles.fonts).toEqual(wb.styles.fonts);
    expect(getDxfs(back.styles)).toEqual(getDxfs(wb.styles));
  });

  it('loads the same fonts and cell values through the streaming reader', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await fixture()));
    try {
      expect(wb.styles.fonts.some((font) => font.underline === 'none')).toBe(true);
      expect(getDxfs(wb.styles)[0]?.font?.underline).toBe('none');
      const rows: unknown[][] = [];
      for await (const row of wb.openWorksheet('Report').iterValues()) rows.push(row);
      expect(rows).toEqual([[42]]);
    } finally {
      await wb.close();
    }
  });
});
