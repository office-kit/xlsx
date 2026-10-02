import { strToU8, unzipSync, zipSync } from 'fflate';
import { assert, describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { fromTree, toTree } from '../../src/schema/serialize.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { getCellFont, setFontSize } from '../../src/styles/cell-style.js';
import { getDxfs } from '../../src/styles/differential.js';
import { FontSchema } from '../../src/styles/fonts.schema.js';
import { parseStylesheetXml } from '../../src/styles/stylesheet-reader.js';
import { stylesheetToBytes } from '../../src/styles/stylesheet-writer.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';
import { SHEET_MAIN_NS } from '../../src/xml/namespaces.js';
import { parseXml } from '../../src/xml/parser.js';
import { serializeXml } from '../../src/xml/serializer.js';
import { validateXlsx } from '../conformance/validate.js';

const toggles = [
  ['b', 'bold'],
  ['i', 'italic'],
  ['strike', 'strike'],
  ['outline', 'outline'],
  ['shadow', 'shadow'],
  ['condense', 'condense'],
  ['extend', 'extend'],
] as const;
const xfFlags = [
  'applyFont', 'applyFill', 'applyBorder', 'applyNumberFormat',
  'applyAlignment', 'applyProtection', 'pivotButton', 'quotePrefix',
] as const;
const parseFont = (children: string) => fromTree(parseXml(`<font xmlns="${SHEET_MAIN_NS}">${children}</font>`), FontSchema);
// The stylesheet from the report, with only whitespace added for readability.
const reportedStyles = `<styleSheet xml:space="preserve" xmlns="${SHEET_MAIN_NS}">
  <numFmts count="0"/>
  <fonts count="1"><font><b val="0"/><i val="0"/><strike val="0"/><u val="none"/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf xfId="0" fontId="0" numFmtId="0" fillId="0" borderId="0" applyFont="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotTableStyle1"/>
</styleSheet>`;

async function fixture(): Promise<Uint8Array> {
  const wb = createWorkbook();
  setCell(addWorksheet(wb, 'Report'), 1, 1, 42);
  const parts = unzipSync(await workbookToBytes(wb));
  parts['xl/styles.xml'] = strToU8(reportedStyles.replace(
    '<dxfs count="0"/>',
    '<dxfs count="1"><dxf><font><b val="false"/><i val="0"/><strike val="0"/></font></dxf></dxfs>',
  ));
  return zipSync(parts);
}

describe('issue #209 — stylesheet boolean values', () => {
  describe.each(toggles)('%s / %s', (tag, key) => {
    it.each([
      ['', undefined], [`<${tag}/>`, true], [`<${tag} val="1"/>`, true],
      [`<${tag} val="true"/>`, true], [`<${tag} val="0"/>`, false],
      [`<${tag} val="false"/>`, false], [`<${tag} val=" 0 "/>`, false],
    ] as const)('preserves %s through a round trip', (xml, expected) => {
      const font = parseFont(xml);
      expect(font).toEqual(expected === undefined ? {} : { [key]: expected });
      expect(fromTree(parseXml(serializeXml(toTree(font, FontSchema))), FontSchema)).toEqual(font);
    });

    it.each(['', '2', 'yes'])('rejects invalid val=%j', (value) => {
      expect(() => parseFont(`<${tag} val="${value}"/>`)).toThrow(OpenXmlSchemaError);
    });
  });

  it.each([undefined, false, true])('preserves XF flags set to %s in both XF tables', (value) => {
    const ss = parseStylesheetXml(strToU8(reportedStyles));
    const flags = value === undefined ? {} : Object.fromEntries(xfFlags.map((key) => [key, value]));
    const xf = { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, ...flags };
    ss.cellXfs = [xf];
    ss.cellStyleXfs = [xf];
    const back = parseStylesheetXml(stylesheetToBytes(ss));
    expect(back.cellXfs).toEqual([xf]);
    expect(back.cellStyleXfs).toEqual([xf]);
  });

  it('preserves the reported stylesheet and differential fonts through load, edit and save', async () => {
    const bytes = await fixture();
    expect((await validateXlsx(bytes)).issues).toEqual([]);
    const wb = await loadWorkbook(fromBuffer(bytes));
    const sheet = wb.sheets[0];
    assert(sheet?.kind === 'worksheet');
    const cell = getCell(sheet.sheet, 1, 1);
    assert(cell);
    const font = {
      name: 'Calibri', size: 11, color: { rgb: 'FF000000' },
      bold: false, italic: false, strike: false, underline: 'none',
    };
    expect(getCellFont(wb, cell)).toEqual(font);
    expect(getDxfs(wb.styles)[0]?.font).toEqual({ bold: false, italic: false, strike: false });
    const saved = await workbookToBytes(wb);
    expect((await validateXlsx(saved)).issues).toEqual([]);
    const back = await loadWorkbook(fromBuffer(saved));
    expect(back.styles.fonts).toEqual(wb.styles.fonts);
    expect(back.styles.cellXfs).toEqual(wb.styles.cellXfs);
    expect(back.styles.cellStyleXfs).toEqual(wb.styles.cellStyleXfs);
    expect(getDxfs(back.styles)).toEqual(getDxfs(wb.styles));
    setFontSize(wb, cell, 12);
    const edited = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    expect(edited.styles.fonts).toContainEqual({ ...font, size: 12 });
  });

  it('loads false font and XF values through the streaming reader', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await fixture()));
    try {
      expect(wb.styles.fonts[0]).toMatchObject({ bold: false, italic: false, strike: false });
      expect(wb.styles.cellXfs[0]).toMatchObject({ applyFont: false, applyAlignment: false });
      expect(getDxfs(wb.styles)[0]?.font).toEqual({ bold: false, italic: false, strike: false });
      const rows: unknown[][] = [];
      for await (const row of wb.openWorksheet('Report').iterValues()) rows.push(row);
      expect(rows).toEqual([[42]]);
    } finally {
      await wb.close();
    }
  });
});
