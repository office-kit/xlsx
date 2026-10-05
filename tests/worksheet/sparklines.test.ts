import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { Chartsheet } from '../../src/chartsheet/chartsheet.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { makeSparklineGroup, type SparklineGroup } from '../../src/worksheet/sparklines.js';
import { setCell, type Worksheet } from '../../src/worksheet/worksheet.js';
import { openZip } from '../../src/zip/reader.js';
import { validateXlsx } from '../conformance/validate.js';

const enc = new TextEncoder();
const dec = new TextDecoder();

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';

const expectSheet = (ws: Worksheet | Chartsheet | undefined): Worksheet => {
  if (!ws || !('rows' in ws)) throw new Error('expected worksheet');
  return ws;
};

const sheetXml = async (bytes: Uint8Array): Promise<string> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    return dec.decode(archive.read('xl/worksheets/sheet1.xml'));
  } finally {
    archive.close();
  }
};

// Excel for Mac's output for a column sparkline group next to an x14
// data-validation extension, which has to survive untouched and stay ahead of
// the sparkline entry.
const DV_EXT =
  `<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14_NS}">` +
  `<x14:dataValidations count="1" xmlns:xm="${XM_NS}"><x14:dataValidation type="list" allowBlank="1">` +
  '<x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>H1</xm:sqref></x14:dataValidation></x14:dataValidations></ext>';
const SPARKLINE_EXT =
  `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14_NS}">` +
  `<x14:sparklineGroups xmlns:xm="${XM_NS}">` +
  '<x14:sparklineGroup type="column" displayEmptyCellsAs="gap" high="1" negative="1">' +
  '<x14:colorSeries theme="4" tint="-0.499984740745262"/><x14:colorNegative theme="5"/><x14:colorAxis rgb="FF000000"/>' +
  '<x14:colorMarkers theme="4" tint="-0.499984740745262"/><x14:colorFirst theme="4" tint="0.39997558519241921"/>' +
  '<x14:colorLast theme="4" tint="0.39997558519241921"/><x14:colorHigh rgb="FF00B050"/><x14:colorLow theme="4"/>' +
  '<x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:E1</xm:f><xm:sqref>F1</xm:sqref></x14:sparkline>' +
  '<x14:sparkline><xm:f>Sheet1!A2:E2</xm:f><xm:sqref>F2</xm:sqref></x14:sparkline></x14:sparklines>' +
  '</x14:sparklineGroup></x14:sparklineGroups></ext>';

const buildFixture = (extLst: string): Uint8Array =>
  zipSync({
    '[Content_Types].xml': enc.encode(
      `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    'xl/workbook.xml': enc.encode(
      `${XML_DECL}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': enc.encode(
      `${XML_DECL}<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    'xl/worksheets/sheet1.xml': enc.encode(
      `${XML_DECL}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` +
        `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><extLst>${extLst}</extLst></worksheet>`,
    ),
  });

describe('sparklines', () => {
  it('reads an Excel sparkline group out of the worksheet extLst', async () => {
    const wb = await loadWorkbook(fromBuffer(buildFixture(DV_EXT + SPARKLINE_EXT)));
    const ws = expectSheet(wb.sheets[0]?.sheet);
    expect(ws.sparklineGroups).toEqual([
      {
        type: 'column',
        high: true,
        negative: true,
        displayEmptyCellsAs: 'gap',
        colorSeries: { theme: 4, tint: -0.499984740745262 },
        colorNegative: { theme: 5 },
        colorAxis: { rgb: 'FF000000' },
        colorMarkers: { theme: 4, tint: -0.499984740745262 },
        colorFirst: { theme: 4, tint: 0.39997558519241921 },
        colorLast: { theme: 4, tint: 0.39997558519241921 },
        colorHigh: { rgb: 'FF00B050' },
        colorLow: { theme: 4 },
        sparklines: [
          { formula: 'Sheet1!A1:E1', location: 'F1' },
          { formula: 'Sheet1!A2:E2', location: 'F2' },
        ],
      },
    ]);
  });

  it('writes the sparkline entry back after the other extensions, which survive verbatim', async () => {
    const bytes = await workbookToBytes(await loadWorkbook(fromBuffer(buildFixture(SPARKLINE_EXT + DV_EXT))));
    const xml = await sheetXml(bytes);
    expect(xml.match(/<extLst>/g)).toHaveLength(1);
    expect(xml).toContain('Lists!$A$1:$A$3');
    expect(xml.indexOf('{CCE6A557')).toBeLessThan(xml.indexOf('{05C60535'));
    // CT_Worksheet puts extLst last.
    expect(xml).toMatch(/<\/extLst><\/worksheet>$/);
    const again = expectSheet((await loadWorkbook(fromBuffer(bytes))).sheets[0]?.sheet);
    expect(again.sparklineGroups?.[0]?.sparklines).toHaveLength(2);
    expect(again.sparklineGroups?.[0]?.colorHigh).toEqual({ rgb: 'FF00B050' });
  });

  it('drops the extLst once its only entry, the sparklines, is cleared', async () => {
    const wb = await loadWorkbook(fromBuffer(buildFixture(SPARKLINE_EXT)));
    expectSheet(wb.sheets[0]?.sheet).sparklineGroups = [];
    expect(await sheetXml(await workbookToBytes(wb))).not.toContain('extLst');
  });

  it('round-trips every group setting from a built workbook that passes the conformance validator', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Data');
    for (let c = 1; c <= 5; c++) {
      setCell(ws, 1, c, c * 2 - 5);
      setCell(ws, 2, c, (c % 2) * 2 - 1);
    }
    const groups: SparklineGroup[] = [
      makeSparklineGroup({
        sparklines: [{ formula: 'Data!A1:E1', location: 'F1' }],
        markers: true,
        first: true,
        last: true,
        low: true,
        lineWeight: 1.5,
        displayXAxis: true,
        minAxisType: 'custom',
        manualMin: -10,
        maxAxisType: 'group',
      }),
      makeSparklineGroup({ type: 'stacked', sparklines: [{ formula: 'Data!A2:E2', location: 'F2' }], negative: true }),
    ];
    ws.sparklineGroups = groups;
    const bytes = await workbookToBytes(wb);
    expect(await validateXlsx(bytes)).toMatchObject({ status: 'valid' });
    const again = expectSheet((await loadWorkbook(fromBuffer(bytes))).sheets[0]?.sheet);
    expect(again.sparklineGroups).toEqual(groups);
  });
});
