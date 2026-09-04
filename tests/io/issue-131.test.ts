// Regression for https://github.com/office-kit/xlsx/issues/131 — openpyxl
// writes non-ASCII inline strings as numeric XML character references. They
// must be decoded like every other XML text-node entity.

import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load';
import { fromBuffer } from '../../src/io/node';

const enc = new TextEncoder();
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const fixture = zipSync({
  '[Content_Types].xml': enc.encode(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '</Types>',
  ),
  '_rels/.rels': enc.encode(
    `<Relationships xmlns="${PKG_REL_NS}">` +
      `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>` +
      '</Relationships>',
  ),
  'xl/workbook.xml': enc.encode(
    `<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
      '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
      '</workbook>',
  ),
  'xl/_rels/workbook.xml.rels': enc.encode(
    `<Relationships xmlns="${PKG_REL_NS}">` +
      `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>` +
      '</Relationships>',
  ),
  'xl/worksheets/sheet1.xml': enc.encode(
    `<worksheet xmlns="${MAIN_NS}"><sheetData><row r="1">` +
      '<c r="A1" t="inlineStr"><is><t>&#20219;&#x52A1;</t></is></c>' +
      '<c r="B1" t="s"><v>0</v></c>' +
      '</row></sheetData></worksheet>',
  ),
  'xl/sharedStrings.xml': enc.encode(`<sst xmlns="${MAIN_NS}" count="1" uniqueCount="1"><si><t>&#20219;&#x52A1;</t></si></sst>`),
});

describe('issue #131 — numeric XML character references', () => {
  it('decodes decimal and hexadecimal references in inline and shared strings', async () => {
    const workbook = await loadWorkbook(fromBuffer(fixture));
    const sheetRef = workbook.sheets[0];
    if (sheetRef?.kind !== 'worksheet') throw new Error('expected one worksheet');

    expect(sheetRef.sheet.rows.get(1)?.get(1)?.value).toBe('任务');
    expect(sheetRef.sheet.rows.get(1)?.get(2)?.value).toBe('任务');
  });
});
