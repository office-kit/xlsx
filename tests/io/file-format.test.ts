import { readFileSync } from 'node:fs';
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { XLSM_TYPE, XLSX_TYPE, XLTM_TYPE, XLTX_TYPE } from '../../src/xml/namespaces.js';

const workbookType = (bytes: Uint8Array): string | undefined => {
  const ct = new TextDecoder().decode(unzipSync(bytes)['[Content_Types].xml']);
  return /PartName="\/xl\/workbook\.xml" ContentType="([^"]+)"/.exec(ct)?.[1];
};

describe('workbook file format', () => {
  it('keeps a macro-enabled workbook without macros macro-enabled', async () => {
    const wb = await loadWorkbook(fromBuffer(readFileSync('reference/openpyxl/openpyxl/writer/tests/data/empty.xlsm')));
    expect(wb.vbaProject).toBeUndefined();
    expect(wb.fileFormat).toBe('xlsm');
    expect(workbookType(await workbookToBytes(wb))).toBe(XLSM_TYPE);
  });

  it.each([
    [undefined, false, XLSX_TYPE],
    ['xltx', false, XLTX_TYPE],
    ['xltx', true, XLTM_TYPE],
    ['xlsx', true, XLSM_TYPE],
  ] as const)('format %s with macros %s is written as %s and read back', async (format, macros, type) => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    if (format !== undefined) wb.fileFormat = format;
    if (macros) wb.vbaProject = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]);
    const bytes = await workbookToBytes(wb);
    expect(workbookType(bytes)).toBe(type);
    const back = await loadWorkbook(fromBuffer(bytes));
    expect(back.fileFormat ?? 'xlsx').toBe({ [XLSX_TYPE]: 'xlsx', [XLSM_TYPE]: 'xlsm', [XLTX_TYPE]: 'xltx', [XLTM_TYPE]: 'xltm' }[type]);
  });
});
