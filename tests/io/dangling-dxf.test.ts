import { describe, expect, it } from 'vitest';
import { workbookToBytes } from '../../src/io/save.js';
import { addDxf, makeDifferentialStyle } from '../../src/styles/differential.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { makeCfRule, makeConditionalFormatting } from '../../src/worksheet/conditional-formatting.js';

describe('saveWorkbook: conditional formatting dxfId', () => {
  it('refuses a rule pointing past the differential formats, which Excel will not open', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    ws.conditionalFormatting.push(makeConditionalFormatting({ sqref: 'A1:A5', rules: [makeCfRule({ type: 'cellIs', operator: 'greaterThan', priority: 1, dxfId: 0, formulas: ['3'] })] }));
    await expect(workbookToBytes(wb)).rejects.toThrow(OpenXmlSchemaError);
    await expect(workbookToBytes(wb)).rejects.toThrow(/sheet "S" conditional formatting A1:A5 uses dxfId 0/);
  });

  it('saves once the format is in the stylesheet', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    const dxfId = addDxf(wb.styles, makeDifferentialStyle({ font: { bold: true } }));
    ws.conditionalFormatting.push(makeConditionalFormatting({ sqref: 'A1:A5', rules: [makeCfRule({ type: 'cellIs', operator: 'greaterThan', priority: 1, dxfId, formulas: ['3'] })] }));
    await expect(workbookToBytes(wb)).resolves.toBeInstanceOf(Uint8Array);
  });
});
