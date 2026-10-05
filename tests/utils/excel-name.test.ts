import { describe, expect, it } from 'vitest';
import { renameDefinedName } from '../../src/workbook/defined-names.js';
import { addDefinedName, addWorksheet, createWorkbook } from '../../src/workbook/index.js';
import { addExcelTable, writeRange } from '../../src/worksheet/index.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';

// Each of these makes Excel offer to repair the file on open.
const REJECTED = ['T1', 'xfd1048576', 'R2C3', 'r', 'C', 'R5', 'my name', '1abc', 'a-b', '', 'x'.repeat(256)];
const ACCEPTED = ['Sales', '_tmp', '\\back', 'a.b', 'Q?', 'XFE1', 'A1B', '売上', 'RC1x'];

describe('Excel name rules', () => {
  it('addDefinedName rejects names Excel would repair away', () => {
    const wb = createWorkbook();
    for (const name of REJECTED) expect(() => addDefinedName(wb, { name, value: '1' }), name).toThrow(OpenXmlSchemaError);
    for (const name of ACCEPTED) expect(() => addDefinedName(wb, { name, value: '1' }), name).not.toThrow();
    expect(() => addDefinedName(wb, { name: '_xlnm.Print_Area', value: 'Sheet1!$A$1', scope: 0 })).not.toThrow();
  });

  it('renameDefinedName validates the new name', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'Sales', value: '1' });
    expect(() => renameDefinedName(wb, 'Sales', 'B2')).toThrow(OpenXmlSchemaError);
    expect(renameDefinedName(wb, 'Sales', 'Revenue')).toBe(true);
  });

  it('tables reject a displayName that reads as a cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    writeRange(ws, 'A1', [['x', 'y']]);
    expect(() => addExcelTable(wb, ws, { name: 'T1', ref: 'A1:B2', columns: ['x', 'y'] })).toThrow(/reads as a cell reference/);
    expect(ws.tables).toEqual([]);
    expect(addExcelTable(wb, ws, { name: 'Sales', ref: 'A1:B2', columns: ['x', 'y'] }).displayName).toBe('Sales');
  });
});
