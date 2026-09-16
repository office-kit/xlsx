// Tests for listHyperlinks / listTables / listDataValidations / listDefinedNames.

import { describe, expect, it } from 'vitest';
import { addDefinedName, listDefinedNames } from '../../src/workbook/defined-names.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  addUrlHyperlink,
} from '../../src/worksheet/hyperlinks.js';
import {
  addListValidation,
} from '../../src/worksheet/data-validations.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import {
  listDataValidations,
  listHyperlinks,
  listTables,
  setCell,
  writeRange,
} from '../../src/worksheet/worksheet.js';

describe('listHyperlinks', () => {
  it('returns the hyperlinks array', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addUrlHyperlink(ws, 'A1', 'https://example.com');
    addUrlHyperlink(ws, 'A2', 'https://anthropic.com', { tooltip: 'docs' });
    const list = listHyperlinks(ws);
    expect(list.length).toBe(2);
    expect(list[0]?.ref).toBe('A1');
  });

  it('empty worksheet → empty array', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(listHyperlinks(ws)).toEqual([]);
  });
});

describe('listTables', () => {
  it('lists added tables', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'h');
    writeRange(ws, 'A1', [['col1', 'col2']]);
    addExcelTable(wb, ws, {
      name: 'Tbl',
      ref: 'A1:B2',
      columns: ['col1', 'col2'],
    });
    const list = listTables(ws);
    expect(list.length).toBe(1);
    expect(list[0]?.name).toBe('Tbl');
  });

  it('empty worksheet → empty array', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(listTables(ws)).toEqual([]);
  });
});

describe('listDataValidations', () => {
  it('lists added data validations', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    addListValidation(ws, 'A1:A10', ['Yes', 'No']);
    const list = listDataValidations(ws);
    expect(list.length).toBe(1);
    expect(list[0]?.type).toBe('list');
  });
});

describe('listDefinedNames', () => {
  it('returns all defined names by default', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'WbName', value: '$A$1' });
    addDefinedName(wb, { name: 'SheetName', value: '$A$1', scope: 0 });
    expect(listDefinedNames(wb).length).toBe(2);
  });

  it('scope: "workbook" filters to workbook-scope only', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'WbName', value: '$A$1' });
    addDefinedName(wb, { name: 'SheetName', value: '$A$1', scope: 0 });
    const wbOnly = listDefinedNames(wb, { scope: 'workbook' });
    expect(wbOnly.length).toBe(1);
    expect(wbOnly[0]?.name).toBe('WbName');
  });

  it('scope: <sheetIndex> filters to that sheet only', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'WbName', value: '$A$1' });
    addDefinedName(wb, { name: 'A0', value: '$A$1', scope: 0 });
    addDefinedName(wb, { name: 'A1', value: '$A$1', scope: 1 });
    const sheet0 = listDefinedNames(wb, { scope: 0 });
    expect(sheet0.map((d) => d.name)).toEqual(['A0']);
  });

  it('empty workbook → empty array', () => {
    const wb = createWorkbook();
    expect(listDefinedNames(wb)).toEqual([]);
  });
});
