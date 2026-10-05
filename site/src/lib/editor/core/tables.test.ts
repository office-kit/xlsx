import { addExcelTable, getCell, setCell } from '@office-kit/xlsx/worksheet';
import { describe, expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';
import { structuralEdit } from './structure.ts';
import {
  convertToRange,
  deleteTableRows,
  renameTable,
  renameTableInFormula,
  resizeTable,
  setHeaderRow,
  setTotalFunction,
  setTotalRow,
  structuredToA1Formula,
  tableNameError,
  totalCellAt,
} from './tables.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

function formula(ctl: EditorController, row: number, col: number): string | undefined {
  const v = getCell(ctl.doc.ws, row, col)?.value;
  return v && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.formula : undefined;
}

function value(ctl: EditorController, row: number, col: number): unknown {
  const v = getCell(ctl.doc.ws, row, col)?.value;
  if (v && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula') return v.cachedValue;
  return v;
}

/** Sales table in A1:B4 (Region / Sales), named Sales. */
function salesTable(): EditorController {
  const ctl = new EditorController();
  const rows = [
    ['Region', 'Sales'],
    ['East', '10'],
    ['West', '20'],
    ['North', '5'],
  ];
  rows.forEach((r, i) => r.forEach((text, j) => type(ctl, i + 1, j + 1, text)));
  ctl.doc.transact('Create Table', (tx) => {
    tx.sheet(ctl.doc.ws, 'tables');
    addExcelTable(ctl.doc.wb, ctl.doc.ws, { name: 'Sales', ref: 'A1:B4', columns: ['Region', 'Sales'], style: 'TableStyleMedium2', autoFilter: { ref: 'A1:B4', filterColumns: [] } });
  });
  return ctl;
}

const table = (ctl: EditorController) => {
  const def = ctl.doc.ws.tables[0];
  if (!def) throw new Error('no table');
  return def;
};

describe('table names', () => {
  it('rejects cell-like, malformed and taken names', () => {
    const ctl = salesTable();
    expect(tableNameError(ctl.doc.wb, 'T1')).toBe('dtTableNameInvalid');
    expect(tableNameError(ctl.doc.wb, 'R1C1')).toBe('dtTableNameInvalid');
    expect(tableNameError(ctl.doc.wb, 'My Table')).toBe('dtTableNameInvalid');
    expect(tableNameError(ctl.doc.wb, '1st')).toBe('dtTableNameInvalid');
    expect(tableNameError(ctl.doc.wb, 'sales')).toBe('dtTableNameTaken');
    expect(tableNameError(ctl.doc.wb, 'sales', table(ctl))).toBeUndefined();
    expect(tableNameError(ctl.doc.wb, '売上_2024')).toBeUndefined();
  });

  it('renames the table and the structured references to it', () => {
    const ctl = salesTable();
    type(ctl, 6, 1, '=SUM(Sales[Sales])');
    expect(renameTable(ctl, table(ctl), 'Revenue')).toBeUndefined();
    expect(table(ctl).displayName).toBe('Revenue');
    expect(formula(ctl, 6, 1)).toBe('SUM(Revenue[Sales])');
    expect(value(ctl, 6, 1)).toBe(35);
    A.undo(ctl);
    expect(table(ctl).displayName).toBe('Sales');
    expect(formula(ctl, 6, 1)).toBe('SUM(Sales[Sales])');
  });

  it('leaves strings and other tables alone when renaming', () => {
    expect(renameTableInFormula('"Sales[x]"&Sales[Region]&MySales[x]', 'Sales', 'T')).toBe('"Sales[x]"&T[Region]&MySales[x]');
  });
});

describe('structured references to A1', () => {
  const def = { id: 1, displayName: 'Sales', ref: 'B2:D6', columns: [{ id: 1, name: 'A' }, { id: 2, name: 'B' }, { id: 3, name: 'C' }], totalsRowCount: 1 };
  const range = { r1: 2, c1: 2, r2: 6, c2: 4 };
  it('resolves column, special-item and this-row forms', () => {
    expect(structuredToA1Formula('SUM(Sales[B])', def, range, 9, false)).toBe('SUM($C$3:$C$5)');
    expect(structuredToA1Formula('ROWS(Sales[#All])', def, range, 9, false)).toBe('ROWS($B$2:$D$6)');
    expect(structuredToA1Formula('Sales[[#Totals],[C]]', def, range, 9, false)).toBe('$D$6');
    expect(structuredToA1Formula('SUM(Sales[[A]:[B]])', def, range, 9, false)).toBe('SUM($B$3:$C$5)');
    expect(structuredToA1Formula('[@C]*2', def, range, 4, true)).toBe('$D4*2');
    expect(structuredToA1Formula('[@C]*2', def, range, 4, false)).toBe('[@C]*2');
  });
});

describe('total row', () => {
  it('adds a labelled total row with a SUBTOTAL that the drop-down can change', () => {
    const ctl = salesTable();
    setTotalRow(ctl, table(ctl), true, 'Total');
    expect(table(ctl).ref).toBe('A1:B5');
    expect(table(ctl).autoFilter?.ref).toBe('A1:B4');
    expect(value(ctl, 5, 1)).toBe('Total');
    expect(formula(ctl, 5, 2)).toBe('SUBTOTAL(109,Sales[Sales])');
    expect(value(ctl, 5, 2)).toBe(35);
    expect(totalCellAt(ctl.doc.ws, 5, 2)?.column.name).toBe('Sales');
    setTotalFunction(ctl, table(ctl), 2, 'max');
    expect(value(ctl, 5, 2)).toBe(20);
    setTotalRow(ctl, table(ctl), false, 'Total');
    expect(table(ctl).ref).toBe('A1:B4');
    expect(getCell(ctl.doc.ws, 5, 2)).toBeUndefined();
  });

  it('inserts cells below the table when the next row is taken', () => {
    const ctl = salesTable();
    // Set directly: typing it would grow the table to take it in.
    setCell(ctl.doc.ws, 5, 2, 'note');
    setTotalRow(ctl, table(ctl), true, 'Total');
    expect(value(ctl, 6, 2)).toBe('note');
    expect(formula(ctl, 5, 2)).toBe('SUBTOTAL(109,Sales[Sales])');
  });

  it('keeps the filter off the total row when rows move', () => {
    const ctl = salesTable();
    setTotalRow(ctl, table(ctl), true, 'Total');
    ctl.doc.transact('Insert', (tx) => structuralEdit(tx, ctl.doc.wb, ctl.doc.ws, { axis: 'row', at: 1, count: 2 }));
    expect(table(ctl).ref).toBe('A3:B7');
    expect(table(ctl).autoFilter?.ref).toBe('A3:B6');
  });
});

describe('header row', () => {
  it('removes the header row and writes it back above the data', () => {
    const ctl = salesTable();
    setHeaderRow(ctl, table(ctl), false);
    expect(table(ctl)).toMatchObject({ ref: 'A2:B4', headerRowCount: 0 });
    expect(table(ctl).autoFilter).toBeUndefined();
    expect(getCell(ctl.doc.ws, 1, 1)).toBeUndefined();
    setHeaderRow(ctl, table(ctl), true);
    expect(table(ctl).ref).toBe('A1:B4');
    expect(value(ctl, 1, 2)).toBe('Sales');
  });
});

describe('resize and convert', () => {
  it('grows a table by a column named from its header cell', () => {
    const ctl = salesTable();
    type(ctl, 1, 3, 'Qty');
    expect(resizeTable(ctl, table(ctl), { r1: 1, c1: 1, r2: 4, c2: 3 })).toBeUndefined();
    expect(table(ctl).columns.map((c) => c.name)).toEqual(['Region', 'Sales', 'Qty']);
    expect(resizeTable(ctl, table(ctl), { r1: 2, c1: 1, r2: 4, c2: 3 })).toBe('dtResizeHeaderRow');
  });

  it('converts to a range, keeping the look and turning references into A1', () => {
    const ctl = salesTable();
    type(ctl, 6, 1, '=SUM(Sales[Sales])');
    convertToRange(ctl, table(ctl));
    expect(ctl.doc.ws.tables).toHaveLength(0);
    expect(formula(ctl, 6, 1)).toBe('SUM($B$2:$B$4)');
    expect(value(ctl, 6, 1)).toBe(35);
    const header = ctl.doc.styles.get(getCell(ctl.doc.ws, 1, 1)?.styleId ?? 0);
    expect(header.bold).toBe(true);
    expect(header.fill).not.toBeNull();
    A.undo(ctl);
    expect(ctl.doc.ws.tables).toHaveLength(1);
    expect(ctl.doc.styles.get(getCell(ctl.doc.ws, 1, 1)?.styleId ?? 0).fill).toBeNull();
  });
});

describe('Delete Table Rows', () => {
  it('shifts only the table columns up and shrinks the ref and filter', () => {
    const ctl = salesTable();
    setCell(ctl.doc.ws, 3, 3, 'outside');
    ctl.selectCell({ row: 2, col: 1 });
    deleteTableRows(ctl, { r1: 2, c1: 1, r2: 2, c2: 1 });
    expect(table(ctl).ref).toBe('A1:B3');
    expect(table(ctl).autoFilter?.ref).toBe('A1:B3');
    expect(getCell(ctl.doc.ws, 2, 1)?.value).toBe('West');
    // Column C is outside the table's band and does not move.
    expect(getCell(ctl.doc.ws, 3, 3)?.value).toBe('outside');
    A.undo(ctl);
    expect(table(ctl).ref).toBe('A1:B4');
    expect(getCell(ctl.doc.ws, 2, 1)?.value).toBe('East');
  });

  it('keeps one empty data row when every data row is deleted', () => {
    const ctl = salesTable();
    ctl.selectCell({ row: 2, col: 1 });
    deleteTableRows(ctl, { r1: 1, c1: 1, r2: 4, c2: 2 });
    expect(table(ctl).ref).toBe('A1:B2');
    expect(getCell(ctl.doc.ws, 1, 1)?.value).toBe('Region');
    expect(getCell(ctl.doc.ws, 2, 1)?.value ?? null).toBeNull();
  });
});
