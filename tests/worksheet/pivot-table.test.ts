// PivotTable creation, rendering, save and reload.
//
// The expected report grids below were checked against Mac Excel: each
// configuration was written by this library, opened in Excel (which refreshes
// the pivot on load), saved back, and Excel's cells and <location> /
// <rowItems> / <colItems> compared with ours.

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { CellValue } from '../../src/cell/cell.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { getCellNumberFormat, setCellNumberFormat } from '../../src/styles/cell-style.js';
import { tupleToCoordinate } from '../../src/utils/coordinate.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import type { Workbook } from '../../src/workbook/workbook.js';
import { listPassthroughPivotTables } from '../../src/worksheet/pivot-reader.js';
import {
  addPivotTable,
  type AddPivotTableOptions,
  computePivotTable,
  getPivotTableAt,
  refreshPivotTable,
  removePivotTable,
} from '../../src/worksheet/pivot-table.js';
import { getCell, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';
import { validateXlsx } from '../conformance/validate.js';

const DATA: CellValue[][] = [
  ['Region', 'Product', 'Year', 'Sales', 'Qty'],
  ['East', 'Apple', 2023, 100, 3],
  ['West', 'Apple', 2023, 150, 5],
  ['East', 'Banana', 2024, 80, 2],
  ['West', 'Cherry', 2024, 60, 1],
  ['East', 'Cherry', 2023, 30, null],
  ['North', 'Banana', 2024, 45, 7],
  // Excel matches text items case-insensitively: this row joins "West".
  ['west', 'Banana', 2023, 20, 4],
];

function setup(opts: Omit<AddPivotTableOptions, 'source' | 'anchor'> = {}): { wb: Workbook; ws: Worksheet; src: Worksheet } {
  const wb = createWorkbook();
  const src = addWorksheet(wb, 'Data');
  DATA.forEach((row, r) => row.forEach((v, c) => setCell(src, r + 1, c + 1, v)));
  const ws = addWorksheet(wb, 'Pivot');
  addPivotTable(wb, ws, { ...opts, source: { sheet: 'Data', ref: 'A1:E8' }, anchor: 'A3' });
  return { wb, ws, src };
}

function grid(ws: Worksheet): string[] {
  const out: string[] = [];
  for (const r of [...ws.rows.keys()].sort((a, b) => a - b)) {
    const cells = [...(ws.rows.get(r)?.values() ?? [])].sort((a, b) => a.col - b.col);
    out.push(cells.map((c) => `${tupleToCoordinate(c.col, r)}=${String(c.value)}`).join(' '));
  }
  return out;
}

const part = (bytes: Uint8Array, path: string): string => {
  const entry = unzipSync(bytes)[path];
  if (entry === undefined) throw new Error(`missing ${path}`);
  return strFromU8(entry);
};

describe('addPivotTable report cells', () => {
  it('lays out rows × columns with two value fields (compact)', () => {
    const { ws } = setup({
      rows: [0],
      columns: [1],
      values: [
        { field: 3, aggregate: 'sum' },
        { field: 4, aggregate: 'max' },
      ],
    });
    expect(grid(ws)).toEqual([
      'B3=Column Labels',
      'B4=Apple D4=Banana F4=Cherry H4=Total Sum of Sales I4=Total Max of Qty',
      'A5=Row Labels B5=Sum of Sales C5=Max of Qty D5=Sum of Sales E5=Max of Qty F5=Sum of Sales G5=Max of Qty',
      'A6=East B6=100 C6=3 D6=80 E6=2 F6=30 H6=210 I6=3',
      'A7=North D7=45 E7=7 H7=45 I7=7',
      'A8=West B8=150 C8=5 D8=20 E8=4 F8=60 G8=1 H8=230 I8=5',
      'A9=Grand Total B9=250 C9=5 D9=145 E9=7 F9=90 G9=1 H9=485 I9=7',
    ]);
  });

  it('puts each row field in its own column with bottom subtotals (tabular)', () => {
    const { ws } = setup({ rows: [0, 1], values: [{ field: 3, aggregate: 'sum' }], layout: 'tabular' });
    expect(grid(ws)).toEqual([
      'A3=Region B3=Product C3=Sum of Sales',
      'A4=East B4=Apple C4=100',
      'B5=Banana C5=80',
      'B6=Cherry C6=30',
      'A7=East Total C7=210',
      'A8=North B8=Banana C8=45',
      'A9=North Total C9=45',
      'A10=West B10=Apple C10=150',
      'B11=Banana C11=20',
      'B12=Cherry C12=60',
      'A13=West Total C13=230',
      'A14=Grand Total C14=485',
    ]);
  });

  it('stacks report filters above the table and filters the records', () => {
    const { ws } = setup({ rows: [1], filters: [{ field: 0, selected: 'East' }], values: [{ field: 3, aggregate: 'sum' }] });
    expect(grid(ws)).toEqual([
      'A3=Region B3=East',
      'A5=Row Labels B5=Sum of Sales',
      'A6=Apple B6=100',
      'A7=Banana B7=80',
      'A8=Cherry B8=30',
      'A9=Grand Total B9=210',
    ]);
  });

  it('computes count / average / min, skipping blanks', () => {
    const { ws } = setup({
      values: [
        { field: 4, aggregate: 'count' },
        { field: 4, aggregate: 'average' },
        { field: 3, aggregate: 'min' },
      ],
    });
    expect(grid(ws)).toEqual([
      'B3=Column Labels',
      'B4=Count of Qty C4=Average of Qty D4=Min of Sales',
      'B5=6 C5=3.6666666666666665 D5=20',
    ]);
  });

  it('clears the previous report when refreshed with a smaller layout', () => {
    const { wb, ws } = setup({ rows: [0, 1], values: [{ field: 3, aggregate: 'sum' }] });
    const pt = ws.pivotTables?.[0];
    if (pt === undefined) throw new Error('no pivot');
    pt.rows = [0];
    refreshPivotTable(wb, ws, pt);
    expect(grid(ws)).toEqual([
      'A3=Row Labels B3=Sum of Sales',
      'A4=East B4=210',
      'A5=North B5=45',
      'A6=West B6=230',
      'A7=Grand Total B7=485',
    ]);
    expect(getPivotTableAt(ws, 7, 2)).toBe(pt);
    expect(getPivotTableAt(ws, 8, 2)).toBeUndefined();
    removePivotTable(ws, pt);
    expect(ws.rows.size).toBe(0);
    expect(ws.pivotTables).toEqual([]);
  });

  it('rejects a field placed on two axes, a blank header and an overlap with the source', () => {
    expect(() => setup({ rows: [0], columns: [0] })).toThrow(OpenXmlSchemaError);
    const { wb, src } = setup();
    setCell(src, 1, 6, null);
    expect(() =>
      addPivotTable(wb, src, { source: { sheet: 'Data', ref: 'A1:F8' }, anchor: 'H1' }),
    ).toThrow(OpenXmlSchemaError);
    expect(() => addPivotTable(wb, src, { source: { sheet: 'Data', ref: 'A1:E8' }, anchor: 'C2' })).toThrow(
      OpenXmlSchemaError,
    );
  });

  // Correctness at scale; the single-pass cost is guarded by tests/perf.
  it('aggregates 100k records', { timeout: 60_000 }, () => {
    const wb = createWorkbook();
    const src = addWorksheet(wb, 'Data');
    setCell(src, 1, 1, 'Key');
    setCell(src, 1, 2, 'Group');
    setCell(src, 1, 3, 'Value');
    const N = 100_000;
    for (let i = 0; i < N; i++) {
      setCell(src, i + 2, 1, `k${i % 500}`);
      setCell(src, i + 2, 2, `g${i % 7}`);
      setCell(src, i + 2, 3, i);
    }
    const ws = addWorksheet(wb, 'Pivot');
    const pt = addPivotTable(wb, ws, {
      source: { sheet: 'Data', ref: `A1:C${N + 1}` },
      anchor: 'A1',
      rows: [0],
      columns: [1],
      values: [{ field: 2, aggregate: 'sum' }],
    });
    expect(computePivotTable(wb, pt).rowEntries).toHaveLength(501);
    // Grand total of 0..N-1.
    expect(ws.rows.get(503)?.get(9)?.value).toBe((N * (N - 1)) / 2);
  });
});

describe('PivotTable save / load', () => {
  it('writes cache, records and table parts wired into workbook, rels and content types', async () => {
    const { wb } = setup({ rows: [0], columns: [1], values: [{ field: 3, aggregate: 'sum' }] });
    const bytes = await workbookToBytes(wb);
    expect(part(bytes, 'xl/workbook.xml')).toMatch(/<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/);
    expect(part(bytes, 'xl/_rels/workbook.xml.rels')).toContain('Target="pivotCache/pivotCacheDefinition1.xml"');
    expect(part(bytes, 'xl/worksheets/_rels/sheet2.xml.rels')).toContain('Target="../pivotTables/pivotTable1.xml"');
    expect(part(bytes, 'xl/pivotTables/_rels/pivotTable1.xml.rels')).toContain('Target="../pivotCache/pivotCacheDefinition1.xml"');
    expect(part(bytes, 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels')).toContain('Target="pivotCacheRecords1.xml"');
    const ct = part(bytes, '[Content_Types].xml');
    expect(ct).toContain('PartName="/xl/pivotTables/pivotTable1.xml"');
    expect(ct).toContain('PartName="/xl/pivotCache/pivotCacheRecords1.xml"');
    const cache = part(bytes, 'xl/pivotCache/pivotCacheDefinition1.xml');
    expect(cache).toContain('<worksheetSource ref="A1:E8" sheet="Data"/>');
    expect(cache).toContain('<sharedItems count="3"><s v="East"/><s v="North"/><s v="West"/></sharedItems>');
    expect(cache).toContain(
      '<cacheField name="Qty" numFmtId="0"><sharedItems containsString="0" containsBlank="1" containsNumber="1" containsInteger="1" minValue="1" maxValue="7"/>',
    );
    expect(part(bytes, 'xl/pivotCache/pivotCacheRecords1.xml')).toContain('<r><x v="0"/><x v="0"/><n v="2023"/><n v="100"/><n v="3"/></r>');
    const table = part(bytes, 'xl/pivotTables/pivotTable1.xml');
    expect(table).toContain('<location ref="A3:E8" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>');
    expect(table).toContain('<colItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i t="grand"><x/></i></colItems>');
    expect((await validateXlsx(bytes)).issues).toEqual([]);
  });

  it('reloads a pivot it wrote as an editable definition', async () => {
    const { wb } = setup({
      rows: [0, 1],
      filters: [{ field: 2, selected: 2024 }],
      values: [{ field: 3, aggregate: 'average' }],
      layout: 'outline',
      subtotals: 'bottom',
      columnGrandTotals: false,
    });
    const before = getSheet(wb, 'Pivot')?.pivotTables?.[0];
    const wb2 = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    const ws2 = getSheet(wb2, 'Pivot');
    expect(ws2?.pivotTables).toEqual([{ ...before, values: [{ field: 3, aggregate: 'average', name: 'Average of Sales' }] }]);
    expect(wb2.pivotCaches).toBeUndefined();
    expect([...(wb2.passthrough?.keys() ?? [])].filter((p) => p.includes('pivot'))).toEqual([]);
    // Saving again regenerates the parts rather than duplicating them.
    const again = await workbookToBytes(wb2);
    expect(Object.keys(unzipSync(again)).filter((p) => p.includes('pivot')).sort()).toEqual([
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
      'xl/pivotCache/pivotCacheDefinition1.xml',
      'xl/pivotCache/pivotCacheRecords1.xml',
      'xl/pivotTables/_rels/pivotTable1.xml.rels',
      'xl/pivotTables/pivotTable1.xml',
    ]);
  });

  it('lifts a pivot Excel saved, keeping its definition', async () => {
    const bytes = readFileSync(resolve(__dirname, '../fixtures/pivot/excel-saved-filter-pivot.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));
    const pt = getSheet(wb, 'Pivot')?.pivotTables?.[0];
    expect(pt).toMatchObject({
      name: 'PivotTable1',
      source: { sheet: 'Data', ref: 'A1:E8' },
      anchor: 'A3',
      rows: [1],
      columns: [],
      filters: [{ field: 0, selected: 'East' }],
      values: [{ field: 3, aggregate: 'sum', name: 'Sum of Sales' }],
      layout: 'compact',
      renderedRef: 'A3:B9',
    });
  });

  it('keeps a pivot it cannot model as passthrough and describes it', async () => {
    const bytes = readFileSync(resolve(__dirname, '../../reference/openpyxl/openpyxl/reader/tests/data/pivot.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));
    const ws = getSheet(wb, 'ptsheet');
    if (ws === undefined) throw new Error('no sheet');
    expect(ws.pivotTables).toBeUndefined();
    expect(listPassthroughPivotTables(wb, ws)).toEqual([
      {
        name: 'PivotTable1',
        ref: 'A3:E14',
        fields: ['ID', 'campaign_name', 'budget_date', 'hour', 'impressions', 'owner'],
        rows: ['campaign_name', 'owner'],
        columns: ['hour'],
        filters: [],
        values: ['Sum of impressions'],
      },
    ]);
  });

  it('numbers new parts past the passthrough pivot parts it carries', async () => {
    const bytes = readFileSync(resolve(__dirname, '../../reference/openpyxl/openpyxl/reader/tests/data/pivot.xlsx'));
    const wb = await loadWorkbook(fromBuffer(bytes));
    const ws = addWorksheet(wb, 'New');
    addPivotTable(wb, ws, { source: { sheet: 'raw', ref: 'A1:F18' }, anchor: 'A1', rows: [5], values: [{ field: 4, aggregate: 'sum' }] });
    const out = await workbookToBytes(wb);
    const files = Object.keys(unzipSync(out));
    expect(files).toContain('xl/pivotTables/pivotTable1.xml');
    expect(files).toContain('xl/pivotTables/pivotTable2.xml');
    expect(files).toContain('xl/pivotCache/pivotCacheDefinition2.xml');
    expect(part(out, 'xl/workbook.xml')).toContain('<pivotCaches><pivotCache cacheId="68" r:id="rId3"/><pivotCache cacheId="69"');
  });
});

describe('date items', () => {
  // Dates are serials in the sheet; without the source format the row labels read 45658, 45689.
  it('keep the source column\'s date format on their labels', () => {
    const wb = createWorkbook();
    const src = addWorksheet(wb, 'Data');
    setCell(src, 1, 1, 'Day');
    setCell(src, 1, 2, 'Sales');
    [45658, 45689, 45658].forEach((serial, i) => {
      setCellNumberFormat(wb, setCell(src, i + 2, 1, serial), 'yyyy-mm-dd');
      setCell(src, i + 2, 2, 10);
    });
    const ws = addWorksheet(wb, 'Pivot');
    addPivotTable(wb, ws, { source: { sheet: 'Data', ref: 'A1:B4' }, anchor: 'A3', rows: [0], values: [{ field: 1, aggregate: 'sum' }] });
    const first = getCell(ws, 4, 1);
    expect(first?.value).toBe(45658);
    expect(first && getCellNumberFormat(wb, first)).toBe('yyyy-mm-dd');
    // Values and captions keep the default format.
    const total = getCell(ws, 4, 2);
    expect(total && getCellNumberFormat(wb, total)).toBe('General');
  });
});
