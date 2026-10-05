// PivotTables built from a worksheet range.
//
// A pivot is modelled as its *definition*: the source range, which source
// columns sit on which axis, how values aggregate, and the layout toggles.
// Everything Excel stores beside it (the pivot cache and its records, the
// row / column item lists, the rendered cells) is derived from that definition
// plus the current source data by `computePivotTable`. `refreshPivotTable`
// writes the derived cells into the host sheet; `saveWorkbook` derives the
// cache and table parts at save time through `computePivotStructure`, which
// skips the aggregation and report cells so they stay out of the io bundle.
//
// Aggregation is one pass over the source records: every record adds itself
// to each (row prefix, column prefix) bucket it belongs to, so the cost is
// records × (row depth + 1) × (column depth + 1), never records × output cells.

import type { CellValue, ExcelErrorCode } from '../cell/cell.js';
import type { Workbook } from '../workbook/workbook.js';
import { getSheet } from '../workbook/workbook.js';
import { isExcelErrorToken } from '../utils/cell-error.js';
import { coordinateToTuple, MAX_COL, MAX_ROW } from '../utils/coordinate.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { getCellNumberFormat, setCellNumberFormat } from '../styles/cell-style.js';
import { isDateFormat } from '../styles/numbers.js';
import { makeCellRange, parseRange, rangesOverlap, rangeToString } from './cell-range.js';
import { clearRange, setCell, type Worksheet } from './worksheet.js';

/** How a value field summarises its source column. */
export type PivotAggregate = 'sum' | 'count' | 'average' | 'max' | 'min';

/** Excel's three report layouts (Design ▸ Report Layout). */
export type PivotLayout = 'compact' | 'outline' | 'tabular';

/**
 * Where row-field subtotals go. Column fields always subtotal after their
 * items, and the tabular layout always subtotals at the bottom, as in Excel.
 */
export type PivotSubtotals = 'top' | 'bottom' | 'off';

/** A cell value as a pivot item sees it. `null` is the "(blank)" item. */
export type PivotItemValue = string | number | boolean | null;

export interface PivotValueField {
  /** 0-based column offset of the source field within `source.ref`. */
  field: number;
  aggregate: PivotAggregate;
  /** Caption, e.g. "Sum of Sales". Derived from the aggregate and field name when unset. */
  name?: string;
}

export interface PivotFilterField {
  /** 0-based column offset of the source field within `source.ref`. */
  field: number;
  /** The one item the report is filtered to; unset shows "(All)". */
  selected?: PivotItemValue;
}

export interface PivotTable {
  /** Workbook-unique name, e.g. "PivotTable1". */
  name: string;
  /** Source range including its header row; the header cells name the fields. */
  source: { sheet: string; ref: string };
  /** Top-left cell of the report on the host sheet, filter rows included. */
  anchor: string;
  /** Row fields, outermost first, as 0-based source column offsets. */
  rows: number[];
  /** Column fields, outermost first. */
  columns: number[];
  /** Report filters (Excel's page fields). */
  filters: PivotFilterField[];
  /** Value fields. Two or more lay out side by side across the columns. */
  values: PivotValueField[];
  layout: PivotLayout;
  subtotals: PivotSubtotals;
  rowGrandTotals: boolean;
  columnGrandTotals: boolean;
  /** Range the last refresh wrote into, so the next refresh can clear it first. */
  renderedRef?: string;
}

export interface AddPivotTableOptions {
  name?: string;
  source: { sheet: string; ref: string };
  anchor: string;
  rows?: number[];
  columns?: number[];
  filters?: PivotFilterField[];
  values?: PivotValueField[];
  layout?: PivotLayout;
  subtotals?: PivotSubtotals;
  rowGrandTotals?: boolean;
  columnGrandTotals?: boolean;
}

/** A source value as the pivot cache stores it. */
export type PivotCacheValue =
  | { t: 's'; v: string }
  | { t: 'n'; v: number }
  | { t: 'b'; v: boolean }
  | { t: 'e'; v: ExcelErrorCode }
  | { t: 'm' };

/** Type facts about one cache field, for the `<sharedItems>` attributes. */
interface PivotFieldStats {
  hasString: boolean;
  hasNumber: boolean;
  /** Every number is an integer (meaningful only when `hasNumber`). */
  allIntegers: boolean;
  hasBoolean: boolean;
  hasError: boolean;
  hasBlank: boolean;
  min: number;
  max: number;
}

export interface PivotCacheField {
  name: string;
  stats: PivotFieldStats;
  /** Per record, the source value. */
  records: PivotCacheValue[];
  /** Distinct items in display order; present only for fields on an axis or filter. */
  items?: PivotCacheValue[];
  /** Per record, the index into `items`; present with `items`. */
  itemIndex?: number[];
}

/** One `<i>` entry of `<rowItems>` / `<colItems>`. */
export interface PivotAxisEntry {
  type: 'data' | 'default' | 'grand';
  /** Leading item indices shared with the previous entry. */
  r: number;
  /** Item indices after the shared ones. */
  x: number[];
  /** Value field index this entry shows (only meaningful with several value fields). */
  i: number;
}

export interface PivotComputation {
  fields: PivotCacheField[];
  recordCount: number;
  /** Value field captions, in `values` order. */
  valueNames: string[];
  /** The column axis fields as written to `<colFields>`; -2 stands for the Σ Values field. */
  colFields: number[];
  rowEntries: PivotAxisEntry[];
  colEntries: PivotAxisEntry[];
  /** Table body (filter rows excluded), as written to `<location>`. */
  location: { ref: string; firstHeaderRow: number; firstDataRow: number; firstDataCol: number };
  /** Everything the report occupies, filter rows included. */
  outputRef: string;
  /** Cells to write, absolute coordinates. `numFmt` is set on item labels of date fields. */
  cells: Array<{ row: number; col: number; value: CellValue; numFmt?: string }>;
}

/** Excel's `<field x="-2"/>`: the Σ Values pseudo-field on an axis. */
export const PIVOT_VALUES_FIELD = -2;

const AGGREGATE_LABEL: Record<PivotAggregate, string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  max: 'Max',
  min: 'Min',
};

// Excel's English captions. Excel rewrites them in its UI language on the next
// refresh, so these only matter until then (and to non-Excel readers).
const ROW_LABELS = 'Row Labels';
const COLUMN_LABELS = 'Column Labels';
const GRAND_TOTAL = 'Grand Total';
const BLANK_ITEM = '(blank)';
const ALL_ITEMS = '(All)';
const VALUES_CAPTION = 'Values';

// A pivot with no fields placed yet still owns a block of the sheet: Excel
// draws its drop-zone placeholder over 3 columns × 18 rows.
const EMPTY_REPORT_COLS = 3;
const EMPTY_REPORT_ROWS = 18;

// ---- definition ----------------------------------------------------------

/**
 * Create a PivotTable on `ws` and write its report cells. Field references are
 * 0-based column offsets within `source.ref`, whose first row holds the field
 * names.
 */
export function addPivotTable(wb: Workbook, ws: Worksheet, opts: AddPivotTableOptions): PivotTable {
  const pt: PivotTable = {
    name: opts.name ?? nextPivotTableName(wb),
    source: { sheet: opts.source.sheet, ref: opts.source.ref },
    anchor: opts.anchor,
    rows: opts.rows ?? [],
    columns: opts.columns ?? [],
    filters: opts.filters ?? [],
    values: opts.values ?? [],
    layout: opts.layout ?? 'compact',
    subtotals: opts.subtotals ?? 'top',
    rowGrandTotals: opts.rowGrandTotals ?? true,
    columnGrandTotals: opts.columnGrandTotals ?? true,
  };
  for (const s of wb.sheets) {
    if (s.kind !== 'worksheet') continue;
    for (const other of s.sheet.pivotTables ?? []) {
      if (other.name.toLowerCase() === pt.name.toLowerCase()) {
        throw new OpenXmlSchemaError(`addPivotTable: a PivotTable named "${pt.name}" already exists`);
      }
    }
  }
  refreshPivotTable(wb, ws, pt);
  (ws.pivotTables ??= []).push(pt);
  return pt;
}

/** First free "PivotTableN" name in the workbook. */
export function nextPivotTableName(wb: Workbook): string {
  const taken = new Set<string>();
  for (const s of wb.sheets) {
    if (s.kind !== 'worksheet') continue;
    for (const pt of s.sheet.pivotTables ?? []) taken.add(pt.name.toLowerCase());
  }
  let n = 1;
  while (taken.has(`pivottable${n}`)) n++;
  return `PivotTable${n}`;
}

/**
 * Recompute `pt` from its source and rewrite its report cells on `ws`: the
 * range the previous refresh wrote is cleared first. Call after changing the
 * definition or the source data.
 */
export function refreshPivotTable(wb: Workbook, ws: Worksheet, pt: PivotTable): PivotComputation {
  const result = computePivotTable(wb, pt);
  const sourceWs = getSheet(wb, pt.source.sheet);
  if (sourceWs === ws && rangesOverlap(parseRange(result.outputRef), parseRange(pt.source.ref))) {
    throw new OpenXmlSchemaError(
      `refreshPivotTable: "${pt.name}" at ${result.outputRef} would overwrite its source ${pt.source.ref}`,
    );
  }
  if (pt.renderedRef !== undefined) clearRange(ws, pt.renderedRef);
  for (const c of result.cells) {
    const cell = setCell(ws, c.row, c.col, c.value);
    if (c.numFmt !== undefined) setCellNumberFormat(wb, cell, c.numFmt);
  }
  pt.renderedRef = result.outputRef;
  return result;
}

/**
 * The range a refresh of `pt` would write, filter rows included, without
 * writing it — for checking what the report would cover first.
 */
export function getPivotTableOutputRef(wb: Workbook, pt: PivotTable): string {
  return computePivotTable(wb, pt).outputRef;
}

/** Distinct items of source field `field` in the order the report lists them, e.g. for a filter drop-down. */
export function getPivotFieldItems(wb: Workbook, source: { sheet: string; ref: string }, field: number): PivotItemValue[] {
  const ws = getSheet(wb, source.sheet);
  if (ws === undefined) throw new OpenXmlSchemaError(`PivotTable source sheet "${source.sheet}" does not exist`);
  const range = parseRange(source.ref);
  const col = range.minCol + field;
  if (!Number.isInteger(field) || col > range.maxCol || field < 0) {
    throw new OpenXmlSchemaError(`PivotTable source ${source.ref} has no field ${field}`);
  }
  const values: PivotCacheValue[] = [];
  for (let row = range.minRow + 1; row <= range.maxRow; row++) values.push(toCacheValue(ws.rows.get(row)?.get(col)?.value ?? null, wb));
  return itemize(values).items.map(pivotItemValue);
}

/** Remove `pt` from `ws` and clear the cells its last refresh wrote. */
export function removePivotTable(ws: Worksheet, pt: PivotTable): void {
  const i = ws.pivotTables?.indexOf(pt) ?? -1;
  if (i < 0) return;
  if (pt.renderedRef !== undefined) clearRange(ws, pt.renderedRef);
  ws.pivotTables?.splice(i, 1);
}

/** The PivotTable on `ws` whose report covers (row, col), if any. */
export function getPivotTableAt(ws: Worksheet, row: number, col: number): PivotTable | undefined {
  for (const pt of ws.pivotTables ?? []) {
    if (pt.renderedRef === undefined) continue;
    const r = parseRange(pt.renderedRef);
    if (row >= r.minRow && row <= r.maxRow && col >= r.minCol && col <= r.maxCol) return pt;
  }
  return undefined;
}

/**
 * Field names for a source range: its header row, with duplicates numbered
 * the way Excel does ("Sales", "Sales2"). Throws when the sheet is missing or
 * a header cell is blank, since Excel requires every field to have a name.
 */
export function getPivotSourceFields(wb: Workbook, source: { sheet: string; ref: string }): string[] {
  const names = sourceFieldNames(wb, source);
  if (names === undefined) {
    throw new OpenXmlSchemaError(
      `PivotTable source ${source.sheet}!${source.ref}: the sheet is missing or a header cell is empty; every field needs a name`,
    );
  }
  return names;
}

/** {@link getPivotSourceFields}, answering `undefined` where that throws. */
export function sourceFieldNames(wb: Workbook, source: { sheet: string; ref: string }): string[] | undefined {
  const ws = getSheet(wb, source.sheet);
  if (ws === undefined) return undefined;
  const range = parseRange(source.ref);
  const names: string[] = [];
  const taken = new Set<string>();
  for (let col = range.minCol; col <= range.maxCol; col++) {
    const base = displayText(ws.rows.get(range.minRow)?.get(col)?.value ?? null).trim();
    if (base.length === 0) return undefined;
    let name = base;
    for (let n = 2; taken.has(name.toLowerCase()); n++) name = `${base}${n}`;
    taken.add(name.toLowerCase());
    names.push(name);
  }
  return names;
}

// ---- computation ---------------------------------------------------------

interface Accumulator {
  sum: number;
  numbers: number;
  /** Non-blank values, the "Count" aggregate. */
  count: number;
  min: number;
  max: number;
}

interface AxisNode {
  /** Per-axis serial number; the root is 0. Buckets are keyed by (row id, column id). */
  id: number;
  children: Map<number, AxisNode>;
}

/** What a save writes: the cache, the item lists and the location; no report cells. */
export type PivotStructure = Omit<PivotComputation, 'cells'>;

/** Called once per record that passes the report filters, with that record's row and column node ids. */
type RecordVisitor = (record: number, rowIds: readonly number[], colIds: readonly number[]) => void;
/** Builds the visitor once the cache fields are read. */
type MakeVisitor = (fields: PivotCacheField[]) => RecordVisitor;

interface DerivedPivot extends PivotStructure {
  names: string[];
  rowTree: AxisLine[];
  colTree: AxisLine[];
  geometry: ReportGeometry;
}

/**
 * The cache, item lists and location of `pt`, from its definition and the
 * current source data. `saveWorkbook` calls this on every save, so the cache it
 * writes always matches the sheet even when no refresh ran after an edit; it
 * leaves out the report cells, which only a refresh writes.
 */
export function computePivotStructure(wb: Workbook, pt: PivotTable): PivotStructure {
  const { names: _names, rowTree: _rows, colTree: _cols, geometry: _geometry, ...structure } = derivePivot(wb, pt);
  return structure;
}

/** Derive the cache, item lists, location and report cells for `pt`. */
export function computePivotTable(wb: Workbook, pt: PivotTable): PivotComputation {
  // Indexed by row node id (dense), then keyed by column node id.
  const buckets: Array<Map<number, Accumulator[]>> = [];
  const D = pt.values.length;
  const derived = derivePivot(wb, pt, (fields) => {
    const valueRecords = pt.values.map((v) => fields[v.field]?.records ?? []);
    return (i, rowIds, colIds) => {
      for (const rk of rowIds) {
        let row = buckets[rk];
        if (row === undefined) {
          row = new Map();
          buckets[rk] = row;
        }
        for (const ck of colIds) {
          let accs = row.get(ck);
          if (accs === undefined) {
            accs = pt.values.map(() => ({ sum: 0, numbers: 0, count: 0, min: Infinity, max: -Infinity }));
            row.set(ck, accs);
          }
          for (let v = 0; v < D; v++) {
            const value = valueRecords[v]?.[i];
            const acc = accs[v];
            if (value !== undefined && acc !== undefined) accumulate(acc, value);
          }
        }
      }
    };
  });
  const { names, rowTree, colTree, geometry, ...structure } = derived;
  const cells = layoutReport(pt, geometry, {
    fields: structure.fields,
    names,
    valueNames: structure.valueNames,
    colFields: structure.colFields,
    rowTree,
    colTree,
    buckets,
    itemFormats: itemFormats(wb, pt, structure.fields),
  });
  return { ...structure, cells };
}

/**
 * Dates are stored as serials, so their items would read as plain numbers
 * without the source column's date format: the first numeric record of each
 * axis or filter field decides.
 */
function itemFormats(wb: Workbook, pt: PivotTable, fields: PivotCacheField[]): Array<string | undefined> {
  const sourceWs = getSheet(wb, pt.source.sheet);
  const src = parseRange(pt.source.ref);
  return fields.map((field, f) => {
    if (field.items === undefined) return undefined;
    const i = field.records.findIndex((v) => v.t === 'n');
    const cell = i < 0 ? undefined : sourceWs?.rows.get(src.minRow + 1 + i)?.get(src.minCol + f);
    if (cell === undefined) return undefined;
    const fmt = getCellNumberFormat(wb, cell);
    return isDateFormat(fmt) ? fmt : undefined;
  });
}

function derivePivot(
  wb: Workbook,
  pt: PivotTable,
  makeVisitor?: MakeVisitor,
): DerivedPivot {
  const sourceWs = getSheet(wb, pt.source.sheet);
  if (sourceWs === undefined) {
    throw new OpenXmlSchemaError(`PivotTable "${pt.name}": source sheet "${pt.source.sheet}" does not exist`);
  }
  const src = parseRange(pt.source.ref);
  if (src.maxRow <= src.minRow) {
    throw new OpenXmlSchemaError(`PivotTable "${pt.name}": source ${pt.source.ref} needs a header row and at least one data row`);
  }
  const names = getPivotSourceFields(wb, pt.source);
  validateDefinition(pt, names.length);

  // Fields that need distinct items: everything on an axis or a filter.
  const itemized = new Set<number>([...pt.rows, ...pt.columns, ...pt.filters.map((f) => f.field)]);
  const recordCount = src.maxRow - src.minRow;
  const fields: PivotCacheField[] = names.map((name, f) => {
    // Grown by push: `new Array(n)` past ~100k elements starts out in V8's
    // slow dictionary mode, which made a 100k-row source several times slower.
    const records: PivotCacheValue[] = [];
    const stats: PivotFieldStats = {
      hasString: false,
      hasNumber: false,
      allIntegers: true,
      hasBoolean: false,
      hasError: false,
      hasBlank: false,
      min: Infinity,
      max: -Infinity,
    };
    const col = src.minCol + f;
    for (let i = 0; i < recordCount; i++) {
      const v = toCacheValue(sourceWs.rows.get(src.minRow + 1 + i)?.get(col)?.value ?? null, wb);
      records.push(v);
      noteStats(stats, v);
    }
    if (!itemized.has(f)) return { name, stats, records };
    return { name, stats, records, ...itemize(records) };
  });
  const visit = makeVisitor?.(fields);

  const valueNames = valueFieldNames(pt, names);
  const D = pt.values.length;
  const colFields = D > 1 ? [...pt.columns, PIVOT_VALUES_FIELD] : [...pt.columns];

  // Records surviving the report filters.
  const pageChecks = pt.filters
    .filter((f) => f.selected !== undefined)
    .map((f) => {
      const field = fields[f.field];
      const want = itemKey(fromItemValue(f.selected ?? null));
      const idx = field?.items?.findIndex((it) => itemKey(it) === want) ?? -1;
      return { itemIndex: field?.itemIndex ?? [], idx };
    });

  // One pass: build the row / column trees and hand each record to `visit`.
  // Numeric node ids rather than joined item paths as bucket keys: building a
  // key string per (record, row prefix, column prefix) dominated the pass.
  const rowRoot: AxisNode = { id: 0, children: new Map() };
  const colRoot: AxisNode = { id: 0, children: new Map() };
  const rowCount = { next: 1 };
  const colCount = { next: 1 };
  const rowIds: number[] = pt.rows.map(() => 0).concat(0);
  const colIds: number[] = pt.columns.map(() => 0).concat(0);
  const rowRecords = pt.rows.map((f) => fields[f]?.itemIndex ?? []);
  const colRecords = pt.columns.map((f) => fields[f]?.itemIndex ?? []);
  const passesFilters = (i: number): boolean => pageChecks.every((check) => check.itemIndex[i] === check.idx);
  for (let i = 0; i < recordCount; i++) {
    if (!passesFilters(i)) continue;
    let node = rowRoot;
    for (let d = 0; d < rowRecords.length; d++) {
      node = childOf(node, rowRecords[d]?.[i] ?? 0, rowCount);
      rowIds[d + 1] = node.id;
    }
    node = colRoot;
    for (let d = 0; d < colRecords.length; d++) {
      node = childOf(node, colRecords[d]?.[i] ?? 0, colCount);
      colIds[d + 1] = node.id;
    }
    visit?.(i, rowIds, colIds);
  }

  const rowTree = flattenAxis(rowRoot, pt.rows.length, 0, pt.layout, pt.subtotals, true);
  const colTree = flattenAxis(colRoot, pt.columns.length, D > 1 ? D : 0, 'tabular', pt.subtotals === 'off' ? 'off' : 'bottom', false);
  if (pt.rows.length > 0 && pt.rowGrandTotals) rowTree.push({ type: 'grand', path: [], dataIndex: 0, depth: 0, node: 0 });
  if (pt.columns.length > 0 && pt.columnGrandTotals) {
    for (let v = 0; v < Math.max(D, 1); v++) colTree.push({ type: 'grand', path: [], dataIndex: v, depth: 0, node: 0 });
  }
  // An axis with no fields still has one (empty) entry: the single value row / column.
  if (rowTree.length === 0) rowTree.push({ type: 'data', path: [], dataIndex: 0, depth: 0, node: 0 });
  if (colTree.length === 0) colTree.push({ type: 'data', path: [], dataIndex: 0, depth: 0, node: 0 });

  const geometry = reportGeometry(pt, colFields.length, rowTree.length, colTree.length);
  return {
    fields,
    recordCount,
    valueNames,
    colFields,
    rowEntries: toAxisEntries(rowTree),
    colEntries: toAxisEntries(colTree),
    location: geometry.location,
    outputRef: geometry.outputRef,
    names,
    rowTree,
    colTree,
    geometry,
  };
}

function validateDefinition(pt: PivotTable, fieldCount: number): void {
  const check = (f: number, where: string): void => {
    if (!Number.isInteger(f) || f < 0 || f >= fieldCount) {
      throw new OpenXmlSchemaError(`PivotTable "${pt.name}": ${where} field ${f} is outside the ${fieldCount} source fields`);
    }
  };
  const placed = new Set<number>();
  const place = (f: number, where: string): void => {
    check(f, where);
    if (placed.has(f)) {
      throw new OpenXmlSchemaError(`PivotTable "${pt.name}": field ${f} is placed on more than one of rows / columns / filters`);
    }
    placed.add(f);
  };
  for (const f of pt.rows) place(f, 'row');
  for (const f of pt.columns) place(f, 'column');
  for (const f of pt.filters) place(f.field, 'filter');
  for (const v of pt.values) check(v.field, 'value');
  coordinateToTuple(pt.anchor);
}

function childOf(node: AxisNode, x: number, counter: { next: number }): AxisNode {
  let child = node.children.get(x);
  if (child === undefined) {
    child = { id: counter.next++, children: new Map() };
    node.children.set(x, child);
  }
  return child;
}

function accumulate(acc: Accumulator, v: PivotCacheValue): void {
  if (v.t === 'm') return;
  acc.count++;
  if (v.t !== 'n') return;
  acc.sum += v.v;
  acc.numbers++;
  if (v.v < acc.min) acc.min = v.v;
  if (v.v > acc.max) acc.max = v.v;
}

function aggregateValue(acc: Accumulator | undefined, aggregate: PivotAggregate): CellValue {
  if (acc === undefined || acc.count === 0) return null;
  switch (aggregate) {
    case 'sum':
      return acc.sum;
    case 'count':
      return acc.count;
    case 'average':
      return acc.numbers === 0 ? { kind: 'error', code: '#DIV/0!' } : acc.sum / acc.numbers;
    case 'max':
      return acc.numbers === 0 ? 0 : acc.max;
    case 'min':
      return acc.numbers === 0 ? 0 : acc.min;
  }
}

function toCacheValue(value: CellValue, wb: Workbook): PivotCacheValue {
  if (value === null) return { t: 'm' };
  if (typeof value === 'number') return { t: 'n', v: value };
  if (typeof value === 'boolean') return { t: 'b', v: value };
  if (typeof value === 'string') return value.length === 0 ? { t: 'm' } : { t: 's', v: value };
  if (value instanceof Date) {
    // Dates group as their serial numbers; the cache has no date-typed items
    // in this writer, and Excel re-types them on its next refresh.
    const epoch = wb.date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    return { t: 'n', v: (value.getTime() - epoch) / 86_400_000 };
  }
  switch (value.kind) {
    case 'duration':
      return { t: 'n', v: value.ms / 86_400_000 };
    case 'error':
      return { t: 'e', v: value.code };
    case 'rich-text':
      return { t: 's', v: value.runs.map((r) => r.text).join('') };
    case 'formula': {
      const cached = value.cachedValue;
      if (cached === undefined) return { t: 'm' };
      if (value.cachedValueType === 'error' && typeof cached === 'string' && isExcelErrorToken(cached)) {
        return { t: 'e', v: cached };
      }
      return toCacheValue(cached, wb);
    }
  }
}

function noteStats(s: PivotFieldStats, v: PivotCacheValue): void {
  switch (v.t) {
    case 's':
      s.hasString = true;
      return;
    case 'n':
      s.hasNumber = true;
      if (!Number.isInteger(v.v)) s.allIntegers = false;
      if (v.v < s.min) s.min = v.v;
      if (v.v > s.max) s.max = v.v;
      return;
    case 'b':
      s.hasBoolean = true;
      return;
    case 'e':
      s.hasError = true;
      return;
    case 'm':
      s.hasBlank = true;
  }
}

// Excel matches text items case-insensitively: "east" and "East" are one item.
function itemKey(v: PivotCacheValue): string {
  switch (v.t) {
    case 's':
      return `s${v.v.toLowerCase()}`;
    case 'n':
      return `n${v.v}`;
    case 'b':
      return v.v ? 'b1' : 'b0';
    case 'e':
      return `e${v.v}`;
    case 'm':
      return 'm';
  }
}

function fromItemValue(v: PivotItemValue): PivotCacheValue {
  if (v === null) return { t: 'm' };
  if (typeof v === 'number') return { t: 'n', v };
  if (typeof v === 'boolean') return { t: 'b', v };
  return v.length === 0 ? { t: 'm' } : { t: 's', v };
}

/** The public form of a cache item, for filter dropdowns. Errors read as their code. */
function pivotItemValue(v: PivotCacheValue): PivotItemValue {
  return v.t === 'm' ? null : v.v;
}

// Excel's default item order: numbers, text, booleans, errors, then (blank).
const TYPE_RANK: Record<PivotCacheValue['t'], number> = { n: 0, s: 1, b: 2, e: 3, m: 4 };

function compareItems(a: PivotCacheValue, b: PivotCacheValue): number {
  if (a.t !== b.t) return TYPE_RANK[a.t] - TYPE_RANK[b.t];
  if (a.t === 'n' && b.t === 'n') return a.v - b.v;
  if (a.t === 's' && b.t === 's') {
    const x = a.v.toLowerCase();
    const y = b.v.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (a.t === 'b' && b.t === 'b') return Number(a.v) - Number(b.v);
  if (a.t === 'e' && b.t === 'e') return a.v < b.v ? -1 : a.v > b.v ? 1 : 0;
  return 0;
}

/** Distinct values of a field in display order, and each record's index into them. */
function itemize(values: PivotCacheValue[]): { items: PivotCacheValue[]; itemIndex: number[] } {
  const firstSeen = new Map<string, number>();
  const distinct: Array<{ value: PivotCacheValue; rank: number }> = [];
  const seen = values.map((value) => {
    const key = itemKey(value);
    let idx = firstSeen.get(key);
    if (idx === undefined) {
      idx = distinct.length;
      firstSeen.set(key, idx);
      distinct.push({ value, rank: 0 });
    }
    return idx;
  });
  const sorted = [...distinct].sort((a, b) => compareItems(a.value, b.value));
  sorted.forEach((d, rank) => {
    d.rank = rank;
  });
  return {
    items: sorted.map((d) => d.value),
    itemIndex: seen.map((idx) => distinct[idx]?.rank ?? 0),
  };
}

function valueFieldNames(pt: PivotTable, fieldNames: string[]): string[] {
  // Excel rejects a value caption that repeats another caption or a field name.
  const taken = new Set(fieldNames.map((n) => n.toLowerCase()));
  return pt.values.map((v) => {
    const base = v.name ?? `${AGGREGATE_LABEL[v.aggregate]} of ${fieldNames[v.field] ?? ''}`;
    let name = base;
    for (let n = 2; taken.has(name.toLowerCase()); n++) name = `${base}${n}`;
    taken.add(name.toLowerCase());
    return name;
  });
}

// ---- axes ----------------------------------------------------------------

interface AxisLine {
  type: 'data' | 'default' | 'grand';
  /** Item indices from the outermost field; for 'data' lines with Σ, excludes the value index. */
  path: number[];
  /** Value field index (Σ lines, per-value subtotal / grand columns). */
  dataIndex: number;
  /** Field level the line's label sits at (0-based). */
  depth: number;
  /** Axis node whose bucket the line shows (0, the root, for grand totals). */
  node: number;
  /** Compact / outline row that heads a group: no values unless subtotals are on top. */
  head?: boolean;
  /** Σ is a level of this axis: the line's path is followed by `dataIndex`. */
  sigma?: boolean;
}

/**
 * Walk an axis tree into display lines. `sigma` > 0 appends a Σ Values level
 * with that many children under every innermost item. Rows in compact / outline
 * layout get a head line per group; tabular rows and all columns put the group
 * label on the first leaf line instead.
 */
function flattenAxis(
  root: AxisNode,
  depth: number,
  sigma: number,
  layout: PivotLayout,
  subtotals: PivotSubtotals,
  isRow: boolean,
): AxisLine[] {
  const lines: AxisLine[] = [];
  if (depth === 0) {
    for (let v = 0; v < sigma; v++) lines.push({ type: 'data', path: [], dataIndex: v, depth: 0, node: 0, sigma: true });
    return lines;
  }
  const headed = isRow && layout !== 'tabular';
  const subtotalAfter = subtotals === 'bottom' || (subtotals === 'top' && !headed);
  const walk = (parent: AxisNode, prefix: number[]): void => {
    const level = prefix.length;
    const sorted = [...parent.children].sort((a, b) => a[0] - b[0]);
    for (const [x, child] of sorted) {
      const path = [...prefix, x];
      const node = child.id;
      if (level === depth - 1) {
        if (sigma > 0) {
          for (let v = 0; v < sigma; v++) lines.push({ type: 'data', path, dataIndex: v, depth: level, node, sigma: true });
        } else {
          lines.push({ type: 'data', path, dataIndex: 0, depth: level, node });
        }
        continue;
      }
      if (headed) lines.push({ type: 'data', path, dataIndex: 0, depth: level, node, head: true });
      walk(child, path);
      if (subtotals !== 'off' && subtotalAfter) {
        for (let v = 0; v < Math.max(sigma, 1); v++) lines.push({ type: 'default', path, dataIndex: v, depth: level, node });
      }
    }
  };
  walk(root, []);
  return lines;
}

function fullPath(line: AxisLine): number[] {
  return line.sigma ? [...line.path, line.dataIndex] : line.path;
}

function toAxisEntries(lines: AxisLine[]): PivotAxisEntry[] {
  let prev: number[] = [];
  return lines.map((line) => {
    if (line.type === 'grand') {
      prev = [];
      return { type: 'grand', r: 0, x: [0], i: line.dataIndex };
    }
    const path = fullPath(line);
    let common = 0;
    while (common < path.length - 1 && common < prev.length && path[common] === prev[common]) common++;
    prev = path;
    return { type: line.type, r: common, x: path.slice(common), i: line.dataIndex };
  });
}

// ---- report layout -------------------------------------------------------

/** Where the report sits; a save needs `location` and `outputRef`, a refresh all of it. */
interface ReportGeometry {
  location: PivotComputation['location'];
  outputRef: string;
  anchorRow: number;
  top: number;
  left: number;
  valueCols: number;
  dataTop: number;
  dataLeft: number;
  /** No field placed anywhere: Excel's empty drop-zone block. */
  empty: boolean;
}

function reportGeometry(pt: PivotTable, colFieldCount: number, rowLines: number, colLines: number): ReportGeometry {
  const { col: anchorCol, row: anchorRow } = coordinateToTuple(pt.anchor);
  // Report filters stack above the table, then one spacer row.
  const top = anchorRow + (pt.filters.length > 0 ? pt.filters.length + 1 : 0);
  const left = anchorCol;
  const R = pt.rows.length;
  const D = pt.values.length;
  const C = colFieldCount;

  if (R === 0 && C === 0 && D === 0) {
    const body = makeCellRange(top, left, clampRow(top + EMPTY_REPORT_ROWS - 1), clampCol(left + EMPTY_REPORT_COLS - 1));
    return {
      location: { ref: rangeToString(body), firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 0 },
      outputRef: rangeToString(makeCellRange(anchorRow, anchorCol, body.maxRow, body.maxCol)),
      anchorRow,
      top,
      left,
      valueCols: 0,
      dataTop: top,
      dataLeft: left,
      empty: true,
    };
  }

  const labelCols = R === 0 ? (C > 0 ? 1 : 0) : pt.layout === 'compact' ? 1 : R;
  const headerRows = C === 0 ? 1 : C + 1;
  const valueCols = D === 0 && C === 0 ? 0 : colLines;
  const dataTop = top + headerRows;
  const dataLeft = left + labelCols;
  const bodyBottom = dataTop + rowLines - 1;
  const bodyRight = Math.max(left, dataLeft + valueCols - 1);
  const body = makeCellRange(top, left, clampRow(bodyBottom), clampCol(bodyRight));
  return {
    location: {
      ref: rangeToString(body),
      firstHeaderRow: 1,
      firstDataRow: headerRows,
      firstDataCol: labelCols,
    },
    outputRef: rangeToString(makeCellRange(anchorRow, anchorCol, body.maxRow, Math.max(body.maxCol, pt.filters.length > 0 ? clampCol(anchorCol + 1) : body.maxCol))),
    anchorRow,
    top,
    left,
    valueCols,
    dataTop,
    dataLeft,
    empty: false,
  };
}

interface LayoutInput {
  fields: PivotCacheField[];
  names: string[];
  valueNames: string[];
  colFields: number[];
  rowTree: AxisLine[];
  colTree: AxisLine[];
  buckets: Array<Map<number, Accumulator[]>>;
  /** Date format of each field's items, by field index. */
  itemFormats: Array<string | undefined>;
}

function layoutReport(pt: PivotTable, g: ReportGeometry, input: LayoutInput): PivotComputation['cells'] {
  const cells: PivotComputation['cells'] = [];
  const put = (row: number, col: number, value: CellValue): void => {
    if (value !== null) cells.push({ row, col, value });
  };
  const putItem = (row: number, col: number, axis: number[], path: number[], level: number): void => {
    const value = itemLabel(input.fields, axis, path, level);
    const numFmt = typeof value === 'number' ? input.itemFormats[axis[level] ?? -1] : undefined;
    if (numFmt === undefined) put(row, col, value);
    else cells.push({ row, col, value, numFmt });
  };

  pt.filters.forEach((f, i) => {
    put(g.anchorRow + i, g.left, input.names[f.field] ?? '');
    const sel = f.selected;
    put(g.anchorRow + i, g.left + 1, sel === undefined ? ALL_ITEMS : sel === null ? BLANK_ITEM : sel);
  });
  if (g.empty) return cells;

  const { top, left, dataTop, dataLeft, valueCols } = g;
  const R = pt.rows.length;
  const D = pt.values.length;
  const C = input.colFields.length;
  const caption = D === 1 ? (input.valueNames[0] ?? '') : null;

  // Header block.
  if (C === 0) {
    if (R > 0) {
      if (pt.layout === 'compact') put(top, left, ROW_LABELS);
      else pt.rows.forEach((f, i) => put(top, left + i, input.names[f] ?? ''));
    }
    if (D === 1) put(top, dataLeft, caption);
    else input.valueNames.forEach((n, v) => put(top, dataLeft + v, n));
  } else {
    put(top, left, caption);
    if (pt.layout === 'compact') put(top, dataLeft, COLUMN_LABELS);
    else input.colFields.forEach((f, i) => put(top, dataLeft + i, f === PIVOT_VALUES_FIELD ? VALUES_CAPTION : (input.names[f] ?? '')));
    if (R > 0) {
      const labelRow = top + C;
      if (pt.layout === 'compact') put(labelRow, left, ROW_LABELS);
      else pt.rows.forEach((f, i) => put(labelRow, left + i, input.names[f] ?? ''));
    }
    // Column item labels: each line labels the levels it starts.
    let prev: number[] = [];
    input.colTree.forEach((line, j) => {
      const col = dataLeft + j;
      if (line.type === 'grand') {
        put(top + 1, col, D > 1 ? `Total ${input.valueNames[line.dataIndex] ?? ''}` : GRAND_TOTAL);
        prev = [];
        return;
      }
      if (line.type === 'default') {
        const label = itemLabel(input.fields, pt.columns, line.path, line.depth);
        put(top + 1 + line.depth, col, D > 1 ? `${label} ${input.valueNames[line.dataIndex] ?? ''}` : `${label} Total`);
        prev = line.path;
        return;
      }
      const path = fullPath(line);
      let common = 0;
      while (common < path.length && common < prev.length && path[common] === prev[common]) common++;
      for (let lvl = common; lvl < path.length; lvl++) {
        if (line.sigma && lvl === path.length - 1) put(top + 1 + lvl, col, input.valueNames[line.dataIndex] ?? '');
        else putItem(top + 1 + lvl, col, pt.columns, line.path, lvl);
      }
      prev = path;
    });
  }

  // Body: row labels then values.
  const colLines = input.colTree;
  let prevRow: number[] = [];
  input.rowTree.forEach((line, i) => {
    const row = dataTop + i;
    if (R > 0) {
      if (line.type === 'grand') {
        put(row, left, GRAND_TOTAL);
        prevRow = [];
      } else if (line.type === 'default') {
        const col = pt.layout === 'compact' ? left : left + line.depth;
        put(row, col, `${itemLabel(input.fields, pt.rows, line.path, line.depth)} Total`);
        prevRow = line.path;
      } else if (pt.layout === 'tabular') {
        let common = 0;
        while (common < line.path.length && common < prevRow.length && line.path[common] === prevRow[common]) common++;
        for (let lvl = common; lvl < line.path.length; lvl++) putItem(row, left + lvl, pt.rows, line.path, lvl);
        prevRow = line.path;
      } else {
        const col = pt.layout === 'compact' ? left : left + line.depth;
        putItem(row, col, pt.rows, line.path, line.depth);
        prevRow = line.path;
      }
    }
    const showValues = !(line.head && pt.subtotals !== 'top');
    if (!showValues || valueCols === 0) return;
    colLines.forEach((cl, j) => {
      const v = D > 1 ? cl.dataIndex : 0;
      const value = pt.values[v];
      if (value === undefined) return;
      const accs = input.buckets[line.node]?.get(cl.node);
      put(row, dataLeft + j, aggregateValue(accs?.[v], value.aggregate));
    });
  });

  return cells.filter((c) => c.row <= MAX_ROW && c.col <= MAX_COL);
}

function itemLabel(fields: PivotCacheField[], axis: number[], path: number[], level: number): CellValue {
  const item = fields[axis[level] ?? -1]?.items?.[path[level] ?? -1];
  if (item === undefined || item.t === 'm') return BLANK_ITEM;
  if (item.t === 'e') return { kind: 'error', code: item.v };
  return item.v;
}

function displayText(v: CellValue): string {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof Date) return v.toISOString();
  switch (v.kind) {
    case 'rich-text':
      return v.runs.map((r) => r.text).join('');
    case 'formula':
      return v.cachedValue === undefined ? '' : String(v.cachedValue);
    case 'error':
      return v.code;
    case 'duration':
      return String(v.ms);
  }
}

const clampRow = (r: number): number => Math.min(r, MAX_ROW);
const clampCol = (c: number): number => Math.min(c, MAX_COL);
