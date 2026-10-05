// PivotTables in the editor: Insert ▸ PivotTable, the field pane's edits, and
// the PivotTable Analyze / Design commands.
//
// The library keeps a pivot as a definition and writes its report cells on
// refresh. Every edit here clones the definition, works out where the new
// report lands, and then — as one undo step — swaps the clone in and rewrites
// the cells. The history snapshots `pivotTables` by value, so after an undo
// the pivot objects are new: callers address a pivot by its sheet and index,
// never by holding on to the object.

import type { PivotAggregate, PivotItemValue, PivotTable, PivotTableSummary } from '@office-kit/xlsx/worksheet';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import {
  addPivotTable,
  getPivotSourceFields,
  getPivotTableAt,
  getPivotTableOutputRef,
  listPassthroughPivotTables,
  nextPivotTableName,
  refreshPivotTable,
} from '@office-kit/xlsx/worksheet';
import { addWorksheet } from '@office-kit/xlsx/workbook';
import type { Workbook } from '@office-kit/xlsx/workbook';
import { OpenXmlError } from '@office-kit/xlsx/utils';
import type { Range } from './address.ts';
import { cellAddress, inRange, parseRangeAddress, quoteSheetName, rangeAddress, rangesIntersect, unionRange } from './address.ts';
import { forEachCellInRange, getCellAt, isBlank } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import { singleCell } from './selection.ts';
import { t, type MessageKey } from '../i18n/i18n.svelte.ts';

export type PivotArea = 'filters' | 'columns' | 'rows' | 'values';

/** The pivot under the active cell: an editable one, or a summary of one the library keeps verbatim. */
export type ActivePivot =
  | { readonly kind: 'model'; readonly ws: Worksheet; readonly index: number; readonly pt: PivotTable; readonly fields: string[] }
  | { readonly kind: 'readonly'; readonly summary: PivotTableSummary };

export function pivotAt(wb: Workbook, ws: Worksheet, row: number, col: number): ActivePivot | undefined {
  const pt = getPivotTableAt(ws, row, col);
  if (pt) {
    const index = ws.pivotTables?.indexOf(pt) ?? -1;
    let fields: string[];
    try {
      fields = getPivotSourceFields(wb, pt.source);
    } catch (err) {
      // The source header was edited into something Excel can't use as field
      // names (a blank cell); the pivot stays, its pane shows numbered fields.
      if (!(err instanceof OpenXmlError)) throw err;
      fields = [];
    }
    return { kind: 'model', ws, index, pt, fields };
  }
  if (!ws.relsExtras?.length) return undefined;
  for (const summary of listPassthroughPivotTables(wb, ws)) {
    const r = parseRangeAddress(summary.ref)?.range;
    if (r && inRange(r, row, col)) return { kind: 'readonly', summary };
  }
  return undefined;
}

const toRange = (ref: string): Range | undefined => parseRangeAddress(ref)?.range;

function rangeHasData(ws: Worksheet, range: Range, except: Range | undefined): boolean {
  let found = false;
  forEachCellInRange(ws, range, (c) => {
    if (!found && !isBlank(c) && !(except && inRange(except, c.row, c.col))) found = true;
  });
  return found;
}

/** Error key for a source the PivotTable can't use, or undefined when it is fine. */
function sourceProblem(wb: Workbook, source: { sheet: string; ref: string }): MessageKey | undefined {
  const r = toRange(source.ref);
  if (!r || r.r2 <= r.r1) return 'pvInvalidSource';
  try {
    getPivotSourceFields(wb, source);
  } catch (err) {
    if (!(err instanceof OpenXmlError)) throw err;
    return 'pvInvalidSource';
  }
  return undefined;
}

/**
 * Write `next` over the pivot at `index` of `ws` (or add it when `index` is
 * undefined) as one undo step. Asks before overwriting other data, as Excel
 * does. Returns an error message key, or 'cancelled' when the user declined
 * or protection refused, whenever nothing was done.
 */
function commit(
  ctl: EditorController,
  ws: Worksheet,
  index: number | undefined,
  next: PivotTable,
  label: string,
): MessageKey | 'cancelled' | undefined {
  const doc = ctl.doc;
  let outRef: string;
  try {
    outRef = getPivotTableOutputRef(doc.wb, next);
  } catch (err) {
    if (!(err instanceof OpenXmlError)) throw err;
    return 'pvInvalidSource';
  }
  const out = toRange(outRef);
  if (!out) return 'pvInvalidSource';
  const prev = index === undefined ? undefined : ws.pivotTables?.[index]?.renderedRef;
  const prevRange = prev === undefined ? undefined : toRange(prev);
  const source = toRange(next.source.ref);
  if (next.source.sheet === ws.title && source && rangesIntersect(source, out)) return 'pvOverlapSource';
  if (rangeHasData(ws, out, prevRange) && !window.confirm(t('pvReplaceData', { ref: `${quoteSheetName(ws.title)}!${rangeAddress(out, true)}` }))) {
    return 'cancelled';
  }
  // transact answers undefined when sheet protection refused the edit (it has
  // already rolled back and said so).
  const done = doc.transact(label, (tx) => {
    tx.sheet(ws, 'pivotTables');
    tx.cells(ws, prevRange ? unionRange(prevRange, out) : out);
    if (index === undefined) {
      addPivotTable(doc.wb, ws, next);
    } else {
      const list = ws.pivotTables ?? [];
      list[index] = next;
      refreshPivotTable(doc.wb, ws, next);
    }
    return true;
  });
  return done ? undefined : 'cancelled';
}

export interface CreatePivotOptions {
  source: { sheet: string; ref: string };
  /** Undefined places the report on a new sheet, at A3 as Excel does. */
  destination?: { sheet: string; cell: string };
}

/** Insert ▸ PivotTable. Returns an error message key, 'cancelled', or undefined once created. */
export function createPivotTable(ctl: EditorController, opts: CreatePivotOptions): MessageKey | 'cancelled' | undefined {
  const doc = ctl.doc;
  const problem = sourceProblem(doc.wb, opts.source);
  if (problem) return problem;
  const def: PivotTable = {
    name: nextPivotTableName(doc.wb),
    source: opts.source,
    anchor: opts.destination?.cell ?? 'A3',
    rows: [],
    columns: [],
    filters: [],
    values: [],
    layout: 'compact',
    subtotals: 'top',
    rowGrandTotals: true,
    columnGrandTotals: true,
  };
  let ws: Worksheet | undefined;
  if (opts.destination) {
    const ref = doc.wb.sheets.find((s) => s.kind === 'worksheet' && s.sheet.title === opts.destination?.sheet);
    if (ref?.kind !== 'worksheet') return 'invalidReference';
    ws = ref.sheet;
    const err = commit(ctl, ws, undefined, def, 'Insert PivotTable');
    if (err) return err;
  } else {
    // A new sheet goes in front of the active one, as Excel inserts it.
    const at = doc.activeSheetIndex;
    const done = doc.transact('Insert PivotTable', (tx) => {
      tx.structural = true;
      tx.workbook('sheets', 'activeSheetIndex', 'definedNames');
      ws = addWorksheet(doc.wb, uniqueSheetName(doc.wb), { index: at });
      addPivotTable(doc.wb, ws, def);
      return true;
    });
    if (!done) return 'cancelled';
    doc.activateSheet(at);
  }
  const anchor = parseRangeAddress(def.anchor)?.range;
  if (anchor) doc.setSelection(singleCell({ row: anchor.r1, col: anchor.c1 }));
  ctl.pivotFieldList = true;
  ctl.ribbonTab = 'pivotAnalyze';
  return undefined;
}

function uniqueSheetName(wb: Workbook): string {
  const names = new Set(wb.sheets.map((s) => s.sheet.title.toLowerCase()));
  let n = wb.sheets.length + 1;
  while (names.has(`sheet${n}`)) n++;
  return `Sheet${n}`;
}

/** Apply `edit` to a copy of the pivot and commit it. Shows an alert on failure. */
export function editPivot(ctl: EditorController, ws: Worksheet, index: number, label: string, edit: (pt: PivotTable) => void): void {
  const current = ws.pivotTables?.[index];
  if (!current) return;
  const next = structuredClone(current);
  edit(next);
  keepTableInPlace(current, next);
  const err = commit(ctl, ws, index, next, label);
  if (err && err !== 'cancelled') ctl.dialog = { kind: 'alert', props: { message: err } };
}

/**
 * Report filters stack above the table with a spacer row. Excel keeps the
 * table where it is when filters come and go and moves the anchor instead,
 * pushing the table down only when there is no room above it.
 */
function keepTableInPlace(before: PivotTable, after: PivotTable): void {
  const anchor = parseRangeAddress(before.anchor)?.range;
  if (!anchor) return;
  const above = (pt: PivotTable): number => (pt.filters.length > 0 ? pt.filters.length + 1 : 0);
  const tableTop = anchor.r1 + above(before);
  after.anchor = cellAddress(Math.max(1, tableTop - above(after)), anchor.c1);
}

// ---- field placement ------------------------------------------------------

/** Which area holds `field`, and where in it. Values can hold a field more than once; this finds the first. */
export function placementOf(pt: PivotTable, field: number): { area: PivotArea; position: number } | undefined {
  const rows = pt.rows.indexOf(field);
  if (rows >= 0) return { area: 'rows', position: rows };
  const cols = pt.columns.indexOf(field);
  if (cols >= 0) return { area: 'columns', position: cols };
  const filters = pt.filters.findIndex((f) => f.field === field);
  if (filters >= 0) return { area: 'filters', position: filters };
  const values = pt.values.findIndex((v) => v.field === field);
  if (values >= 0) return { area: 'values', position: values };
  return undefined;
}

/** Take the entry at `position` of `area` out of the definition. */
export function removeFromArea(pt: PivotTable, area: PivotArea, position: number): void {
  if (area === 'rows') pt.rows.splice(position, 1);
  else if (area === 'columns') pt.columns.splice(position, 1);
  else if (area === 'filters') pt.filters.splice(position, 1);
  else pt.values.splice(position, 1);
}

/**
 * Put `field` into `area` at `position` (end when undefined). A field sits on
 * one of rows / columns / filters at a time, so placing it on one takes it off
 * the others, as dragging does in Excel. Values may repeat a field.
 */
export function placeField(pt: PivotTable, field: number, area: PivotArea, position?: number, aggregate: PivotAggregate = 'sum'): void {
  const insert = <T>(list: T[], item: T): void => {
    list.splice(position === undefined ? list.length : Math.min(position, list.length), 0, item);
  };
  if (area === 'values') {
    insert(pt.values, { field, aggregate });
    return;
  }
  pt.rows = pt.rows.filter((f) => f !== field);
  pt.columns = pt.columns.filter((f) => f !== field);
  pt.filters = pt.filters.filter((f) => f.field !== field);
  if (area === 'rows') insert(pt.rows, field);
  else if (area === 'columns') insert(pt.columns, field);
  else insert(pt.filters, { field });
}

/** Ticking a field in the list: numbers go to Values (Sum), anything else to Rows — Excel's rule. */
export function defaultAreaFor(wb: Workbook, pt: PivotTable, field: number): PivotArea {
  const ws = wb.sheets.find((s) => s.kind === 'worksheet' && s.sheet.title === pt.source.sheet);
  const r = toRange(pt.source.ref);
  if (ws?.kind !== 'worksheet' || !r) return 'rows';
  const col = r.c1 + field;
  let numbers = 0;
  for (let row = r.r1 + 1; row <= r.r2; row++) {
    const v = getCellAt(ws.sheet, row, col)?.value;
    if (v === null || v === undefined) continue;
    const plain = typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.cachedValue : v;
    if (typeof plain === 'number' || plain instanceof Date) numbers++;
    else if (plain !== undefined && plain !== '') return 'rows';
  }
  return numbers > 0 ? 'values' : 'rows';
}

export function itemLabel(v: PivotItemValue): string {
  if (v === null) return t('pvBlank');
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

/** `Sheet1!$A$1:$D$20` for the Create / Change Data Source dialogs. */
export function sourceText(source: { sheet: string; ref: string }): string {
  const r = toRange(source.ref);
  return r ? `${quoteSheetName(source.sheet)}!${rangeAddress(r, true)}` : source.ref;
}

/** Parse what the user typed for a source range; an unqualified range is on `defaultSheet`. */
export function parseSource(wb: Workbook, text: string, defaultSheet: string): { sheet: string; ref: string } | undefined {
  const parsed = parseRangeAddress(text.trim().replace(/^=/, ''));
  if (!parsed) return undefined;
  const sheet = parsed.sheet ?? defaultSheet;
  const ws = wb.sheets.find((s) => s.kind === 'worksheet' && s.sheet.title.toLowerCase() === sheet.toLowerCase());
  if (ws?.kind !== 'worksheet') return undefined;
  let r = parsed.range;
  // Whole columns shrink to the used rows, as Excel does.
  if (r.r1 === 1 && r.r2 === 1_048_576) {
    let last = 1;
    for (const row of ws.sheet.rows.keys()) if (row > last) last = row;
    r = { ...r, r2: last };
  }
  return { sheet: ws.sheet.title, ref: rangeAddress(r) };
}

/** Parse the Location box of the Create dialog: a cell, optionally sheet-qualified. */
export function parseDestination(wb: Workbook, text: string, defaultSheet: string): { sheet: string; cell: string } | undefined {
  const parsed = parseRangeAddress(text.trim().replace(/^=/, ''));
  if (!parsed) return undefined;
  const sheet = parsed.sheet ?? defaultSheet;
  const ws = wb.sheets.find((s) => s.kind === 'worksheet' && s.sheet.title.toLowerCase() === sheet.toLowerCase());
  if (ws?.kind !== 'worksheet') return undefined;
  return { sheet: ws.sheet.title, cell: cellAddress(parsed.range.r1, parsed.range.c1) };
}
