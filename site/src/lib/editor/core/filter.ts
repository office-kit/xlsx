// AutoFilter: the sheet's own filter and each table's header filter. Value
// lists persist in the model (the only filter kind the file model carries);
// text/number conditions are kept for the session alongside them.

import type { Cell } from '@office-kit/xlsx/cell';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import type { AutoFilter, Worksheet } from '@office-kit/xlsx/worksheet';
import { parseRangeAddress, rangeAddress, type Range } from './address.ts';
import { getCellAt, isBlank } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import { dataRange } from './data.ts';

/** A filter and the range it covers (header row first). */
export interface FilterOwner {
  readonly range: Range;
  readonly autoFilter: AutoFilter;
}

function rangeOfRef(ref: string): Range | undefined {
  return parseRangeAddress(ref)?.range;
}

export function autoFilterRange(ws: Worksheet): Range | undefined {
  return ws.autoFilter ? rangeOfRef(ws.autoFilter.ref) : undefined;
}

/** Every active filter on the sheet: the sheet AutoFilter and table filters. */
export function filterOwners(ws: Worksheet): FilterOwner[] {
  const out: FilterOwner[] = [];
  const sheetRange = autoFilterRange(ws);
  if (ws.autoFilter && sheetRange) out.push({ range: sheetRange, autoFilter: ws.autoFilter });
  for (const table of ws.tables) {
    if (!table.autoFilter || table.headerRowCount === 0) continue;
    const range = rangeOfRef(table.autoFilter.ref) ?? rangeOfRef(table.ref);
    if (range) out.push({ range, autoFilter: table.autoFilter });
  }
  return out;
}

/** Hidden by a filter with criteria, not by hand: copy and fill skip these rows. */
export function isRowFiltered(ws: Worksheet, row: number): boolean {
  if (ws.rowDimensions.get(row)?.hidden !== true) return false;
  return filterOwners(ws).some((o) => o.autoFilter.filterColumns.length > 0 && row > o.range.r1 && row <= o.range.r2);
}

/** The filter whose header row holds (row, col), if any. */
export function filterOwnerAt(ws: Worksheet, row: number, col: number): FilterOwner | undefined {
  return filterOwners(ws).find((o) => o.range.r1 === row && col >= o.range.c1 && col <= o.range.c2);
}

export function toggleAutoFilter(ctl: EditorController): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  if (ws.autoFilter) {
    const range = autoFilterRange(ws);
    criteria.delete(ws.autoFilter);
    doc.transact('Filter', (tx) => {
      tx.sheet(ws, 'autoFilter', 'rowDimensions');
      if (range) for (let r = range.r1 + 1; r <= range.r2; r++) unhideRow(ws, r);
      delete ws.autoFilter;
    });
    return;
  }
  const range = dataRange(ctl);
  doc.transact('Filter', (tx) => {
    tx.sheet(ws, 'autoFilter');
    ws.autoFilter = { ref: rangeAddress(range), filterColumns: [] };
  });
}

function unhideRow(ws: Worksheet, r: number): void {
  const dim = ws.rowDimensions.get(r);
  if (!dim?.hidden) return;
  const { hidden: _hidden, ...rest } = dim;
  if (Object.keys(rest).length === 0) ws.rowDimensions.delete(r);
  else ws.rowDimensions.set(r, rest);
}

/** Text a filter list shows and matches for a cell: its displayed value. */
export function filterKey(ctl: EditorController, cell: Cell | undefined): string {
  return cell && !isBlank(cell) ? getCellDisplayText(ctl.doc.wb, cell) : '';
}

/** Distinct displayed values of a filter column (blank as ''), in Excel's list order. */
export function filterValues(ctl: EditorController, owner: FilterOwner, col: number): string[] {
  const ws = ctl.doc.ws;
  const set = new Set<string>();
  for (let r = owner.range.r1 + 1; r <= owner.range.r2; r++) set.add(filterKey(ctl, getCellAt(ws, r, col)));
  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
  return [...set].sort((a, b) => (a === '' ? 1 : b === '' ? -1 : collator.compare(a, b)));
}

export interface FilterCriterion {
  /** Allowed displayed values ('' = blanks); undefined = no value filter. */
  readonly values?: ReadonlySet<string>;
  /** Custom condition (text/number filters), evaluated on the cell. */
  readonly test?: (cell: Cell | undefined, text: string) => boolean;
}

const criteria = new WeakMap<AutoFilter, Map<number, FilterCriterion>>();

/** Criteria per sheet column; value lists loaded from a file are picked up from the model. */
export function activeCriteria(owner: FilterOwner): ReadonlyMap<number, FilterCriterion> {
  let map = criteria.get(owner.autoFilter);
  if (!map) {
    map = new Map();
    for (const fc of owner.autoFilter.filterColumns) {
      const values = new Set(fc.values);
      if (fc.blank) values.add('');
      map.set(owner.range.c1 + fc.colId, { values });
    }
    criteria.set(owner.autoFilter, map);
  }
  return map;
}

/**
 * Apply (or clear, with `undefined`) a column's filter and re-evaluate row
 * visibility across all of the owner's filtered columns.
 */
export function setColumnFilter(ctl: EditorController, owner: FilterOwner, col: number, criterion: FilterCriterion | undefined): void {
  const ws = ctl.doc.ws;
  ctl.doc.transact('Filter', (tx) => {
    tx.sheet(ws, 'autoFilter', 'tables', 'rowDimensions');
    writeColumnFilter(ctl, owner, col, criterion);
  });
}

/**
 * Cell menu ▸ Filter ▸ Filter by Selected Cell's Value: keep the rows whose
 * cell in the active column shows the active cell's value. A cell in no
 * filter gets a sheet AutoFilter over its data region first, in the same
 * undo step; a cell outside an existing sheet AutoFilter is left alone.
 */
export function filterBySelectedValue(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const { row, col } = ctl.doc.selection.active;
  const existing = filterOwners(ws).find((o) => row > o.range.r1 && row <= o.range.r2 && col >= o.range.c1 && col <= o.range.c2);
  if (!existing && ws.autoFilter) return;
  const values = new Set([filterKey(ctl, getCellAt(ws, row, col))]);
  ctl.doc.transact('Filter', (tx) => {
    tx.sheet(ws, 'autoFilter', 'tables', 'rowDimensions');
    let owner = existing;
    if (!owner) {
      const range = dataRange(ctl);
      ws.autoFilter = { ref: rangeAddress(range), filterColumns: [] };
      owner = { range, autoFilter: ws.autoFilter };
    }
    writeColumnFilter(ctl, owner, col, { values });
  });
}

function writeColumnFilter(ctl: EditorController, owner: FilterOwner, col: number, criterion: FilterCriterion | undefined): void {
  const map = new Map(activeCriteria(owner));
  if (criterion) map.set(col, criterion);
  else map.delete(col);
  criteria.set(owner.autoFilter, map);
  const af = owner.autoFilter;
  const colId = col - owner.range.c1;
  af.filterColumns = af.filterColumns.filter((fc) => fc.colId !== colId);
  if (criterion?.values) {
    const values = [...criterion.values].filter((v) => v !== '');
    af.filterColumns.push({ kind: 'filters', colId, values, ...(criterion.values.has('') ? { blank: true } : {}) });
  }
  applyCriteria(ctl, owner.range, map);
}

/** Data ▸ Reapply: re-run every filter after the data changed. */
export function reapplyFilters(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const owners = filterOwners(ws);
  if (owners.length === 0) return;
  ctl.doc.transact('Reapply', (tx) => {
    tx.sheet(ws, 'rowDimensions');
    for (const owner of owners) applyCriteria(ctl, owner.range, activeCriteria(owner));
  });
}

function applyCriteria(ctl: EditorController, range: Range, map: ReadonlyMap<number, FilterCriterion>): void {
  const ws = ctl.doc.ws;
  for (let r = range.r1 + 1; r <= range.r2; r++) {
    let visible = true;
    for (const [col, crit] of map) {
      const cell = getCellAt(ws, r, col);
      const text = filterKey(ctl, cell);
      if ((crit.values && !crit.values.has(text)) || (crit.test && !crit.test(cell, text))) {
        visible = false;
        break;
      }
    }
    if (visible) unhideRow(ws, r);
    else ws.rowDimensions.set(r, { ...ws.rowDimensions.get(r), hidden: true });
  }
}

/** Data ▸ Clear: drop every condition of every filter on the sheet, keeping the buttons. */
export function clearAllFilters(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const owners = filterOwners(ws);
  if (owners.length === 0) return;
  ctl.doc.transact('Clear Filter', (tx) => {
    tx.sheet(ws, 'autoFilter', 'tables', 'rowDimensions');
    for (const owner of owners) {
      criteria.delete(owner.autoFilter);
      owner.autoFilter.filterColumns = [];
      for (let r = owner.range.r1 + 1; r <= owner.range.r2; r++) unhideRow(ws, r);
    }
  });
}
