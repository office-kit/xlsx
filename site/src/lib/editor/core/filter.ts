// AutoFilter: the sheet's own filter and each table's header filter. Value
// lists persist in the model (the only filter kind the file model carries);
// text/number conditions are kept for the session alongside them.

import type { Cell } from '@office-kit/xlsx/cell';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import type { AutoFilter, CustomFilterCondition, FilterColumn, Worksheet } from '@office-kit/xlsx/worksheet';
import { MAX_ROW, parseRangeAddress, rangeAddress, type Range } from './address.ts';
import { getCellAt, isBlank } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import { dataRange } from './data.ts';
import { setFilterButton, tableAt } from './tables.ts';

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

/**
 * `range` without the rows a filter hides, as runs of visible rows: in a
 * filtered list Excel clears, formats and pastes only the rows on show. A
 * whole-column range is left whole, since its format lives on the column.
 */
export function visibleParts(ws: Worksheet, range: Range): Range[] {
  if (range.r1 === 1 && range.r2 === MAX_ROW) return [range];
  const owners = filterOwners(ws).filter((o) => o.autoFilter.filterColumns.length > 0);
  if (owners.length === 0) return [range];
  const hidden: number[] = [];
  for (const [row, dim] of ws.rowDimensions) {
    if (dim.hidden !== true || row < range.r1 || row > range.r2) continue;
    if (owners.some((o) => row > o.range.r1 && row <= o.range.r2)) hidden.push(row);
  }
  if (hidden.length === 0) return [range];
  hidden.sort((a, b) => a - b);
  const parts: Range[] = [];
  let from = range.r1;
  for (const row of hidden) {
    if (row > from) parts.push({ ...range, r1: from, r2: row - 1 });
    from = row + 1;
  }
  if (from <= range.r2) parts.push({ ...range, r1: from });
  return parts;
}

/** The filter whose header row holds (row, col), if any. */
export function filterOwnerAt(ws: Worksheet, row: number, col: number): FilterOwner | undefined {
  return filterOwners(ws).find((o) => o.range.r1 === row && col >= o.range.c1 && col <= o.range.c2);
}

export function toggleAutoFilter(ctl: EditorController): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  // Inside a table, Filter shows or hides the table's own buttons; a sheet filter would overlap it.
  const { row, col } = doc.selection.active;
  const table = tableAt(ws, row, col);
  if (table) {
    setFilterButton(ctl, table.def, table.def.autoFilter === undefined);
    return;
  }
  if (ws.autoFilter) {
    const range = autoFilterRange(ws);
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

/** Criteria per sheet column, compiled from the filter's saved columns (compiled afresh: Top 10 and averages read the data). */
export function activeCriteria(ctl: EditorController, owner: FilterOwner): ReadonlyMap<number, FilterCriterion> {
  return new Map(owner.autoFilter.filterColumns.map((fc) => [owner.range.c1 + fc.colId, compileColumn(ctl, owner, fc)]));
}

function numberOf(cell: Cell | undefined): number | undefined {
  const v = cell?.value;
  if (typeof v === 'number') return v;
  if (v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' && typeof v.cachedValue === 'number') return v.cachedValue;
  return undefined;
}

/** Excel's `*` / `?` wildcards (`~` escapes) as a whole-text, case-insensitive pattern. */
function wildcard(pattern: string): RegExp {
  let source = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern.charAt(i);
    if (ch === '~' && i + 1 < pattern.length) source += pattern.charAt(++i).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (ch === '*') source += '.*';
    else if (ch === '?') source += '.';
    else source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${source}$`, 'is');
}

function conditionTest(c: CustomFilterCondition): (cell: Cell | undefined, text: string) => boolean {
  const op = c.operator ?? 'equal';
  const target = c.val.trim() === '' ? Number.NaN : Number(c.val);
  if (op === 'equal' || op === 'notEqual') {
    const re = wildcard(c.val);
    const match = (cell: Cell | undefined, text: string) => {
      const n = numberOf(cell);
      return Number.isFinite(target) && n !== undefined ? n === target : re.test(text);
    };
    return op === 'equal' ? match : (cell, text) => !match(cell, text);
  }
  return (cell, text) => {
    const n = numberOf(cell);
    const cmp = Number.isFinite(target) ? (n === undefined ? undefined : n - target) : text.localeCompare(c.val, undefined, { sensitivity: 'base' });
    if (cmp === undefined) return false;
    switch (op) {
      case 'lessThan':
        return cmp < 0;
      case 'lessThanOrEqual':
        return cmp <= 0;
      case 'greaterThan':
        return cmp > 0;
      case 'greaterThanOrEqual':
        return cmp >= 0;
    }
  };
}

function columnNumbers(ctl: EditorController, owner: FilterOwner, col: number): number[] {
  const out: number[] = [];
  for (let r = owner.range.r1 + 1; r <= owner.range.r2; r++) {
    const n = numberOf(getCellAt(ctl.doc.ws, r, col));
    if (n !== undefined) out.push(n);
  }
  return out;
}

/** What a saved filter column keeps visible. Criteria the editor cannot evaluate (colour, dates) keep every row. */
function compileColumn(ctl: EditorController, owner: FilterOwner, fc: FilterColumn): FilterCriterion {
  const col = owner.range.c1 + fc.colId;
  switch (fc.kind) {
    case 'filters': {
      const values = new Set(fc.values);
      if (fc.blank) values.add('');
      return { values };
    }
    case 'custom': {
      const tests = fc.conditions.map(conditionTest);
      return { test: (cell, text) => (fc.and ? tests.every((t) => t(cell, text)) : tests.some((t) => t(cell, text))) };
    }
    case 'top10': {
      const nums = columnNumbers(ctl, owner, col).sort((a, b) => (fc.top === false ? a - b : b - a));
      const k = Math.max(1, fc.percent ? Math.floor((nums.length * fc.val) / 100) : Math.floor(fc.val));
      const cut = nums[Math.min(k, nums.length) - 1];
      if (cut === undefined) return {};
      return { test: (cell) => {
        const n = numberOf(cell);
        return n !== undefined && (fc.top === false ? n <= cut : n >= cut);
      } };
    }
    case 'dynamic': {
      if (fc.type !== 'aboveAverage' && fc.type !== 'belowAverage') return {};
      const nums = columnNumbers(ctl, owner, col);
      const avg = nums.reduce((sum, n) => sum + n, 0) / Math.max(1, nums.length);
      return { test: (cell) => {
        const n = numberOf(cell);
        return n !== undefined && (fc.type === 'aboveAverage' ? n > avg : n < avg);
      } };
    }
    case 'raw':
      return {};
  }
}

/** A value-list filter column keeping `values` ('' = blanks) of sheet column `col`. */
export function valuesColumn(owner: FilterOwner, col: number, values: ReadonlySet<string>): FilterColumn {
  const listed = [...values].filter((v) => v !== '');
  return { kind: 'filters', colId: col - owner.range.c1, values: listed, ...(values.has('') ? { blank: true } : {}) };
}

/**
 * Apply (or clear, with `undefined`) a column's filter and re-evaluate row
 * visibility across all of the owner's filtered columns.
 */
export function setColumnFilter(ctl: EditorController, owner: FilterOwner, col: number, column: FilterColumn | undefined): void {
  const ws = ctl.doc.ws;
  ctl.doc.transact('Filter', (tx) => {
    tx.sheet(ws, 'autoFilter', 'tables', 'rowDimensions');
    writeColumnFilter(ctl, owner, col, column);
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
    writeColumnFilter(ctl, owner, col, valuesColumn(owner, col, values));
  });
}

function writeColumnFilter(ctl: EditorController, owner: FilterOwner, col: number, column: FilterColumn | undefined): void {
  const af = owner.autoFilter;
  const colId = col - owner.range.c1;
  af.filterColumns = [...af.filterColumns.filter((fc) => fc.colId !== colId), ...(column ? [{ ...column, colId }] : [])];
  applyCriteria(ctl, owner.range, activeCriteria(ctl, owner));
}

/** Data ▸ Reapply: re-run every filter after the data changed. */
export function reapplyFilters(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const owners = filterOwners(ws);
  if (owners.length === 0) return;
  ctl.doc.transact('Reapply', (tx) => {
    tx.sheet(ws, 'rowDimensions');
    for (const owner of owners) applyCriteria(ctl, owner.range, activeCriteria(ctl, owner));
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
      owner.autoFilter.filterColumns = [];
      for (let r = owner.range.r1 + 1; r <= owner.range.r2; r++) unhideRow(ws, r);
    }
  });
}
