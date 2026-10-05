// Inserting and deleting whole rows/columns, and shifting cells within a band
// (Insert/Delete Cells ▸ Shift down/right/up/left).
//
// The library has no structural edit, so this moves everything that carries a
// position: the sparse cell maps, row/column dimensions, merges, hyperlinks,
// notes, threaded comments, data validations, conditional formats, tables, the AutoFilter,
// drawing anchors, defined names and every formula in the workbook (via the
// formula engine's reference adjuster, which turns references to deleted
// cells into #REF! the way Excel does).

import type { Cell } from '@office-kit/xlsx/cell';
import type { Workbook } from '@office-kit/xlsx/workbook';
import type { PivotTable, TableColumn, TableDefinition, Worksheet } from '@office-kit/xlsx/worksheet';
import { adjustFormulaForStructure } from '../calc/index.ts';
import type { Range } from './address.ts';
import { colLetter, MAX_COL, MAX_ROW, parseRangeAddress } from './address.ts';
import type { Transaction } from './history.ts';
import { mapTableColumnsInFormula } from './tables.ts';

export type Axis = 'row' | 'col';

interface Edit {
  readonly axis: Axis;
  /** First index affected. */
  readonly at: number;
  /** > 0 inserts before `at`; < 0 deletes |count| indices from `at`. */
  readonly count: number;
  /** Restrict the shift to this band of the other axis (Insert/Delete Cells); undefined = whole rows/columns. */
  readonly band?: { readonly from: number; readonly to: number };
}

/** New index of `i` after the edit, or undefined when `i` was deleted. */
function mapIndex(i: number, e: Edit): number | undefined {
  if (i < e.at) return i;
  if (e.count > 0) return i + e.count;
  const deletedEnd = e.at - e.count - 1;
  if (i <= deletedEnd) return undefined;
  return i + e.count;
}

/** Map a 1-D span; undefined when the whole span was deleted. */
function mapSpan(lo: number, hi: number, e: Edit, limit: number): [number, number] | undefined {
  if (e.count > 0) {
    const nlo = lo >= e.at ? lo + e.count : lo;
    const nhi = hi >= e.at ? hi + e.count : hi;
    if (nlo > limit) return undefined;
    return [nlo, Math.min(nhi, limit)];
  }
  const delLo = e.at;
  const delHi = e.at - e.count - 1;
  if (lo >= delLo && hi <= delHi) return undefined;
  const nlo = lo < delLo ? lo : lo <= delHi ? delLo : lo + e.count;
  const nhi = hi < delLo ? hi : hi <= delHi ? delLo - 1 : hi + e.count;
  return nhi < nlo ? undefined : [nlo, nhi];
}

function inBand(e: Edit, other: number): boolean {
  return !e.band || (other >= e.band.from && other <= e.band.to);
}

export function mapRange(r: Range, e: Edit): Range | undefined {
  const limit = e.axis === 'row' ? MAX_ROW : MAX_COL;
  if (e.band) {
    // A band shift moves only ranges lying entirely within the band.
    const [olo, ohi] = e.axis === 'row' ? [r.c1, r.c2] : [r.r1, r.r2];
    if (olo < e.band.from || ohi > e.band.to) return r;
  }
  const [lo, hi] = e.axis === 'row' ? [r.r1, r.r2] : [r.c1, r.c2];
  // Whole-column ranges are unaffected by row edits (and vice versa).
  if (e.axis === 'row' && r.r1 === 1 && r.r2 === MAX_ROW) return r;
  if (e.axis === 'col' && r.c1 === 1 && r.c2 === MAX_COL) return r;
  const span = mapSpan(lo, hi, e, limit);
  if (!span) return undefined;
  return e.axis === 'row' ? { ...r, r1: span[0], r2: span[1] } : { ...r, c1: span[0], c2: span[1] };
}

// ---- A1 text helpers for the string-typed refs --------------------------------

function parseRef(ref: string): Range | undefined {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(ref.trim());
  if (!m?.[1] || !m[2]) return undefined;
  const col = (s: string) => {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
    return n;
  };
  const c1 = col(m[1]);
  const r1 = Number(m[2]);
  const c2 = m[3] ? col(m[3]) : c1;
  const r2 = m[4] ? Number(m[4]) : r1;
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

function formatRef(r: Range): string {
  const a = `${colLetter(r.c1)}${r.r1}`;
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:${colLetter(r.c2)}${r.r2}`;
}

function mapRefText(ref: string, e: Edit): string | undefined {
  const r = parseRef(ref);
  if (!r) return ref;
  const next = mapRange(r, e);
  return next ? formatRef(next) : undefined;
}

// ---- the edit -----------------------------------------------------------------

function shiftCells(ws: Worksheet, e: Edit): void {
  if (e.axis === 'row') {
    const next = new Map<number, Map<number, Cell>>();
    const moved: Cell[] = [];
    for (const [r, rowMap] of ws.rows) {
      if (!e.band) {
        const nr = mapIndex(r, e);
        if (nr === undefined || nr > MAX_ROW) continue;
        for (const cell of rowMap.values()) cell.row = nr;
        next.set(nr, rowMap);
        continue;
      }
      // Band shift: only cells in the band move; others stay.
      for (const [c, cell] of rowMap) {
        if (!inBand(e, c)) {
          moved.push(cell);
          continue;
        }
        const nr = mapIndex(r, e);
        if (nr === undefined || nr > MAX_ROW) continue;
        cell.row = nr;
        moved.push(cell);
      }
    }
    if (e.band) {
      for (const cell of moved) {
        let m = next.get(cell.row);
        if (!m) {
          m = new Map();
          next.set(cell.row, m);
        }
        m.set(cell.col, cell);
      }
    }
    ws.rows = next;
    return;
  }
  for (const [r, rowMap] of ws.rows) {
    if (e.band && !inBand(e, r)) continue;
    const next = new Map<number, Cell>();
    for (const [c, cell] of rowMap) {
      const nc = mapIndex(c, e);
      if (nc === undefined || nc > MAX_COL) continue;
      cell.col = nc;
      next.set(nc, cell);
    }
    if (next.size === 0) ws.rows.delete(r);
    else ws.rows.set(r, next);
  }
}

function shiftDimensions(ws: Worksheet, e: Edit): void {
  if (e.band) return;
  if (e.axis === 'row') {
    const next = new Map<number, (typeof ws.rowDimensions extends Map<number, infer D> ? D : never)>();
    for (const [r, dim] of ws.rowDimensions) {
      const nr = mapIndex(r, e);
      if (nr !== undefined && nr <= MAX_ROW) next.set(nr, dim);
    }
    // Inserted rows take the format of the row above (Excel's default "Format Same As Above").
    if (e.count > 0) {
      const above = ws.rowDimensions.get(e.at - 1);
      if (above) for (let i = 0; i < e.count; i++) next.set(e.at + i, { ...above, hidden: false });
    }
    ws.rowDimensions = next;
    return;
  }
  const entries = [...ws.columnDimensions.values()];
  const next = new Map<number, (typeof entries)[number]>();
  for (const dim of entries) {
    const span = mapSpan(dim.min, dim.max, e, MAX_COL);
    if (!span) continue;
    next.set(span[0], { ...dim, min: span[0], max: span[1] });
  }
  if (e.count > 0) {
    const left = entries.find((d) => e.at - 1 >= d.min && e.at - 1 <= d.max);
    if (left && !(e.at >= left.min && e.at <= left.max)) {
      next.set(e.at, { ...left, min: e.at, max: e.at + e.count - 1, hidden: false });
    }
  }
  ws.columnDimensions = new Map([...next].sort((a, b) => a[0] - b[0]));
}

function shiftSheetObjects(ws: Worksheet, e: Edit, dropped: DroppedColumns[]): void {
  ws.mergedCells = ws.mergedCells.flatMap((m) => {
    const r = mapRange({ r1: m.minRow, c1: m.minCol, r2: m.maxRow, c2: m.maxCol }, e);
    if (!r || (r.r1 === r.r2 && r.c1 === r.c2)) return [];
    return [{ minRow: r.r1, minCol: r.c1, maxRow: r.r2, maxCol: r.c2 }];
  });
  ws.hyperlinks = ws.hyperlinks.flatMap((h) => {
    const ref = mapRefText(h.ref, e);
    return ref ? [{ ...h, ref }] : [];
  });
  ws.legacyComments = ws.legacyComments.flatMap((c) => {
    const ref = mapRefText(c.ref, e);
    return ref ? [{ ...c, ref }] : [];
  });
  // Replies share their root's ref, so a deleted cell drops the whole thread.
  ws.threadedComments = (ws.threadedComments ?? []).flatMap((c) => {
    const ref = mapRefText(c.ref, e);
    return ref ? [{ ...c, ref }] : [];
  });
  const mapMulti = (ranges: ReadonlyArray<{ minRow: number; minCol: number; maxRow: number; maxCol: number }>) =>
    ranges.flatMap((m) => {
      const r = mapRange({ r1: m.minRow, c1: m.minCol, r2: m.maxRow, c2: m.maxCol }, e);
      return r ? [{ minRow: r.r1, minCol: r.c1, maxRow: r.r2, maxCol: r.c2 }] : [];
    });
  ws.dataValidations = ws.dataValidations.flatMap((dv) => {
    const ranges = mapMulti(dv.sqref.ranges);
    return ranges.length ? [{ ...dv, sqref: { ranges } }] : [];
  });
  ws.conditionalFormatting = ws.conditionalFormatting.flatMap((cf) => {
    const ranges = mapMulti(cf.sqref.ranges);
    return ranges.length ? [{ ...cf, sqref: { ranges } }] : [];
  });
  ws.tables = ws.tables.flatMap((t) => {
    const ref = mapRefText(t.ref, e);
    if (!ref) {
      dropped.push({ table: t.displayName, columns: t.columns.map((c) => c.name) });
      return [];
    }
    // Mapped on its own: with a total row the filter stops one row short of the table.
    const autoFilter = t.autoFilter ? { ...t.autoFilter, ref: mapRefText(t.autoFilter.ref, e) ?? ref } : undefined;
    const columns = mapTableColumns(t, e, dropped);
    return [{ ...t, ref, columns, ...(autoFilter ? { autoFilter } : {}) }];
  });
  if (ws.autoFilter) {
    const ref = mapRefText(ws.autoFilter.ref, e);
    if (ref) ws.autoFilter = { ...ws.autoFilter, ref };
    else delete ws.autoFilter;
  }
  if (ws.drawing && !e.band) {
    for (const item of ws.drawing.items) {
      const a = item.anchor;
      if (a.kind === 'absolute') continue;
      const key = e.axis === 'row' ? 'row' : 'col';
      // Anchors are 0-based.
      const from = mapIndex(a.from[key] + 1, e);
      a.from[key] = (from ?? e.at) - 1;
      if (a.kind === 'twoCell') {
        const to = mapIndex(a.to[key] + 1, e);
        a.to[key] = (to ?? e.at) - 1;
      }
    }
  }
}

/**
 * A table's column list after whole columns were inserted or deleted through
 * it: deleted columns go (their names are recorded so references to them can
 * become #REF!), inserted ones get Excel's unused ColumnN names.
 */
function mapTableColumns(t: TableDefinition, e: Edit, dropped: DroppedColumns[]): TableColumn[] {
  const range = parseRangeAddress(t.ref)?.range;
  if (!range || e.axis !== 'col' || (e.band && (e.band.from > range.r1 || e.band.to < range.r2))) return t.columns;
  if (e.count < 0) {
    const last = e.at - e.count - 1;
    const kept = t.columns.filter((_, i) => range.c1 + i < e.at || range.c1 + i > last);
    if (kept.length !== t.columns.length && kept.length > 0) {
      dropped.push({ table: t.displayName, columns: t.columns.filter((c) => !kept.includes(c)).map((c) => c.name) });
    }
    return kept.length > 0 ? kept : t.columns;
  }
  if (e.at <= range.c1 || e.at > range.c2) return t.columns;
  const taken = new Set(t.columns.map((c) => c.name.toLowerCase()));
  let nextId = Math.max(0, ...t.columns.map((c) => c.id)) + 1;
  let n = 1;
  const added: TableColumn[] = [];
  for (let k = 0; k < e.count; k++) {
    while (taken.has(`column${n}`)) n++;
    taken.add(`column${n}`);
    added.push({ id: nextId++, name: `Column${n}` });
  }
  const at = e.at - range.c1;
  return [...t.columns.slice(0, at), ...added, ...t.columns.slice(at)];
}

interface DroppedColumns {
  readonly table: string;
  readonly columns: readonly string[];
}

/** References to deleted table columns (or a deleted table) become #REF!, as in Excel. */
function dropTableColumnRefs(wb: Workbook, dropped: readonly DroppedColumns[]): void {
  if (dropped.length === 0) return;
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    for (const rowMap of ref.sheet.rows.values()) {
      for (const cell of rowMap.values()) {
        const v = cell.value;
        if (v === null || typeof v !== 'object' || v instanceof Date || v.kind !== 'formula' || !v.formula.includes('[')) continue;
        let next = v.formula;
        for (const d of dropped) {
          const gone = new Set(d.columns.map((c) => c.toLowerCase()));
          next = mapTableColumnsInFormula(next, d.table, false, (name) => (gone.has(name.toLowerCase()) ? null : name));
        }
        if (next !== v.formula) cell.value = { ...v, formula: next };
      }
    }
  }
}

function adjustAllFormulas(wb: Workbook, sheetTitle: string, e: Edit): void {
  const edit = { sheet: sheetTitle, axis: e.axis, at: e.at, count: e.count, ...(e.band ? { band: e.band } : {}) };
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    for (const rowMap of ref.sheet.rows.values()) {
      for (const cell of rowMap.values()) {
        const v = cell.value;
        if (v === null || typeof v !== 'object' || v instanceof Date || v.kind !== 'formula' || !v.formula) continue;
        const next = adjustFormulaForStructure(v.formula, ref.sheet.title, edit);
        if (next !== v.formula) cell.value = { ...v, formula: next };
      }
    }
  }
  // Rule formulas follow the cells they point at, like cell formulas do; a
  // rule on another sheet may point into this one.
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    const host = ref.sheet;
    const adjust = (f: string): string => adjustFormulaForStructure(f, host.title, edit);
    host.conditionalFormatting = host.conditionalFormatting.map((cf) => {
      const rules = cf.rules.map((rule) => {
        const formulas = rule.formulas.map(adjust);
        return formulas.every((f, i) => f === rule.formulas[i]) ? rule : { ...rule, formulas };
      });
      return rules.every((r, i) => r === cf.rules[i]) ? cf : { ...cf, rules };
    });
    host.dataValidations = host.dataValidations.map((dv) => {
      const formula1 = dv.formula1 === undefined ? undefined : adjust(dv.formula1);
      const formula2 = dv.formula2 === undefined ? undefined : adjust(dv.formula2);
      if (formula1 === dv.formula1 && formula2 === dv.formula2) return dv;
      return { ...dv, ...(formula1 === undefined ? {} : { formula1 }), ...(formula2 === undefined ? {} : { formula2 }) };
    });
  }
  wb.definedNames = wb.definedNames.map((dn) => {
    const value = adjustFormulaForStructure(dn.value, sheetTitle, edit);
    return value === dn.value ? dn : { ...dn, value };
  });
}

/**
 * Apply a structural edit to `ws`. The caller declared nothing yet; this
 * declares the whole sheet plus every other sheet's cells (their formulas
 * may reference this one) and the workbook names.
 */
export function structuralEdit(tx: Transaction, wb: Workbook, ws: Worksheet, e: Edit): void {
  declareStructural(tx, wb, ws);
  applyStructuralEdit(wb, ws, e);
}

/** What a structural edit of `ws` can reach; declare it once before a run of edits. */
export function declareStructural(tx: Transaction, wb: Workbook, ws: Worksheet): void {
  tx.wholeSheet(ws);
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet' || ref.sheet === ws) continue;
    tx.cells(ref.sheet, { r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL });
    tx.sheet(ref.sheet, 'conditionalFormatting', 'dataValidations');
    // A PivotTable elsewhere may read its source from this sheet.
    if (ref.sheet.pivotTables) tx.sheet(ref.sheet, 'pivotTables');
  }
  tx.workbook('definedNames');
}

/**
 * The mutation half of {@link structuralEdit}, for commands that insert many
 * rows in one step (Subtotal): they declare once, then apply each edit,
 * instead of snapshotting the whole sheet per edit.
 */
export function applyStructuralEdit(wb: Workbook, ws: Worksheet, e: Edit): void {
  shiftCells(ws, e);
  shiftDimensions(ws, e);
  const dropped: DroppedColumns[] = [];
  shiftSheetObjects(ws, e, dropped);
  shiftPivotTables(wb, ws, e);
  adjustAllFormulas(wb, ws.title, e);
  dropTableColumnRefs(wb, dropped);
}

/**
 * PivotTables follow the edit: a source on `ws` grows, shrinks or moves with
 * its cells, and a report on `ws` moves with its anchor. A report whose source
 * or anchor was deleted outright is dropped, leaving its last values as plain
 * cells (the editor can only save a PivotTable it can rebuild from its source).
 */
function shiftPivotTables(wb: Workbook, ws: Worksheet, e: Edit): void {
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet' || !ref.sheet.pivotTables) continue;
    const host = ref.sheet;
    host.pivotTables = ref.sheet.pivotTables.flatMap((pt) => {
      let { source, anchor, renderedRef } = pt;
      if (source.sheet.toLowerCase() === ws.title.toLowerCase()) {
        const next = mapRefText(source.ref, e);
        if (next === undefined) return [];
        source = { ...source, ref: next };
      }
      if (host === ws) {
        const next = mapRefText(anchor, e);
        if (next === undefined) return [];
        anchor = next;
        renderedRef = renderedRef === undefined ? undefined : mapRefText(renderedRef, e);
      }
      const moved: PivotTable = { ...pt, source, anchor };
      if (renderedRef === undefined) delete moved.renderedRef;
      else moved.renderedRef = renderedRef;
      return [moved];
    });
  }
}

export type { Edit as StructuralEdit };
