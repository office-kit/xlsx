// Data tab operations: AutoFilter, sorting, drop-down pick lists and Remove
// Duplicates.

import type { Cell, CellValue } from '@office-kit/xlsx/cell';
import { makeCell } from '@office-kit/xlsx/cell';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { translateFormula } from '../calc/index.ts';
import type { Range } from './address.ts';
import { colLetter, MAX_ROW, parseRangeAddress, rangeAddress } from './address.ts';
import { getCellAt, isBlank } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import { currentRegion } from './navigation.ts';
import { currentRange, selectRange } from './selection.ts';
import { autoFilterRange, filterKey } from './filter.ts';
import { headerRows, tableAt, totalRows } from './tables.ts';

/** The data block a sort/filter applies to: the selection if it spans cells, else the current region. */
export function dataRange(ctl: EditorController): Range {
  const r = currentRange(ctl.doc.selection);
  if (r.r1 !== r.r2 || r.c1 !== r.c2) {
    const used = ctl.doc.ws;
    // Whole columns shrink to the populated rows.
    if (r.r2 === MAX_ROW) {
      let last = r.r1;
      for (const row of used.rows.keys()) if (row > last) last = row;
      return { ...r, r2: last };
    }
    return r;
  }
  return currentRegion(ctl.doc.ws, ctl.doc.selection.active);
}

// ---- sorting ------------------------------------------------------------------------

export interface SortKey {
  readonly col: number;
  readonly descending: boolean;
  /** Sort on cell colour / font colour is out of scope; values only. */
  readonly by?: 'value';
  /** Custom-list order (Sort ▸ Order ▸ Custom List); text not in the list sorts after it. */
  readonly list?: readonly string[];
  /** Sort ▸ Options ▸ Case sensitive: lowercase sorts before uppercase. */
  readonly caseSensitive?: boolean;
}

function effective(v: CellValue | undefined): CellValue | number | string | boolean | undefined {
  if (v !== null && v !== undefined && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula') {
    if (v.cachedValueType === 'error') return { kind: 'error', code: String(v.cachedValue) as `#${string}` };
    return v.cachedValue;
  }
  return v ?? undefined;
}

function compareText(a: string, b: string, caseSensitive = false): number {
  return caseSensitive ? a.localeCompare(b, undefined, { sensitivity: 'variant', caseFirst: 'lower', numeric: false }) : a.localeCompare(b, undefined, { sensitivity: 'base', numeric: false });
}

/**
 * Excel's ascending order: numbers < text < FALSE < TRUE < errors, blanks
 * always last (in both directions).
 */
export function compareValues(a: CellValue | undefined, b: CellValue | undefined, descending: boolean, opts: Pick<SortKey, 'list' | 'caseSensitive'> = {}): number {
  const va = effective(a);
  const vb = effective(b);
  const blankA = va === undefined || va === null || va === '';
  const blankB = vb === undefined || vb === null || vb === '';
  if (blankA || blankB) return blankA === blankB ? 0 : blankA ? 1 : -1;
  const rank = (v: unknown): number => (typeof v === 'number' || v instanceof Date ? 0 : typeof v === 'string' || (typeof v === 'object' && v !== null && 'runs' in v) ? 1 : typeof v === 'boolean' ? 2 : 3);
  const ra = rank(va);
  const rb = rank(vb);
  let cmp: number;
  if (ra !== rb) cmp = ra - rb;
  else if (ra === 0) cmp = Number(va instanceof Date ? va.getTime() : va) - Number(vb instanceof Date ? vb.getTime() : vb);
  else if (ra === 1) cmp = opts.list ? compareInList(textOf(va), textOf(vb), opts.list) || compareText(textOf(va), textOf(vb), opts.caseSensitive) : compareText(textOf(va), textOf(vb), opts.caseSensitive);
  else if (ra === 2) cmp = Number(va) - Number(vb);
  else cmp = 0;
  return descending ? -cmp : cmp;
}

function compareInList(a: string, b: string, list: readonly string[]): number {
  const ia = list.findIndex((item) => item.toLowerCase() === a.toLowerCase());
  const ib = list.findIndex((item) => item.toLowerCase() === b.toLowerCase());
  return (ia < 0 ? list.length : ia) - (ib < 0 ? list.length : ib);
}

function textOf(v: unknown): string {
  if (typeof v === 'string') return v;
  if (v && typeof v === 'object' && 'runs' in v) return (v as { runs: Array<{ text: string }> }).runs.map((r) => r.text).join('');
  return String(v);
}

/** Does the block's first row look like a header (text over non-text data)? */
export function guessHeader(ws: Worksheet, range: Range): boolean {
  if (range.r2 <= range.r1) return false;
  let textTop = 0;
  let typed = 0;
  for (let c = range.c1; c <= range.c2; c++) {
    const top = effective(getCellAt(ws, range.r1, c)?.value);
    const below = effective(getCellAt(ws, range.r1 + 1, c)?.value);
    if (typeof top === 'string') textTop++;
    if (typeof below === 'number' || typeof below === 'boolean') typed++;
  }
  return textTop === range.c2 - range.c1 + 1 && typed > 0;
}

export function sortRange(ctl: EditorController, range: Range, keys: readonly SortKey[], hasHeader: boolean, orientation: 'rows' | 'columns' = 'rows'): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const body: Range = orientation === 'rows' ? { ...range, r1: range.r1 + (hasHeader ? 1 : 0) } : { ...range, c1: range.c1 + (hasHeader ? 1 : 0) };
  if (body.r1 > body.r2 || body.c1 > body.c2) return;
  doc.transact('Sort', (tx) => {
    tx.cells(ws, body);
    const byRows = orientation === 'rows';
    // A line is a row (byRows) or a column; its cells are keyed by the other axis.
    const lines: Array<{ src: number; cells: Map<number, Cell> }> = [];
    if (byRows) {
      for (let r = body.r1; r <= body.r2; r++) {
        const cells = new Map<number, Cell>();
        const rowMap = ws.rows.get(r);
        if (rowMap) for (const [c, cell] of rowMap) if (c >= body.c1 && c <= body.c2) cells.set(c, cell);
        lines.push({ src: r, cells });
      }
    } else {
      for (let c = body.c1; c <= body.c2; c++) {
        const cells = new Map<number, Cell>();
        for (let r = body.r1; r <= body.r2; r++) {
          const cell = getCellAt(ws, r, c);
          if (cell) cells.set(r, cell);
        }
        lines.push({ src: c, cells });
      }
    }
    lines.sort((a, b) => {
      for (const k of keys) {
        const cmp = compareValues(a.cells.get(k.col)?.value, b.cells.get(k.col)?.value, k.descending, k);
        if (cmp !== 0) return cmp;
      }
      return a.src - b.src;
    });
    for (let r = body.r1; r <= body.r2; r++) {
      const rowMap = ws.rows.get(r);
      if (rowMap) for (let c = body.c1; c <= body.c2; c++) rowMap.delete(c);
    }
    lines.forEach((line, i) => {
      const at = (byRows ? body.r1 : body.c1) + i;
      for (const [k, cell] of line.cells) {
        if (byRows) placeMoved(ws, cell, at, k, at - line.src, 0);
        else placeMoved(ws, cell, k, at, 0, at - line.src);
      }
    });
  });
  ctl.doc.setSelection(selectRange(range, ctl.doc.selection.active));
}

/** Put a sorted cell at its new position; relative references move with it, as in Excel. */
function placeMoved(ws: Worksheet, cell: Cell, row: number, col: number, dRow: number, dCol: number): void {
  let value = cell.value;
  if ((dRow || dCol) && value !== null && typeof value === 'object' && !(value instanceof Date) && value.kind === 'formula' && value.formula) {
    value = { ...value, formula: translateFormula(value.formula, dRow, dCol) };
  }
  const next = makeCell(row, col, value, cell.styleId);
  if (cell.hyperlinkId !== undefined) next.hyperlinkId = cell.hyperlinkId;
  if (cell.commentId !== undefined) next.commentId = cell.commentId;
  let rowMap = ws.rows.get(row);
  if (!rowMap) {
    rowMap = new Map();
    ws.rows.set(row, rowMap);
  }
  rowMap.set(col, next);
}

/** Sort A→Z / Z→A on the active cell's column, as the ribbon buttons do. */
export function quickSort(ctl: EditorController, descending: boolean): void {
  const ws = ctl.doc.ws;
  const af = autoFilterRange(ws);
  const active = ctl.doc.selection.active;
  const range = af && active.row >= af.r1 && active.row <= af.r2 && active.col >= af.c1 && active.col <= af.c2 ? af : dataRange(ctl);
  const header = af === range ? true : guessHeader(ws, range);
  const col = Math.min(Math.max(active.col, range.c1), range.c2);
  sortRange(ctl, range, [{ col, descending }], header);
}

// ---- pick lists ---------------------------------------------------------------------

/** Items for Alt+Down: a list validation's source if one applies, else the column's texts. */
export function pickListItems(ctl: EditorController, row: number, col: number): string[] {
  const ws = ctl.doc.ws;
  const dv = validationAt(ws, row, col);
  if (dv?.type === 'list' && dv.formula1) return listSource(ctl, dv.formula1);
  const seen = new Set<string>();
  for (let r = row - 1; r >= 1 && !isBlank(getCellAt(ws, r, col)); r--) {
    const v = getCellAt(ws, r, col)?.value;
    if (typeof v === 'string') seen.add(v);
  }
  for (let r = row + 1; r <= MAX_ROW && !isBlank(getCellAt(ws, r, col)); r++) {
    const v = getCellAt(ws, r, col)?.value;
    if (typeof v === 'string') seen.add(v);
  }
  return [...seen].sort(compareText);
}

export function validationAt(ws: Worksheet, row: number, col: number) {
  return ws.dataValidations.find((dv) => dv.sqref.ranges.some((r) => row >= r.minRow && row <= r.maxRow && col >= r.minCol && col <= r.maxCol));
}

/** A list validation's items: a quoted comma list, or a range/name reference evaluated to values. */
export function listSource(ctl: EditorController, formula1: string): string[] {
  const text = formula1.trim();
  if (text.startsWith('"') && text.endsWith('"')) return text.slice(1, -1).split(',').map((s) => s.trim());
  const ws = ctl.doc.ws;
  const ref = text.replace(/^=/, '');
  const parsed = parseRangeAddress(ref);
  const targetRef = parsed?.sheet === undefined ? undefined : ctl.doc.wb.sheets.find((s) => s.sheet.title === parsed.sheet);
  const target = parsed?.sheet === undefined ? ws : targetRef?.kind === 'worksheet' ? targetRef.sheet : undefined;
  if (parsed && target) {
    const out: string[] = [];
    const { range } = parsed;
    // A list longer than this is unusable as a drop-down anyway.
    for (let r = range.r1; r <= Math.min(range.r2, range.r1 + 10_000); r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const cell = getCellAt(target, r, c);
        if (cell && !isBlank(cell)) out.push(getCellDisplayText(ctl.doc.wb, cell));
      }
    }
    return out;
  }
  // Defined names and other formulas go through the engine.
  const sheetTitle = ws.title;
  const value = ctl.doc.calc.evaluate(ref, sheetTitle, 1, 1);
  return value === null ? [] : [String(value)];
}

// ---- remove duplicates -----------------------------------------------------------

/**
 * Data ▸ Remove Duplicates, with Excel's rules: rows compare on the displayed
 * text of `cols`, the first of each set is kept, and a removed row goes
 * across every column of the range (or table) while nothing outside it moves.
 * Inside a table the whole table body is the target and the table shrinks,
 * its total row moving up with it.
 *
 * Returns the number of rows removed, `'outlined'` when the rows carry an
 * outline or SUBTOTAL formulas (Excel refuses those), or undefined when sheet
 * protection refused the edit.
 */
export function removeDuplicates(ctl: EditorController, range: Range, cols: readonly number[], hasHeader: boolean): number | 'outlined' | undefined {
  const ws = ctl.doc.ws;
  const table = tableAt(ws, range.r1, range.c1);
  const c1 = table ? table.range.c1 : range.c1;
  const c2 = table ? table.range.c2 : range.c2;
  const start = table ? table.range.r1 + headerRows(table.def) : range.r1 + (hasHeader ? 1 : 0);
  const last = table ? table.range.r2 - totalRows(table.def) : range.r2;
  // Everything that can move: the body plus a table's total row.
  const bottom = table ? table.range.r2 : last;
  // A table's total row is SUBTOTAL formulas by design; only the body is checked.
  if (isOutlined(ws, { r1: start, c1, r2: last, c2 })) return 'outlined';

  const keep: number[] = [];
  const seen = new Set<string>();
  for (let r = start; r <= last; r++) {
    const key = cols.map((c) => filterKey(ctl, getCellAt(ws, r, c)).toLowerCase()).join('\u0000');
    if (seen.has(key)) continue;
    seen.add(key);
    keep.push(r);
  }
  const removed = last - start + 1 - keep.length;
  if (removed === 0) return 0;

  // Old row -> new row for everything that stays; removed rows are absent.
  const moved = new Map<number, number>(keep.map((r, i) => [r, start + i]));
  for (let r = last + 1; r <= bottom; r++) moved.set(r, r - removed);

  return ctl.doc.transact('Remove Duplicates', (tx) => {
    tx.cells(ws, { r1: start, c1, r2: bottom, c2 });
    tx.sheet(ws, 'tables', 'hyperlinks', 'legacyComments', 'threadedComments');
    const rows = [...moved].map(([src, dest]) => {
      const cells: Cell[] = [];
      for (let c = c1; c <= c2; c++) {
        const cell = getCellAt(ws, src, c);
        if (cell) cells.push(cell);
      }
      return { src, dest, cells };
    });
    for (let r = start; r <= bottom; r++) {
      const rowMap = ws.rows.get(r);
      if (!rowMap) continue;
      for (let c = c1; c <= c2; c++) rowMap.delete(c);
      if (rowMap.size === 0) ws.rows.delete(r);
    }
    for (const { src, dest, cells } of rows) {
      // A total row keeps its formulas as written: they point at the column, not at a row.
      const dRow = src > last ? 0 : dest - src;
      for (const cell of cells) placeMoved(ws, cell, dest, cell.col, dRow, 0);
    }
    const remap = (ref: string): string | null => {
      const at = parseRangeAddress(ref)?.range;
      // Only single-row refs inside the moved block follow their row.
      if (!at || at.r1 !== at.r2 || at.r1 < start || at.r1 > bottom || at.c1 < c1 || at.c2 > c2) return ref;
      const dest = moved.get(at.r1);
      return dest === undefined ? null : rangeAddress({ ...at, r1: dest, r2: dest });
    };
    ws.hyperlinks = ws.hyperlinks.flatMap((h) => {
      const ref = remap(h.ref);
      return ref === null ? [] : [ref === h.ref ? h : { ...h, ref }];
    });
    ws.legacyComments = ws.legacyComments.flatMap((n) => {
      const ref = remap(n.ref);
      return ref === null ? [] : [ref === n.ref ? n : { ...n, ref }];
    });
    ws.threadedComments = (ws.threadedComments ?? []).flatMap((n) => {
      const ref = remap(n.ref);
      return ref === null ? [] : [ref === n.ref ? n : { ...n, ref }];
    });
    if (table) {
      const { def } = table;
      def.ref = rangeAddress({ ...table.range, r2: table.range.r2 - removed });
      const filter = def.autoFilter ? parseRangeAddress(def.autoFilter.ref)?.range : undefined;
      if (def.autoFilter && filter) def.autoFilter = { ...def.autoFilter, ref: rangeAddress({ ...filter, r2: filter.r2 - removed }) };
    }
    return removed;
  });
}

/** Excel's Remove Duplicates refuses grouped rows or columns and subtotalled data. */
function isOutlined(ws: Worksheet, range: Range): boolean {
  for (const [r, d] of ws.rowDimensions) if (d.outlineLevel && r >= range.r1 && r <= range.r2) return true;
  for (const d of ws.columnDimensions.values()) if (d.outlineLevel && d.max >= range.c1 && d.min <= range.c2) return true;
  for (let r = range.r1; r <= range.r2; r++) {
    const rowMap = ws.rows.get(r);
    if (!rowMap) continue;
    for (const [c, cell] of rowMap) {
      const v = cell.value;
      if (c >= range.c1 && c <= range.c2 && v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' && /^SUBTOTAL\(/i.test(v.formula)) return true;
    }
  }
  return false;
}

/** Column header label for dialogs: the header text, or "Column B". */
export function columnLabel(ctl: EditorController, range: Range, col: number, hasHeader: boolean, columnWord: string): string {
  if (hasHeader) {
    const cell = getCellAt(ctl.doc.ws, range.r1, col);
    if (cell && !isBlank(cell)) return getCellDisplayText(ctl.doc.wb, cell);
  }
  return `${columnWord} ${colLetter(col)}`;
}

