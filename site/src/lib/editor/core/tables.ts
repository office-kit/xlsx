// Excel tables (Insert ▸ Table) after creation: the Table Design tab's
// commands — rename, resize, style options, the total row and its functions,
// Convert to Range — plus the structured-reference rewriting they need.

import type { CellValue } from '@office-kit/xlsx/cell';
import type { Workbook } from '@office-kit/xlsx/workbook';
import type { TableColumn, TableDefinition, Worksheet } from '@office-kit/xlsx/worksheet';
import type { MessageKey } from '../i18n/i18n.svelte.ts';
import { colLetter, MAX_ROW, parseRangeAddress, rangeAddress, rangesIntersect, type Range } from './address.ts';
import { getCellAt, isBlank, setValueAt } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import type { Transaction } from './history.ts';
import { transformStyle } from './format.ts';
import { structuralEdit } from './structure.ts';
import { hasBuiltInStyle, placeOf, rawLookAt, tableOptions, type Line, type ThemeRef } from './table-style.ts';

export interface TableHit {
  readonly def: TableDefinition;
  readonly range: Range;
}

export function tableRange(def: TableDefinition): Range | undefined {
  return parseRangeAddress(def.ref)?.range;
}

/** The table covering (row, col) on `ws`, if any. */
export function tableAt(ws: Worksheet, row: number, col: number): TableHit | undefined {
  for (const def of ws.tables) {
    const range = tableRange(def);
    if (range && row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2) return { def, range };
  }
  return undefined;
}

export function headerRows(def: TableDefinition): number {
  return def.headerRowCount ?? 1;
}

export function totalRows(def: TableDefinition): number {
  return def.totalsRowCount ?? 0;
}

/** Header + data rows, the part an AutoFilter covers. */
function filterRange(def: TableDefinition, range: Range): Range {
  return { ...range, r2: range.r2 - totalRows(def) };
}

// ---- names ------------------------------------------------------------------

// Excel's name rules: a letter, underscore or backslash first, then letters,
// digits, periods and underscores; never something that reads as a cell
// reference (A1, R1C1, or the bare R / C shorthands).
const NAME_RE = /^[\p{L}_\\][\p{L}\p{N}._]*$/u;
const A1_RE = /^[A-Za-z]{1,3}\d+$/;
const R1C1_RE = /^[Rr](\d*)[Cc](\d*)$|^[Rr]\d*$|^[Cc]\d*$/;
const NAME_MAX = 255;

export function tableNameError(wb: Workbook, name: string, except?: TableDefinition): MessageKey | undefined {
  if (name === '') return 'dtTableNameEmpty';
  if (name.length > NAME_MAX || !NAME_RE.test(name) || A1_RE.test(name) || R1C1_RE.test(name)) return 'dtTableNameInvalid';
  const lower = name.toLowerCase();
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    for (const t of ref.sheet.tables) if (t !== except && t.displayName.toLowerCase() === lower) return 'dtTableNameTaken';
  }
  if (wb.definedNames.some((d) => d.name.toLowerCase() === lower)) return 'dtTableNameTaken';
  return undefined;
}

// ---- structured references ----------------------------------------------------

const IDENT_CHAR = /[\p{L}\p{N}._\\]/u;

/** End index (exclusive) of the bracket group opening at `open`, honouring `'` escapes. */
function bracketEnd(text: string, open: number): number {
  let depth = 0;
  for (let i = open; i < text.length; i++) {
    const ch = text[i];
    if (ch === "'") {
      i++;
      continue;
    }
    if (ch === '[') depth++;
    else if (ch === ']') {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  return -1;
}

interface StructuredRef {
  readonly start: number;
  readonly end: number;
  /** Table name as written, or undefined for an unqualified `[…]` inside the table. */
  readonly table: string | undefined;
  readonly spec: string;
}

/** Every structured reference in `formula`, skipping string literals and quoted sheet names. */
function structuredRefs(formula: string): StructuredRef[] {
  const out: StructuredRef[] = [];
  let i = 0;
  while (i < formula.length) {
    const ch = formula.charAt(i);
    if (ch === '"' || ch === "'") {
      const close = formula.indexOf(ch, i + 1);
      i = close < 0 ? formula.length : close + 1;
      continue;
    }
    if (ch === '[') {
      const end = bracketEnd(formula, i);
      if (end < 0) break;
      // A name right before the bracket qualifies it; otherwise it is unqualified.
      let s = i;
      while (s > 0 && IDENT_CHAR.test(formula.charAt(s - 1))) s--;
      const table = s < i ? formula.slice(s, i) : undefined;
      // `Book.xlsx]Sheet` style external references are not table references.
      if (!(s > 0 && formula.charAt(s - 1) === '!')) out.push({ start: s, end, table, spec: formula.slice(i, end) });
      i = end;
      continue;
    }
    i++;
  }
  return out;
}

const COLUMN_TOKEN = /\[((?:'.|[^\]'])*)\]/g;

/**
 * Rewrite the column names in `table`'s structured references: `rename`
 * returns the new name, the same name, or null for a deleted column, which
 * turns the whole reference into #REF! as Excel does. `inTable` says whether
 * an unqualified `[@Col]` in this formula belongs to `table` (its cell is
 * inside it).
 */
export function mapTableColumnsInFormula(formula: string, table: string, inTable: boolean, rename: (name: string) => string | null): string {
  const lower = table.toLowerCase();
  const refs = structuredRefs(formula).filter((r) => (r.table === undefined ? inTable : r.table.toLowerCase() === lower));
  if (refs.length === 0) return formula;
  let out = '';
  let at = 0;
  for (const r of refs) {
    const prefixLen = r.end - r.start - r.spec.length;
    const inner = r.spec.slice(1, -1);
    let deleted = false;
    const mapName = (raw: string): string => {
      if (raw.startsWith('#')) return raw;
      const thisRow = raw.startsWith('@') ? '@' : '';
      const name = unescapeColumn(raw.slice(thisRow.length));
      const next = rename(name);
      if (next === null) deleted = true;
      return thisRow + escapeColumn(next ?? name);
    };
    // `[Col]` / `[@Col]` hold one name; `[[#Data],[A]:[B]]` a bracketed list.
    const spec = inner.includes('[') ? `[${inner.replace(COLUMN_TOKEN, (_m, raw: string) => `[${mapName(raw)}]`)}]` : `[${mapName(inner)}]`;
    out += formula.slice(at, r.start) + (deleted ? '#REF!' : formula.slice(r.start, r.start + prefixLen) + spec);
    at = r.end;
  }
  return out + formula.slice(at);
}

/**
 * Header cells typed over in this transaction rename their table columns, and
 * every structured reference to the old name follows (Excel keeps formulas
 * working across a header rename).
 */
export function renameEditedHeaders(wb: Workbook, tx: Transaction, displayText: (ws: Worksheet, row: number, col: number) => string): void {
  const renames: Array<{ table: TableDefinition; ws: Worksheet; from: string; to: string }> = [];
  for (const part of tx.parts) {
    if (part.kind !== 'cells') continue;
    for (const def of part.ws.tables) {
      const range = tableRange(def);
      if (!range || headerRows(def) === 0 || range.r1 < part.range.r1 || range.r1 > part.range.r2) continue;
      const taken = new Set(def.columns.map((c) => c.name.toLowerCase()));
      def.columns.forEach((c, i) => {
        const col = range.c1 + i;
        if (col < part.range.c1 || col > part.range.c2) return;
        const text = displayText(part.ws, range.r1, col);
        if (text === '' || text === c.name || taken.has(text.toLowerCase())) return;
        taken.add(text.toLowerCase());
        renames.push({ table: def, ws: part.ws, from: c.name, to: text });
      });
    }
  }
  if (renames.length === 0) return;
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    tx.cells(ref.sheet, { r1: 1, c1: 1, r2: MAX_ROW, c2: 16_384 });
  }
  for (const ws of new Set(renames.map((r) => r.ws))) tx.sheet(ws, 'tables');
  for (const { table, ws, from, to } of renames) {
    const range = tableRange(table);
    for (const ref of wb.sheets) {
      if (ref.kind !== 'worksheet') continue;
      for (const rowMap of ref.sheet.rows.values()) {
        for (const cell of rowMap.values()) {
          const v = cell.value;
          if (v === null || typeof v !== 'object' || v instanceof Date || v.kind !== 'formula' || !v.formula.includes('[')) continue;
          const inTable = ref.sheet === ws && range !== undefined && cell.row >= range.r1 && cell.row <= range.r2 && cell.col >= range.c1 && cell.col <= range.c2;
          const next = mapTableColumnsInFormula(v.formula, table.displayName, inTable, (n) => (n.toLowerCase() === from.toLowerCase() ? to : n));
          if (next !== v.formula) cell.value = { ...v, formula: next };
        }
      }
    }
    const column = table.columns.find((c) => c.name === from);
    if (column) column.name = to;
  }
}

/** Rewrite `Old[` to `New[` in a formula; returns the input when nothing changed. */
export function renameTableInFormula(formula: string, from: string, to: string): string {
  const refs = structuredRefs(formula).filter((r) => r.table !== undefined && r.table.toLowerCase() === from.toLowerCase());
  if (refs.length === 0) return formula;
  let out = '';
  let at = 0;
  for (const r of refs) {
    out += formula.slice(at, r.start) + to;
    at = r.start + (r.table?.length ?? 0);
  }
  return out + formula.slice(at);
}

function unescapeColumn(text: string): string {
  return text.replace(/'(.)/g, '$1');
}

/** Escape a column name for use inside `[...]`, as Excel writes it. */
export function escapeColumn(name: string): string {
  return name.replace(/['#[\]@]/g, "'$&");
}

/**
 * The A1 text a structured reference resolves to inside `def` at `range`, or
 * undefined for a form this does not understand (it is then left alone).
 */
function structuredToA1(def: TableDefinition, range: Range, spec: string, row: number): string | undefined {
  const inner = spec.slice(1, -1).trim();
  const items: string[] = [];
  if (inner.startsWith('[') || inner.startsWith('@[')) {
    const thisRow = inner.startsWith('@');
    if (thisRow) items.push('#This Row');
    for (const m of (thisRow ? inner.slice(1) : inner).matchAll(/\[((?:'.|[^\]])*)\]|([,:])/g)) {
      if (m[1] !== undefined) items.push(m[1]);
      else if (m[2] === ':') items.push(':');
    }
  } else if (inner.startsWith('@')) items.push('#This Row', inner.slice(1));
  else if (inner !== '') items.push(inner);

  const h = headerRows(def);
  const t = totalRows(def);
  let r1 = range.r1 + h;
  let r2 = range.r2 - t;
  const cols: number[] = [];
  let colon = false;
  for (const raw of items) {
    if (raw === ':') {
      colon = true;
      continue;
    }
    const item = unescapeColumn(raw).trim();
    switch (item.toLowerCase()) {
      case '#all':
        r1 = range.r1;
        r2 = range.r2;
        continue;
      case '#data':
        continue;
      case '#headers':
        if (h === 0) return undefined;
        r1 = r2 = range.r1;
        continue;
      case '#totals':
        if (t === 0) return undefined;
        r1 = r2 = range.r2;
        continue;
      case '#this row':
        r1 = r2 = row;
        continue;
    }
    const idx = def.columns.findIndex((c) => c.name.toLowerCase() === item.toLowerCase());
    if (idx < 0) return undefined;
    if (colon && cols.length > 0) {
      const from = cols.pop() ?? idx;
      for (let k = Math.min(from, idx); k <= Math.max(from, idx); k++) cols.push(k);
    } else cols.push(idx);
    colon = false;
  }
  const c1 = range.c1 + (cols.length > 0 ? Math.min(...cols) : 0);
  const c2 = range.c1 + (cols.length > 0 ? Math.max(...cols) : range.c2 - range.c1);
  if (r1 === r2 && items.some((s) => s.toLowerCase() === '#this row')) {
    // A this-row reference stays row-relative, the way Excel converts `[@Col]`.
    const left = `$${colLetter(c1)}${r1}`;
    return c1 === c2 ? left : `${left}:$${colLetter(c2)}${r1}`;
  }
  return rangeAddress({ r1, c1, r2, c2 }, true);
}

/** Replace references to `def` in `formula` (sitting at `row`; `insideTable` allows unqualified ones) by A1 references. */
export function structuredToA1Formula(formula: string, def: TableDefinition, range: Range, row: number, insideTable: boolean): string {
  const name = def.displayName.toLowerCase();
  let out = '';
  let at = 0;
  for (const r of structuredRefs(formula)) {
    const mine = r.table === undefined ? insideTable : r.table.toLowerCase() === name;
    if (!mine) continue;
    const a1 = structuredToA1(def, range, r.spec, row);
    if (a1 === undefined) continue;
    out += formula.slice(at, r.start) + a1;
    at = r.end;
  }
  return at === 0 ? formula : out + formula.slice(at);
}

function isFormulaValue(v: CellValue): v is Extract<CellValue, { kind: 'formula' }> {
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula';
}

/**
 * Rewrite every formula in the workbook through `fn`, declaring only the
 * cells that change on the transaction (a full-sheet snapshot would cost far
 * more than the edit).
 */
function rewriteFormulas(tx: Transaction, wb: Workbook, fn: (formula: string, ws: Worksheet, row: number, col: number) => string): void {
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    const ws = ref.sheet;
    for (const rowMap of ws.rows.values()) {
      for (const cell of rowMap.values()) {
        const v = cell.value;
        if (!isFormulaValue(v) || !v.formula) continue;
        const next = fn(v.formula, ws, cell.row, cell.col);
        if (next === v.formula) continue;
        tx.cells(ws, { r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });
        cell.value = { ...v, formula: next };
      }
    }
  }
}

export function renameTable(ctl: EditorController, def: TableDefinition, name: string): MessageKey | undefined {
  const doc = ctl.doc;
  if (name === def.displayName) return undefined;
  const err = tableNameError(doc.wb, name, def);
  if (err) return err;
  const from = def.displayName;
  doc.transact('Rename Table', (tx) => {
    tx.sheet(doc.ws, 'tables');
    tx.workbook('definedNames');
    tx.structural = true;
    def.displayName = name;
    def.name = name;
    rewriteFormulas(tx, doc.wb, (f) => renameTableInFormula(f, from, name));
    for (const d of doc.wb.definedNames) d.value = renameTableInFormula(d.value, from, name);
  });
  return undefined;
}

// ---- headers ----------------------------------------------------------------------

/** Text a header cell names its column by. */
function headerText(ctl: EditorController, ws: Worksheet, row: number, col: number): string {
  const cell = getCellAt(ws, row, col);
  return cell && !isBlank(cell) ? ctl.displayText(cell).trim() : '';
}

/** Unique column names, blank ones becoming ColumnN, the way Excel names them. */
function uniqueNames(names: readonly string[], taken = new Set<string>()): string[] {
  return names.map((raw, i) => {
    const base = raw === '' ? `Column${i + 1}` : raw;
    let name = base;
    for (let k = 2; taken.has(name.toLowerCase()); k++) name = `${base}${k}`;
    taken.add(name.toLowerCase());
    return name;
  });
}

/**
 * Make each table's column names match its header cells before saving: Excel
 * treats a mismatch as a damaged table and drops it on open. Blank headers get
 * Excel's ColumnN name written into the cell as well.
 */
export function syncTableHeaders(wb: Workbook, displayText: (ws: Worksheet, row: number, col: number) => string): void {
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    const ws = ref.sheet;
    for (const def of ws.tables) {
      const range = tableRange(def);
      if (!range || headerRows(def) === 0) continue;
      const texts = def.columns.map((_, i) => displayText(ws, range.r1, range.c1 + i));
      const names = uniqueNames(texts);
      def.columns.forEach((c, i) => {
        const name = names[i] ?? c.name;
        if (texts[i] !== name) setValueAt(ws, range.r1, range.c1 + i, name);
        c.name = name;
      });
    }
  }
}

// ---- style options -------------------------------------------------------------------

export type StyleOption = 'bandedRows' | 'bandedCols' | 'firstCol' | 'lastCol';

const STYLE_KEYS = {
  bandedRows: 'showRowStripes',
  bandedCols: 'showColumnStripes',
  firstCol: 'showFirstColumn',
  lastCol: 'showLastColumn',
} as const;

export function styleOption(def: TableDefinition, opt: StyleOption): boolean {
  return def.styleInfo?.[STYLE_KEYS[opt]] === true;
}

export function setStyleOption(ctl: EditorController, def: TableDefinition, opt: StyleOption, on: boolean): void {
  ctl.doc.transact('Table Style Options', (tx) => {
    tx.sheet(ctl.doc.ws, 'tables');
    def.styleInfo = { ...def.styleInfo, [STYLE_KEYS[opt]]: on };
  });
}

export function setTableStyle(ctl: EditorController, def: TableDefinition, name: string | undefined): void {
  ctl.doc.transact('Table Style', (tx) => {
    tx.sheet(ctl.doc.ws, 'tables');
    const { name: _old, ...rest } = def.styleInfo ?? {};
    def.styleInfo = name === undefined ? rest : { ...rest, name };
  });
}

export function setFilterButton(ctl: EditorController, def: TableDefinition, on: boolean): void {
  const range = tableRange(def);
  if (!range || headerRows(def) === 0) return;
  ctl.doc.transact('Filter Button', (tx) => {
    tx.sheet(ctl.doc.ws, 'tables', 'rowDimensions');
    if (on) def.autoFilter = { ref: rangeAddress(filterRange(def, range)), filterColumns: [] };
    else {
      // Removing the buttons also drops their filtering, as in Excel.
      for (let r = range.r1 + 1; r <= range.r2; r++) {
        const dim = ctl.doc.ws.rowDimensions.get(r);
        if (!dim?.hidden) continue;
        const { hidden: _h, ...rest } = dim;
        if (Object.keys(rest).length === 0) ctl.doc.ws.rowDimensions.delete(r);
        else ctl.doc.ws.rowDimensions.set(r, rest);
      }
      delete def.autoFilter;
    }
  });
}

/** Whether `row` within the table's columns is free to take a header or total row. */
function rowIsFree(ws: Worksheet, row: number, range: Range): boolean {
  if (row < 1 || row > MAX_ROW) return false;
  for (let c = range.c1; c <= range.c2; c++) if (!isBlank(getCellAt(ws, row, c))) return false;
  return !ws.tables.some((t) => {
    const r = tableRange(t);
    return r !== undefined && rangesIntersect(r, { ...range, r1: row, r2: row });
  });
}

/**
 * Header Row on/off. Off empties the header row and leaves it outside the
 * table; on writes the column names back into the row above, inserting cells
 * when that row is taken.
 */
export function setHeaderRow(ctl: EditorController, def: TableDefinition, on: boolean): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const range = tableRange(def);
  if (!range || (headerRows(def) > 0) === on) return;
  doc.transact('Header Row', (tx) => {
    if (!on) {
      tx.cells(ws, { ...range, r2: range.r1 });
      tx.sheet(ws, 'tables');
      for (let c = range.c1; c <= range.c2; c++) ws.rows.get(range.r1)?.delete(c);
      def.headerRowCount = 0;
      delete def.autoFilter;
      def.ref = rangeAddress({ ...range, r1: range.r1 + 1 });
      return;
    }
    let top = range.r1 - 1;
    if (rowIsFree(ws, top, range)) {
      tx.cells(ws, { ...range, r1: top, r2: top });
      tx.sheet(ws, 'tables');
    } else {
      top = range.r1;
      structuralEdit(tx, doc.wb, ws, { axis: 'row', at: range.r1, count: 1, band: { from: range.c1, to: range.c2 } });
    }
    const full: Range = { r1: top, c1: range.c1, r2: top + (range.r2 - range.r1) + 1, c2: range.c2 };
    def.columns.forEach((c, i) => setValueAt(ws, top, range.c1 + i, c.name, ctl.defaultStyleAt(top, range.c1 + i)));
    delete def.headerRowCount;
    def.ref = rangeAddress(full);
    def.autoFilter = { ref: rangeAddress(filterRange(def, full)), filterColumns: [] };
  });
}

// ---- total row --------------------------------------------------------------------

export type TotalFunction = NonNullable<TableColumn['totalsRowFunction']>;

/** The Total Row drop-down: Excel's SUBTOTAL codes that skip hidden rows. */
export const TOTAL_FUNCTIONS: ReadonlyArray<{ readonly fn: Exclude<TotalFunction, 'custom'>; readonly code: number; readonly label: MessageKey }> = [
  { fn: 'average', code: 101, label: 'dtTotAverage' },
  { fn: 'count', code: 103, label: 'dtTotCount' },
  { fn: 'countNums', code: 102, label: 'dtTotCountNums' },
  { fn: 'max', code: 104, label: 'dtTotMax' },
  { fn: 'min', code: 105, label: 'dtTotMin' },
  { fn: 'sum', code: 109, label: 'dtTotSum' },
  { fn: 'stdDev', code: 107, label: 'dtTotStdDev' },
  { fn: 'var', code: 110, label: 'dtTotVar' },
];

export function totalFormula(def: TableDefinition, column: TableColumn, fn: Exclude<TotalFunction, 'custom'>): string {
  const code = TOTAL_FUNCTIONS.find((f) => f.fn === fn)?.code ?? 109;
  return `SUBTOTAL(${code},${def.displayName}[${escapeColumn(column.name)}])`;
}

/** Excel's default: a label in the first column and a sum (or count, for text) in the last. */
function defaultTotals(ctl: EditorController, def: TableDefinition, range: Range, label: string): void {
  const ws = ctl.doc.ws;
  if (def.columns.some((c) => c.totalsRowFunction !== undefined || c.totalsRowLabel !== undefined)) return;
  const first = def.columns[0];
  const last = def.columns[def.columns.length - 1];
  if (!first || !last) return;
  first.totalsRowLabel = label;
  if (last === first) return;
  let numeric = false;
  for (let r = range.r1 + headerRows(def); r <= range.r2 && !numeric; r++) {
    const v = getCellAt(ws, r, range.c2)?.value ?? null;
    numeric = typeof v === 'number' || (isFormulaValue(v) && typeof v.cachedValue === 'number');
  }
  last.totalsRowFunction = numeric ? 'sum' : 'count';
}

function writeTotalRow(ctl: EditorController, def: TableDefinition, row: number, range: Range): void {
  const ws = ctl.doc.ws;
  def.columns.forEach((c, i) => {
    const col = range.c1 + i;
    let value: CellValue = null;
    if (c.totalsRowFunction && c.totalsRowFunction !== 'custom') value = { kind: 'formula', t: 'normal', formula: totalFormula(def, c, c.totalsRowFunction) };
    else if (c.totalsRowFunction === 'custom' && c.totalsRowFormula) value = { kind: 'formula', t: 'normal', formula: c.totalsRowFormula };
    else if (c.totalsRowLabel !== undefined) value = c.totalsRowLabel;
    if (value !== null) setValueAt(ws, row, col, value, ctl.defaultStyleAt(row, col));
  });
}

/**
 * Total Row on/off. On uses the row below the table, inserting cells when it
 * is taken; `label` is the first column's caption ("Total" / "集計").
 */
export function setTotalRow(ctl: EditorController, def: TableDefinition, on: boolean, label: string): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const range = tableRange(def);
  if (!range || (totalRows(def) > 0) === on) return;
  doc.transact('Total Row', (tx) => {
    if (!on) {
      tx.cells(ws, { ...range, r1: range.r2 });
      tx.sheet(ws, 'tables');
      for (let c = range.c1; c <= range.c2; c++) ws.rows.get(range.r2)?.delete(c);
      delete def.totalsRowCount;
      def.ref = rangeAddress({ ...range, r2: range.r2 - 1 });
      return;
    }
    const row = range.r2 + 1;
    let table = def;
    if (rowIsFree(ws, row, range)) {
      tx.cells(ws, { ...range, r1: row, r2: row });
      tx.sheet(ws, 'tables');
      // A total row's formulas bring new dependencies.
      tx.structural = true;
    } else {
      structuralEdit(tx, doc.wb, ws, { axis: 'row', at: row, count: 1, band: { from: range.c1, to: range.c2 } });
      // The edit rebuilt the sheet's table list; keep changing the live definition.
      table = ws.tables.find((t) => t.id === def.id) ?? def;
    }
    const full = { ...range, r2: row };
    table.totalsRowCount = 1;
    defaultTotals(ctl, table, range, label);
    table.ref = rangeAddress(full);
    if (table.autoFilter) table.autoFilter = { ...table.autoFilter, ref: rangeAddress(range) };
    writeTotalRow(ctl, table, row, full);
  });
}

/** A total-row cell's drop-down choice; 'none' clears it. */
export function setTotalFunction(ctl: EditorController, def: TableDefinition, col: number, fn: Exclude<TotalFunction, 'custom'> | 'none'): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const range = tableRange(def);
  const column = range ? def.columns[col - range.c1] : undefined;
  if (!range || !column || totalRows(def) === 0) return;
  doc.transact('Total Row', (tx) => {
    tx.cells(ws, { r1: range.r2, c1: col, r2: range.r2, c2: col });
    tx.sheet(ws, 'tables');
    delete column.totalsRowFormula;
    delete column.totalsRowLabel;
    if (fn === 'none') {
      delete column.totalsRowFunction;
      const cell = getCellAt(ws, range.r2, col);
      if (cell) cell.value = null;
      return;
    }
    column.totalsRowFunction = fn;
    setValueAt(ws, range.r2, col, { kind: 'formula', t: 'normal', formula: totalFormula(def, column, fn) }, ctl.defaultStyleAt(range.r2, col));
  });
}

/** The total-row function at (row, col) if it is a total-row cell of a table. */
export function totalCellAt(ws: Worksheet, row: number, col: number): { def: TableDefinition; column: TableColumn } | undefined {
  const hit = tableAt(ws, row, col);
  if (!hit || totalRows(hit.def) === 0 || row !== hit.range.r2) return undefined;
  const column = hit.def.columns[col - hit.range.c1];
  return column ? { def: hit.def, column } : undefined;
}

// ---- resize / convert --------------------------------------------------------------

/**
 * Resize Table: the header row stays where it is and the new range must
 * overlap the old one. Columns keep their identity by position; new ones take
 * their header text (or ColumnN).
 */
export function resizeTable(ctl: EditorController, def: TableDefinition, next: Range): MessageKey | undefined {
  const doc = ctl.doc;
  const ws = doc.ws;
  const range = tableRange(def);
  if (!range) return 'invalidReference';
  if (next.r1 !== range.r1) return 'dtResizeHeaderRow';
  if (!rangesIntersect(next, range)) return 'dtResizeOverlap';
  if (next.r2 - next.r1 + 1 <= headerRows(def) + totalRows(def)) return 'dtResizeTooSmall';
  const others = ws.tables.some((t) => {
    if (t === def) return false;
    const r = tableRange(t);
    return r !== undefined && rangesIntersect(r, next);
  });
  if (others) return 'dlgTableOverlap';
  if (ws.mergedCells.some((m) => rangesIntersect({ r1: m.minRow, c1: m.minCol, r2: m.maxRow, c2: m.maxCol }, next))) return 'dlgTableMerged';
  doc.transact('Resize Table', (tx) => {
    tx.cells(ws, { ...next, r2: next.r1 });
    tx.sheet(ws, 'tables');
    tx.structural = true;
    let nextId = def.columns.reduce((m, c) => Math.max(m, c.id), 0) + 1;
    const byOffset = new Map(def.columns.map((c, i) => [range.c1 + i, c]));
    const kept: TableColumn[] = [];
    const fresh: number[] = [];
    for (let col = next.c1; col <= next.c2; col++) {
      const old = byOffset.get(col);
      if (old) kept.push(old);
      else fresh.push(col);
    }
    const taken = new Set(kept.map((c) => c.name.toLowerCase()));
    const columns: TableColumn[] = [];
    for (let col = next.c1; col <= next.c2; col++) {
      const old = byOffset.get(col);
      if (old) {
        columns.push(old);
        continue;
      }
      const text = headerRows(def) > 0 ? headerText(ctl, ws, next.r1, col) : '';
      const [name = `Column${col - next.c1 + 1}`] = uniqueNames([text === '' ? `Column${col - next.c1 + 1}` : text], taken);
      if (headerRows(def) > 0) setValueAt(ws, next.r1, col, name, ctl.defaultStyleAt(next.r1, col));
      columns.push({ id: nextId++, name });
    }
    def.columns = columns;
    def.ref = rangeAddress(next);
    if (def.autoFilter) def.autoFilter = { ...def.autoFilter, ref: rangeAddress(filterRange(def, next)), filterColumns: [] };
  });
  return undefined;
}

function themeColor(r: ThemeRef) {
  return { theme: r.theme, ...(r.tint !== 0 ? { tint: r.tint } : {}) };
}

function side(line: Line) {
  return { style: line.kind, color: themeColor(line.color) };
}

/**
 * Convert to Range: the table goes, its cells keep the look the style gave
 * them as ordinary formatting, and structured references to it become A1
 * references — what Excel does.
 */
/**
 * Home ▸ Delete ▸ Table Rows: delete the selected data rows inside the table's
 * columns only (a Delete Cells ▸ shift up over the table's band), so the
 * table's ref and AutoFilter shrink with it. A table needs one data row, so
 * deleting all of them keeps the first and empties it.
 */
export function deleteTableRows(ctl: EditorController, rows: Range): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const hit = tableAt(ws, doc.selection.active.row, doc.selection.active.col);
  if (!hit) return;
  const { def, range } = hit;
  const first = range.r1 + headerRows(def);
  const last = range.r2 - totalRows(def);
  const lo = Math.max(rows.r1, first);
  const hi = Math.min(rows.r2, last);
  if (lo > hi) return;
  const all = lo === first && hi === last;
  doc.transact('Delete Table Rows', (tx) => {
    const from = all ? lo + 1 : lo;
    if (hi >= from) structuralEdit(tx, doc.wb, ws, { axis: 'row', at: from, count: -(hi - from + 1), band: { from: range.c1, to: range.c2 } });
    else tx.cells(ws, { ...range, r1: lo, r2: lo });
    if (!all) return;
    for (let c = range.c1; c <= range.c2; c++) {
      const cell = getCellAt(ws, lo, c);
      if (cell) cell.value = null;
    }
  });
}

export function convertToRange(ctl: EditorController, def: TableDefinition): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const range = tableRange(def);
  if (!range) return;
  doc.transact('Convert to Range', (tx) => {
    tx.cells(ws, range);
    tx.sheet(ws, 'tables', 'rowDimensions');
    tx.workbook('definedNames');
    tx.structural = true;
    // Formulas first: their cells are declared one by one, and must be
    // snapshotted before the formatting below changes them.
    rewriteFormulas(tx, doc.wb, (f, sheet, row, col) => structuredToA1Formula(f, def, range, row, sheet === ws && row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2));
    doc.wb.definedNames = doc.wb.definedNames.map((d) => {
      const value = structuredToA1Formula(d.value, def, range, range.r1, false);
      return value === d.value ? d : { ...d, value };
    });
    const style = def.styleInfo?.name;
    if (hasBuiltInStyle(style) && style) {
      const opts = tableOptions(def);
      const geometry = { range, headerRows: headerRows(def), totalRows: totalRows(def) };
      const memo = new Map<string, number>();
      for (let r = range.r1; r <= range.r2; r++) {
        for (let c = range.c1; c <= range.c2; c++) {
          const look = rawLookAt(style, opts, placeOf(geometry, r, c));
          if (!look) continue;
          const cell = getCellAt(ws, r, c) ?? setValueAt(ws, r, c, null, ctl.defaultStyleAt(r, c));
          const key = `${cell.styleId}|${JSON.stringify(look)}`;
          let id = memo.get(key);
          if (id === undefined) {
            const base = doc.styles.get(cell.styleId);
            id = transformStyle(
              doc.wb,
              cell.styleId,
              {
                ...(look.fill && !base.fill ? { fill: { kind: 'pattern', patternType: 'solid', fgColor: themeColor(look.fill) } } : {}),
                ...(look.font || look.bold ? { font: { ...(look.font && !base.ownColor ? { color: themeColor(look.font) } : {}), ...(look.bold ? { bold: true } : {}) } } : {}),
                border: (cur) => ({
                  ...cur,
                  ...(!cur.top?.style && look.top ? { top: side(look.top) } : {}),
                  ...(!cur.bottom?.style && look.bottom ? { bottom: side(look.bottom) } : {}),
                  ...(!cur.left?.style && look.left ? { left: side(look.left) } : {}),
                  ...(!cur.right?.style && look.right ? { right: side(look.right) } : {}),
                }),
              },
              { top: false, bottom: false, left: false, right: false },
            );
            memo.set(key, id);
          }
          cell.styleId = id;
        }
      }
    }
    // Filtered-out rows come back, as the filter goes with the table.
    for (let r = range.r1; r <= range.r2; r++) {
      const dim = ws.rowDimensions.get(r);
      if (!dim?.hidden || !def.autoFilter) continue;
      const { hidden: _h, ...rest } = dim;
      if (Object.keys(rest).length === 0) ws.rowDimensions.delete(r);
      else ws.rowDimensions.set(r, rest);
    }
    ws.tables = ws.tables.filter((t) => t !== def);
  });
}

/** Cells of the table's body rows, for commands such as Remove Duplicates that act on the data. */
export function tableDataRange(def: TableDefinition, range: Range): Range {
  return { ...range, r1: range.r1 + headerRows(def), r2: range.r2 - totalRows(def) };
}
