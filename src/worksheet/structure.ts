// Whole-row insertion and deletion. Everything that carries a row position
// follows the edit: the cells, row dimensions and every sheet object that
// holds a ref, plus every formula in the workbook that points at the edited
// sheet (cell formulas, defined names, rule and validation formulas, table
// column formulas, sparklines, charts, x14 extension formulas).
//
// The edit is all-or-nothing: every case that would leave a wrong or corrupt
// workbook is rejected, and every formula is rewritten, before the first
// mutation, so a throw leaves the workbook as it was.

import type { Cell, FormulaValue } from '../cell/cell.js';
import { adjustFormulaForRowEdit, adjustRowSpan, type RowEdit } from '../formula/structure.js';
import { translateFormula } from '../formula/translate.js';
import { MAX_ROW, tupleToCoordinate } from '../utils/coordinate.js';
import { OpenXmlError, OpenXmlSchemaError } from '../utils/exceptions.js';
import type { Workbook } from '../workbook/workbook.js';
import { REL_NS, SHEET_DRAWING_NS, SHEET_MAIN_NS } from '../xml/namespaces.js';
import { parseXml } from '../xml/parser.js';
import type { XmlNode } from '../xml/tree.js';
import { type CellRange, type MultiCellRange, parseRange, rangeToString } from './cell-range.js';
import type { RowDimension } from './dimensions.js';
import type { Worksheet } from './worksheet.js';

const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';
const XM_SQREF = `{${XM_NS}}sqref`;
const XM_F = `{${XM_NS}}f`;
const XDR_ROW = `{${SHEET_DRAWING_NS}}row`;
const XDR_ROW_OFF = `{${SHEET_DRAWING_NS}}rowOff`;
const XDR_FROM = `{${SHEET_DRAWING_NS}}from`;
const XDR_TO = `{${SHEET_DRAWING_NS}}to`;
const XDR_TWO_CELL_ANCHOR = `{${SHEET_DRAWING_NS}}twoCellAnchor`;
const XDR_ABSOLUTE_ANCHOR = `{${SHEET_DRAWING_NS}}absoluteAnchor`;
const PIVOT_TABLE_REL = `${REL_NS}/pivotTable`;
const WORKSHEET_SOURCE = `{${SHEET_MAIN_NS}}worksheetSource`;
const PIVOT_CACHE_PART_RE = /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/;

/**
 * Insert `count` empty rows before row `at` of `ws`, moving row `at` and
 * everything below it down — Excel's Insert ▸ Sheet Rows.
 *
 * Everything that carries a position follows: cells, row dimensions, merges,
 * hyperlinks, notes, threaded comments, data validations, conditional
 * formats, tables, the AutoFilter, sort state, protected ranges, ignored
 * errors, cell watches, smart tags, scenarios, selections, page breaks,
 * sparklines, drawing anchors, charts, pivot tables and defined names. Every
 * formula in the workbook that points at `ws` is rewritten: references move
 * with their cells (absolute ones included) and a range that straddles row
 * `at` grows. Whole-column, 3-D and external references are left alone, as
 * are hyperlink `location`s, frozen panes and the scroll position.
 *
 * The inserted rows take the height and cell styles of the row above, as
 * Excel does by default. A merge is never duplicated into them; one that
 * straddles row `at` grows instead.
 *
 * Shared-formula groups the edit reaches are written back as plain formulas.
 *
 * @throws {OpenXmlSchemaError} when `at` / `count` are not whole numbers in
 * range, a populated cell would be pushed past row 1048576, the insertion
 * splits an array formula, data table or pivot table, or the workbook holds
 * something the edit cannot move: a pivot table the model doesn't carry
 * (located on `ws` or reading from it), or OLE objects / form controls on
 * `ws`. The workbook is unchanged when it throws.
 */
export function insertRows(wb: Workbook, ws: Worksheet, at: number, count = 1): void {
  assertRowArgs('insertRows', at, count);
  if (at + count - 1 > MAX_ROW) {
    throw new OpenXmlSchemaError(`insertRows: rows ${at}..${at + count - 1} run past row ${MAX_ROW}`);
  }
  applyRowEdit(wb, ws, { sheet: ws.title, at, count }, 'insertRows');
}

/**
 * Delete rows `at` .. `at + count - 1` of `ws`, moving the rows below up —
 * Excel's Delete ▸ Sheet Rows. Follows the same rules as {@link insertRows}:
 * a range that straddles the deleted rows shrinks, an object lying wholly in
 * them is dropped, and a formula reference to deleted cells becomes `#REF!`.
 *
 * @throws {OpenXmlSchemaError} when `at` / `count` are out of range, or the
 * deletion would cut through an array formula, data table or pivot table,
 * remove a table's header or totals row, all its data rows or the whole
 * table (remove it with `removeTable` first), remove a pivot table's whole
 * source (`removePivotTable` first), remove a scenario's changing cell or a
 * consolidation source, or the workbook holds something the edit cannot move
 * (see {@link insertRows}). The workbook is unchanged when it throws.
 */
export function deleteRows(wb: Workbook, ws: Worksheet, at: number, count = 1): void {
  assertRowArgs('deleteRows', at, count);
  if (at + count - 1 > MAX_ROW) {
    throw new OpenXmlSchemaError(`deleteRows: rows ${at}..${at + count - 1} run past row ${MAX_ROW}`);
  }
  applyRowEdit(wb, ws, { sheet: ws.title, at, count: -count }, 'deleteRows');
}

function assertRowArgs(fn: string, at: number, count: number): void {
  if (!Number.isInteger(at) || at < 1 || at > MAX_ROW) {
    throw new OpenXmlSchemaError(`${fn}: row ${at} out of range [1, ${MAX_ROW}]`);
  }
  if (!Number.isInteger(count) || count < 1) {
    throw new OpenXmlSchemaError(`${fn}: count must be a positive whole number; got ${count}`);
  }
}

// ---- position mapping ----------------------------------------------------------

/** New index of row `row`, or undefined when it was deleted or pushed off the grid. */
function mapRow(row: number, e: RowEdit): number | undefined {
  const span = adjustRowSpan(row, row, e);
  return span?.[0];
}

/** A whole-column range (`A1:A1048576`) is untouched by a row edit, as in formulas. */
function mapRange(r: CellRange, e: RowEdit): CellRange | undefined {
  if (r.minRow === 1 && r.maxRow === MAX_ROW) return r;
  const span = adjustRowSpan(r.minRow, r.maxRow, e);
  if (span === undefined) return undefined;
  if (span[0] === r.minRow && span[1] === r.maxRow) return r;
  return { ...r, minRow: span[0], maxRow: span[1] };
}

/** Map an A1 ref string; the input string itself when nothing moved. */
function mapRef(ref: string, e: RowEdit): string | undefined {
  const r = parseRange(ref);
  const next = mapRange(r, e);
  if (next === undefined) return undefined;
  return next === r ? ref : rangeToString(next);
}

function mapMulti(m: MultiCellRange, e: RowEdit): MultiCellRange | undefined {
  const ranges = m.ranges.flatMap((r) => mapRange(r, e) ?? []);
  return ranges.length === 0 ? undefined : { ranges };
}

/** Map a space-separated sqref string, dropping the ranges that were deleted. */
function mapSqrefText(sqref: string, e: RowEdit): string | undefined {
  const parts = sqref.split(/\s+/).filter((p) => p.length > 0);
  const mapped = parts.flatMap((p) => mapRef(p, e) ?? []);
  return mapped.length === 0 ? undefined : mapped.join(' ');
}

/**
 * A drawing anchor corner (0-based row plus an offset into it). A corner in
 * deleted rows lands at the top of the first row after them, as Excel's does.
 */
function mapMarker(marker: { row: number; rowOff: number }, e: RowEdit): void {
  const row = mapRow(marker.row + 1, e);
  if (row === undefined) {
    marker.row = Math.min(e.at, MAX_ROW) - 1;
    marker.rowOff = 0;
  } else {
    marker.row = row - 1;
  }
}

/** Old index of a row that survived the edit. */
function unmapRow(row: number, e: RowEdit): number {
  if (row < e.at) return row;
  return row - e.count;
}

const rowOf = (ref: string): number => parseRange(ref).minRow;

/** Does the edit cut through `r` without removing all of it? */
function splits(r: CellRange, e: RowEdit): boolean {
  if (e.count > 0) return e.at > r.minRow && e.at <= r.maxRow;
  const last = e.at - e.count - 1;
  const overlaps = e.at <= r.maxRow && last >= r.minRow;
  const covers = e.at <= r.minRow && last >= r.maxRow;
  return overlaps && !covers;
}

function deletesAny(r: CellRange, e: RowEdit): boolean {
  return e.count < 0 && e.at <= r.maxRow && e.at - e.count - 1 >= r.minRow;
}

const sameSheet = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();

// ---- the edit ----------------------------------------------------------------

function applyRowEdit(wb: Workbook, ws: Worksheet, e: RowEdit, fn: string): void {
  if (!wb.sheets.some((s) => s.kind === 'worksheet' && s.sheet === ws)) {
    throw new OpenXmlSchemaError(`${fn}: worksheet "${ws.title}" does not belong to the workbook`);
  }
  validate(wb, ws, e, fn);
  // Rewriting a formula is the one step that can fail on input (a formula
  // the tokenizer rejects), so every rewrite is computed before anything is
  // touched and committed afterwards.
  const commits: Array<() => void> = [];
  planFormulas(wb, ws, e, fn, commits);
  for (const commit of commits) commit();
  shiftCells(ws, e);
  shiftRowDimensions(ws, e);
  shiftSheetObjects(ws, e);
  shiftWorkbookObjects(wb, ws, e);
}

function validate(wb: Workbook, ws: Worksheet, e: RowEdit, fn: string): void {
  const fail = (what: string): never => {
    throw new OpenXmlSchemaError(`${fn}: ${what}`);
  };
  if (ws.oleObjects.length > 0 || ws.controls.length > 0) {
    fail(`"${ws.title}" holds OLE objects or form controls, whose anchors are not modelled`);
  }
  if (e.count > 0) {
    let lastRow = 0;
    for (const [r, rowMap] of ws.rows) if (rowMap.size > 0 && r > lastRow) lastRow = r;
    if (lastRow >= e.at && lastRow + e.count > MAX_ROW) fail(`row ${lastRow} would be pushed past row ${MAX_ROW}`);
  }
  for (const rowMap of ws.rows.values()) {
    for (const cell of rowMap.values()) {
      const v = cell.value;
      if (!isFormula(v) || (v.t !== 'array' && v.t !== 'dataTable') || v.ref === undefined) continue;
      const where = tupleToCoordinate(cell.col, cell.row);
      if (splits(parseRange(v.ref), e)) fail(`the edit cuts through the ${v.t} formula at ${where} (${v.ref})`);
      for (const input of [v.r1, v.r2]) {
        if (input !== undefined && mapRef(input, e) === undefined) fail(`the edit deletes input cell ${input} of the data table at ${where}`);
      }
    }
  }
  for (const t of ws.tables) {
    const r = parseRange(t.ref);
    if (!deletesAny(r, e)) continue;
    const last = e.at - e.count - 1;
    const header = (t.headerRowCount ?? 1) > 0 ? r.minRow : undefined;
    const totals = (t.totalsRowCount ?? 0) > 0 ? r.maxRow : undefined;
    if (e.at <= r.minRow && last >= r.maxRow) fail(`the edit deletes table "${t.displayName}"; remove it with removeTable first`);
    if (header !== undefined && e.at <= header && last >= header) fail(`the edit deletes the header row of table "${t.displayName}"`);
    if (totals !== undefined && e.at <= totals && last >= totals) fail(`the edit deletes the totals row of table "${t.displayName}"`);
    const firstData = header === undefined ? r.minRow : r.minRow + 1;
    const lastData = totals === undefined ? r.maxRow : r.maxRow - 1;
    if (e.at <= firstData && last >= lastData) fail(`the edit deletes every data row of table "${t.displayName}"`);
  }
  for (const scenario of ws.scenarios?.scenarios ?? []) {
    for (const input of scenario.inputCells) {
      if (mapRef(input.ref, e) === undefined) fail(`the edit deletes changing cell ${input.ref} of scenario "${scenario.name}"`);
    }
  }
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    const host = ref.sheet;
    for (const pt of host.pivotTables ?? []) {
      if (sameSheet(pt.source.sheet, ws.title) && mapRef(pt.source.ref, e) === undefined) {
        fail(`the edit deletes the whole source of pivot table "${pt.name}"; remove it with removePivotTable first`);
      }
      if (host !== ws) continue;
      const area = parseRange(pt.renderedRef ?? pt.anchor);
      if (splits(area, e) || deletesAny(area, e)) fail(`the edit cuts through pivot table "${pt.name}"`);
    }
    for (const source of host.dataConsolidate?.dataRefs ?? []) {
      if (source.ref === undefined || !sameSheet(source.sheet ?? host.title, ws.title)) continue;
      if (mapRef(source.ref, e) === undefined) fail(`the edit deletes consolidation source ${source.ref}`);
    }
  }
  if ((ws.relsExtras ?? []).some((rel) => rel.type === PIVOT_TABLE_REL)) {
    fail(`"${ws.title}" holds a pivot table the model does not carry, which cannot be moved`);
  }
  for (const [path, bytes] of wb.passthrough ?? []) {
    if (!PIVOT_CACHE_PART_RE.test(path)) continue;
    const source = findNode(parseXml(bytes), WORKSHEET_SOURCE);
    const sheet = source?.attrs['sheet'];
    if (sheet !== undefined && source?.attrs['ref'] !== undefined && sameSheet(sheet, ws.title)) {
      fail(`a pivot table the model does not carry reads its data from "${ws.title}"`);
    }
  }
}

function findNode(node: XmlNode, name: string): XmlNode | undefined {
  if (node.name === name) return node;
  for (const child of node.children) {
    const hit = findNode(child, name);
    if (hit !== undefined) return hit;
  }
  return undefined;
}

const isFormula = (v: Cell['value']): v is FormulaValue =>
  typeof v === 'object' && v !== null && !(v instanceof Date) && (v as { kind?: unknown }).kind === 'formula';

// ---- formulas ----------------------------------------------------------------

function planFormulas(wb: Workbook, ws: Worksheet, e: RowEdit, fn: string, commits: Array<() => void>): void {
  const rewrite = (formula: string, host: string | undefined, where: string, rebased?: (f: string) => string): string => {
    try {
      return adjustFormulaForRowEdit(rebased ? rebased(formula) : formula, host, e);
    } catch (cause) {
      if (!(cause instanceof OpenXmlError)) throw cause;
      throw new OpenXmlSchemaError(`${fn}: cannot rewrite the formula ${where}`, { cause });
    }
  };
  for (const ref of wb.sheets) {
    if (ref.kind === 'chartsheet') {
      for (const item of ref.sheet.drawing?.items ?? []) {
        if (item.content.kind === 'chart') planChart(item.content.chart, (f) => rewrite(f, undefined, `in a chart on "${ref.sheet.title}"`), commits);
      }
      continue;
    }
    const host = ref.sheet;
    const at = (where: string): string => `${where} on "${host.title}"`;
    planCellFormulas(host, ws, e, (f, cell) => rewrite(f, host.title, at(`at ${tupleToCoordinate(cell.col, cell.row)}`)), commits);
    for (const cf of host.conditionalFormatting) {
      const anchor = host === ws ? rebase(cf.sqref, e) : undefined;
      for (const rule of cf.rules) {
        const formulas = rule.formulas.map((f) => rewrite(f, host.title, at('of a conditional format'), anchor));
        if (formulas.some((f, i) => f !== rule.formulas[i])) commits.push(() => (rule.formulas = formulas));
      }
    }
    for (const dv of host.dataValidations) {
      const anchor = host === ws ? rebase(dv.sqref, e) : undefined;
      const adjust = (f: string): string => rewrite(f, host.title, at('of a data validation'), anchor);
      const f1 = dv.formula1 === undefined ? undefined : adjust(dv.formula1);
      const f2 = dv.formula2 === undefined ? undefined : adjust(dv.formula2);
      if (f1 !== dv.formula1 || f2 !== dv.formula2) {
        commits.push(() => {
          if (f1 !== undefined) dv.formula1 = f1;
          if (f2 !== undefined) dv.formula2 = f2;
        });
      }
    }
    for (const table of host.tables) {
      for (const column of table.columns) {
        const calc = column.calculatedColumnFormula;
        const totals = column.totalsRowFormula;
        const nextCalc = calc === undefined ? undefined : rewrite(calc, host.title, at(`of table "${table.displayName}"`));
        const nextTotals = totals === undefined ? undefined : rewrite(totals, host.title, at(`of table "${table.displayName}"`));
        if (nextCalc !== calc || nextTotals !== totals) {
          commits.push(() => {
            if (nextCalc !== undefined) column.calculatedColumnFormula = nextCalc;
            if (nextTotals !== undefined) column.totalsRowFormula = nextTotals;
          });
        }
      }
    }
    for (const group of host.sparklineGroups ?? []) {
      for (const sparkline of group.sparklines) {
        const next = rewrite(sparkline.formula, host.title, at('of a sparkline'));
        if (next !== sparkline.formula) commits.push(() => (sparkline.formula = next));
      }
      const date = group.dateFormula;
      if (date !== undefined) {
        const next = rewrite(date, host.title, at('of a sparkline group'));
        if (next !== date) commits.push(() => (group.dateFormula = next));
      }
    }
    for (const item of host.drawing?.items ?? []) {
      if (item.content.kind === 'chart') planChart(item.content.chart, (f) => rewrite(f, undefined, at('in a chart')), commits);
    }
    for (const node of [...(host.bodyExtras?.beforeSheetData ?? []), ...(host.bodyExtras?.afterSheetData ?? [])]) {
      forEachNode(node, (n) => {
        if (n.name !== XM_F || n.text === undefined) return;
        const text = n.text;
        const next = rewrite(text, host.title, at('in an extension list'));
        if (next !== text) commits.push(() => (n.text = next));
      });
    }
  }
  wb.definedNames.forEach((dn) => {
    const scopeRef = dn.scope === undefined ? undefined : wb.sheets[dn.scope];
    const next = rewrite(dn.value, scopeRef?.sheet.title, `of defined name "${dn.name}"`);
    if (next !== dn.value) commits.push(() => (dn.value = next));
  });
}

/**
 * A conditional format's or validation's formula is written for the top-left
 * cell of its first range, its relative references reaching every other
 * cell by offset. When the edit deletes that cell (or the whole first range)
 * the range that survives is anchored somewhere else, so the formula is first
 * moved, copy-style, to the old position of the new anchor; the structural
 * rewrite then sees it from a cell that still exists. Returns undefined when
 * the anchor stays put.
 */
function rebase(sqref: MultiCellRange, e: RowEdit): ((formula: string) => string) | undefined {
  const first = sqref.ranges[0];
  const next = mapMulti(sqref, e)?.ranges[0];
  if (first === undefined || next === undefined) return undefined;
  const oldRow = unmapRow(next.minRow, e);
  if (oldRow === first.minRow && next.minCol === first.minCol) return undefined;
  const origin = tupleToCoordinate(first.minCol, first.minRow);
  const dest = tupleToCoordinate(next.minCol, oldRow);
  return (formula) => translateFormula(`=${formula}`, origin, { dest }).slice(1);
}

/**
 * Cell formulas of `host`. A shared-formula group stores each member's text
 * relative to the group's first cell, so once members move by different
 * amounts, or their references do, the group no longer describes them: such
 * a group is written back as plain formulas. On the edited sheet every group
 * is, since its cells move.
 */
function planCellFormulas(
  host: Worksheet,
  ws: Worksheet,
  e: RowEdit,
  rewrite: (formula: string, cell: Cell) => string,
  commits: Array<() => void>,
): void {
  const masters = new Map<number, Cell>();
  const members = new Map<number, Cell[]>();
  const planned = new Map<Cell, string>();
  for (const rowMap of host.rows.values()) {
    for (const cell of rowMap.values()) {
      const v = cell.value;
      if (!isFormula(v)) continue;
      if (v.t === 'shared' && v.si !== undefined) {
        const group = members.get(v.si);
        if (group === undefined) members.set(v.si, [cell]);
        else group.push(cell);
        if (v.ref !== undefined) masters.set(v.si, cell);
        continue;
      }
      if (v.formula.length === 0) continue;
      const next = rewrite(v.formula, cell);
      if (next !== v.formula) planned.set(cell, next);
    }
  }
  for (const [cell, formula] of planned) {
    const v = cell.value as FormulaValue;
    commits.push(() => (cell.value = Object.freeze({ ...v, formula })));
  }
  for (const [si, group] of members) {
    const master = masters.get(si);
    const masterValue = master?.value as FormulaValue | undefined;
    const texts = group.map((cell) => {
      const own = (cell.value as FormulaValue).formula;
      if (own.length > 0 || master === undefined || masterValue === undefined) return own;
      const origin = tupleToCoordinate(master.col, master.row);
      return translateFormula(`=${masterValue.formula}`, origin, { dest: tupleToCoordinate(cell.col, cell.row) }).slice(1);
    });
    const next = group.map((cell, i) => (texts[i] === '' ? '' : rewrite(texts[i] as string, cell)));
    if (host !== ws && next.every((f, i) => f === texts[i])) continue;
    group.forEach((cell, i) => {
      const { ref: _ref, si: _si, ...rest } = cell.value as FormulaValue;
      const formula = next[i] as string;
      // A member whose text cannot be recovered (no master in the sheet) has
      // nothing to be written back as; leave it for the writer to report.
      if (formula === '') return;
      commits.push(() => (cell.value = Object.freeze({ ...rest, t: 'normal' as const, formula })));
    });
  }
}

/**
 * Chart data references: `ref` on the legacy chart model (series values and
 * categories, series and label text, error-bar data), `f` / `txDataRef` on
 * the chartex model. A reference whose cells were all deleted becomes
 * `#REF!`, as in Excel.
 */
function planChart(chart: object, rewrite: (formula: string) => string, commits: Array<() => void>): void {
  const visit = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (const item of node) visit(item);
      return;
    }
    if (typeof node !== 'object' || node === null || ArrayBuffer.isView(node)) return;
    const record = node as Record<string, unknown>;
    for (const [key, value] of Object.entries(record)) {
      if (typeof value === 'string' && (key === 'ref' || key === 'f' || key === 'txDataRef')) {
        const next = rewrite(value);
        if (next !== value) commits.push(() => (record[key] = next));
      } else if (key !== 'attrs') {
        visit(value);
      }
    }
  };
  visit(chart);
}

function forEachNode(node: XmlNode, fn: (n: XmlNode) => void): void {
  fn(node);
  for (const child of node.children) forEachNode(child, fn);
}

// ---- cells and dimensions --------------------------------------------------------

function shiftCells(ws: Worksheet, e: RowEdit): void {
  const next = new Map<number, Map<number, Cell>>();
  for (const [r, rowMap] of ws.rows) {
    const nr = mapRow(r, e);
    if (nr === undefined) continue;
    if (nr !== r) for (const cell of rowMap.values()) cell.row = nr;
    next.set(nr, rowMap);
  }
  for (const rowMap of next.values()) {
    for (const cell of rowMap.values()) {
      const v = cell.value;
      if (!isFormula(v) || (v.t !== 'array' && v.t !== 'dataTable')) continue;
      const ref = v.ref === undefined ? undefined : mapRef(v.ref, e);
      const r1 = v.r1 === undefined ? undefined : mapRef(v.r1, e);
      const r2 = v.r2 === undefined ? undefined : mapRef(v.r2, e);
      if (ref === v.ref && r1 === v.r1 && r2 === v.r2) continue;
      cell.value = Object.freeze({
        ...v,
        ...(ref !== undefined ? { ref } : {}),
        ...(r1 !== undefined ? { r1 } : {}),
        ...(r2 !== undefined ? { r2 } : {}),
      });
    }
  }
  // Inserted rows take the cell styles of the row above (Excel's default
  // "Format Same As Above"); empty cells only, no values.
  const above = e.count > 0 ? next.get(e.at - 1) : undefined;
  if (above !== undefined) {
    for (let k = 0; k < e.count; k++) {
      const row = e.at + k;
      const rowMap = new Map<number, Cell>();
      for (const cell of above.values()) {
        if (cell.styleId !== 0) rowMap.set(cell.col, { row, col: cell.col, value: null, styleId: cell.styleId });
      }
      if (rowMap.size > 0) next.set(row, rowMap);
    }
  }
  ws.rows = next;
  if (ws._appendRowCursor >= e.at) {
    ws._appendRowCursor = mapRow(ws._appendRowCursor, e) ?? e.at - 1;
  }
}

function shiftRowDimensions(ws: Worksheet, e: RowEdit): void {
  const next = new Map<number, RowDimension>();
  for (const [r, dim] of ws.rowDimensions) {
    const nr = mapRow(r, e);
    if (nr !== undefined) next.set(nr, dim);
  }
  const above = e.count > 0 ? ws.rowDimensions.get(e.at - 1) : undefined;
  if (above !== undefined) {
    const { hidden: _hidden, collapsed: _collapsed, ...inherited } = above;
    if (Object.keys(inherited).length > 0) {
      for (let k = 0; k < e.count; k++) next.set(e.at + k, { ...inherited });
    }
  }
  ws.rowDimensions = next;
}

// ---- objects on the edited sheet -------------------------------------------------

function shiftSheetObjects(ws: Worksheet, e: RowEdit): void {
  ws.mergedCells = ws.mergedCells.flatMap((m) => {
    const r = mapRange(m, e);
    if (r === undefined || (r.minRow === r.maxRow && r.minCol === r.maxCol)) return [];
    return [r];
  });
  ws.hyperlinks = ws.hyperlinks.flatMap((h) => {
    const ref = mapRef(h.ref, e);
    return ref === undefined ? [] : [ref === h.ref ? h : { ...h, ref }];
  });
  ws.legacyComments = ws.legacyComments.flatMap((c) => {
    const ref = mapRef(c.ref, e);
    if (ref === undefined) return [];
    if (ref === c.ref) return [c];
    // The note's box keeps its place next to the cell, so it moves as far.
    const delta = rowOf(ref) - rowOf(c.ref);
    const anchor = c.anchor && {
      ...c.anchor,
      fromRow: Math.max(0, c.anchor.fromRow + delta),
      toRow: Math.max(0, c.anchor.toRow + delta),
    };
    return [{ ...c, ref, ...(anchor ? { anchor } : {}) }];
  });
  if (ws.threadedComments) {
    // Replies share their thread's ref, so deleting the cell drops the thread.
    ws.threadedComments = ws.threadedComments.flatMap((c) => {
      const ref = mapRef(c.ref, e);
      return ref === undefined ? [] : [ref === c.ref ? c : { ...c, ref }];
    });
  }
  ws.dataValidations = ws.dataValidations.flatMap((dv) => {
    const sqref = mapMulti(dv.sqref, e);
    return sqref === undefined ? [] : [{ ...dv, sqref }];
  });
  ws.conditionalFormatting = ws.conditionalFormatting.flatMap((cf) => {
    const sqref = mapMulti(cf.sqref, e);
    return sqref === undefined ? [] : [{ ...cf, sqref }];
  });
  ws.protectedRanges = ws.protectedRanges.flatMap((p) => {
    const sqref = mapMulti(p.sqref, e);
    return sqref === undefined ? [] : [{ ...p, sqref }];
  });
  ws.ignoredErrors = ws.ignoredErrors.flatMap((ie) => {
    const sqref = mapMulti(ie.sqref, e);
    return sqref === undefined ? [] : [{ ...ie, sqref }];
  });
  ws.cellWatches = ws.cellWatches.flatMap((w) => {
    const ref = mapRef(w.ref, e);
    return ref === undefined ? [] : [{ ...w, ref }];
  });
  ws.smartTags = ws.smartTags.flatMap((s) => {
    const ref = mapRef(s.ref, e);
    return ref === undefined ? [] : [{ ...s, ref }];
  });
  // validate() rejected a deleted changing cell, so every one maps.
  if (ws.scenarios) {
    for (const scenario of ws.scenarios.scenarios) {
      for (const input of scenario.inputCells) input.ref = mapRef(input.ref, e) ?? input.ref;
    }
    if (ws.scenarios.sqref) {
      const sqref = mapMulti(ws.scenarios.sqref, e);
      if (sqref) ws.scenarios.sqref = sqref;
      else delete ws.scenarios.sqref;
    }
  }
  ws.tables = ws.tables.map((t) => {
    // validate() rejected every deletion that would remove a table outright.
    const ref = mapRef(t.ref, e) ?? t.ref;
    const autoFilter = t.autoFilter && mapAutoFilter(t.autoFilter, e);
    const sortState = t.sortState && mapSortState(t.sortState, e);
    const { autoFilter: _a, sortState: _s, ...rest } = t;
    return { ...rest, ref, ...(autoFilter ? { autoFilter } : {}), ...(sortState ? { sortState } : {}) };
  });
  if (ws.autoFilter) {
    const autoFilter = mapAutoFilter(ws.autoFilter, e);
    if (autoFilter) ws.autoFilter = autoFilter;
    else delete ws.autoFilter;
  }
  if (ws.sortState) {
    const sortState = mapSortState(ws.sortState, e);
    if (sortState) ws.sortState = sortState;
    else delete ws.sortState;
  }
  ws.views = ws.views.map((v) => (v.selections ? { ...v, selections: v.selections.map((s) => mapSelection(s, e)) } : v));
  ws.rowBreaks = mapBreaks(ws.rowBreaks, e);
  ws.customSheetViews = ws.customSheetViews.map((v) => {
    const next = { ...v };
    if (v.selections) next.selections = v.selections.map((s) => mapSelection(s, e));
    if (v.rowBreaks) next.rowBreaks = mapBreaks(v.rowBreaks, e);
    return next;
  });
  ws.webPublishItems = ws.webPublishItems.map((item) => {
    if (item.sourceRef === undefined) return item;
    const sourceRef = mapRef(item.sourceRef, e);
    return sourceRef === undefined || sourceRef === item.sourceRef ? item : { ...item, sourceRef };
  });
  if (ws.drawing) {
    for (const item of ws.drawing.items) {
      const a = item.anchor;
      if (a.kind === 'oneCell') mapMarker(a.from, e);
      // `editAs="absolute"` is "don't move or size with cells".
      if (a.kind === 'twoCell' && a.editAs !== 'absolute') {
        mapMarker(a.from, e);
        mapMarker(a.to, e);
      }
      if (item.raw) shiftRawAnchors(item.raw, e);
    }
    for (const { node } of ws.drawing.rawChildren ?? []) shiftRawAnchors(node, e);
  }
  if (ws.bodyExtras) {
    ws.bodyExtras.beforeSheetData = shiftExtSqrefs(ws.bodyExtras.beforeSheetData, e);
    ws.bodyExtras.afterSheetData = shiftExtSqrefs(ws.bodyExtras.afterSheetData, e);
  }
}

function mapAutoFilter<T extends { ref: string; sortState?: import('./sort-state.js').SortState }>(f: T, e: RowEdit): T | undefined {
  const ref = mapRef(f.ref, e);
  if (ref === undefined) return undefined;
  const sortState = f.sortState && mapSortState(f.sortState, e);
  const { sortState: _s, ...rest } = f;
  return { ...rest, ref, ...(sortState ? { sortState } : {}) } as T;
}

function mapSortState(s: import('./sort-state.js').SortState, e: RowEdit): import('./sort-state.js').SortState | undefined {
  const ref = mapRef(s.ref, e);
  if (ref === undefined) return undefined;
  const conditions = s.conditions.flatMap((c) => {
    const cref = mapRef(c.ref, e);
    return cref === undefined ? [] : [{ ...c, ref: cref }];
  });
  return { ...s, ref, conditions };
}

function mapSelection(s: import('./views.js').Selection, e: RowEdit): import('./views.js').Selection {
  const next = { ...s };
  // A selection on deleted cells moves to the row that took their place, as Excel's does.
  const clamp = (ref: string): string => {
    const r = parseRange(ref);
    return mapRef(ref, e) ?? rangeToString({ ...r, minRow: Math.min(e.at, MAX_ROW), maxRow: Math.min(e.at, MAX_ROW) });
  };
  if (s.activeCell !== undefined) next.activeCell = clamp(s.activeCell);
  if (s.sqref !== undefined) next.sqref = mapSqrefText(s.sqref, e) ?? next.activeCell ?? clamp(s.sqref.split(/\s+/)[0] as string);
  return next;
}

/** A row break sits below row `id`; it moves with that row and goes with it. */
function mapBreaks(breaks: import('./page-setup.js').PageBreak[], e: RowEdit): import('./page-setup.js').PageBreak[] {
  return breaks.flatMap((b) => {
    if (b.id === undefined) return [b];
    const id = mapRow(b.id, e);
    return id === undefined ? [] : [id === b.id ? b : { ...b, id }];
  });
}

/** The anchors of a drawing element kept verbatim (a group, an AlternateContent wrapper). */
function shiftRawAnchors(node: XmlNode, e: RowEdit): void {
  if (node.name === XDR_ABSOLUTE_ANCHOR) return;
  if (node.name === XDR_TWO_CELL_ANCHOR && node.attrs['editAs'] === 'absolute') return;
  if (node.name === XDR_FROM || node.name === XDR_TO) {
    const row = node.children.find((c) => c.name === XDR_ROW);
    const rowOff = node.children.find((c) => c.name === XDR_ROW_OFF);
    if (row?.text === undefined || !/^\d+$/.test(row.text)) return;
    const marker = { row: Number(row.text), rowOff: Number(rowOff?.text ?? '0') };
    mapMarker(marker, e);
    row.text = String(marker.row);
    if (rowOff !== undefined) rowOff.text = String(marker.rowOff);
    return;
  }
  for (const child of node.children) shiftRawAnchors(child, e);
}

/**
 * x14 extensions (conditional formats with data bars, data validations
 * pointing at other sheets, …) carry their ranges in `xm:sqref`. An entry
 * whose range was deleted goes, and so does a container it leaves empty.
 */
function shiftExtSqrefs(nodes: XmlNode[], e: RowEdit): XmlNode[] {
  const prune = (node: XmlNode): XmlNode | undefined => {
    const sqref = node.children.find((c) => c.name === XM_SQREF);
    if (sqref?.text !== undefined) {
      const text = mapSqrefText(sqref.text, e);
      if (text === undefined) return undefined;
      sqref.text = text;
    }
    if (node.children.length === 0) return node;
    const children = node.children.flatMap((c) => prune(c) ?? []);
    if (children.length === 0) return undefined;
    if (children.length !== node.children.length) {
      node.children = children;
      if (node.attrs['count'] !== undefined) node.attrs['count'] = String(children.length);
    }
    return node;
  };
  return nodes.flatMap((n) => prune(n) ?? []);
}

// ---- objects elsewhere that point at the edited sheet ---------------------------

function shiftWorkbookObjects(wb: Workbook, ws: Worksheet, e: RowEdit): void {
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    const host = ref.sheet;
    if (host.sparklineGroups) {
      host.sparklineGroups = host.sparklineGroups.flatMap((g) => {
        if (host !== ws) return [g];
        const sparklines = g.sparklines.flatMap((s) => {
          const location = mapRef(s.location, e);
          return location === undefined ? [] : [{ ...s, location }];
        });
        return sparklines.length === 0 ? [] : [{ ...g, sparklines }];
      });
    }
    if (host.pivotTables) {
      // validate() rejected a deleted source and any cut through a report.
      host.pivotTables = host.pivotTables.map((pt) => {
        const source = sameSheet(pt.source.sheet, ws.title) ? { ...pt.source, ref: mapRef(pt.source.ref, e) ?? pt.source.ref } : pt.source;
        if (host !== ws) return { ...pt, source };
        const anchor = mapRef(pt.anchor, e) ?? pt.anchor;
        const renderedRef = pt.renderedRef === undefined ? undefined : (mapRef(pt.renderedRef, e) ?? pt.renderedRef);
        return { ...pt, source, anchor, ...(renderedRef !== undefined ? { renderedRef } : {}) };
      });
    }
    for (const source of host.dataConsolidate?.dataRefs ?? []) {
      if (source.ref !== undefined && sameSheet(source.sheet ?? host.title, ws.title)) {
        source.ref = mapRef(source.ref, e) ?? source.ref;
      }
    }
  }
}
