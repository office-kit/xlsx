// Copy / cut / paste.
//
// The system clipboard carries TSV and an HTML table, so data moves to and
// from Excel, Google Sheets and other apps. Inside the editor a richer
// payload (values, formulas, styles, merges) is kept in memory and recognised
// on paste by an id embedded in the HTML, so an internal paste keeps
// formatting and re-anchors formulas the way Excel does.

import type { CellValue } from '@office-kit/xlsx/cell';
import { makeCell } from '@office-kit/xlsx/cell';
import { cellStyleToCss, getCellDisplayText } from '@office-kit/xlsx/styles';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { mergeCells, setColumnDimension, unmergeCells } from '@office-kit/xlsx/worksheet';
import { adjustFormulaForMove, translateFormula } from '../calc/index.ts';
import type { CellPos, Range } from './address.ts';
import { MAX_COL, MAX_ROW, rangesIntersect, toBoundaries } from './address.ts';
import { deleteCellsInRange, forEachCellInRange, getCellAt, usedRange } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import { isRowFiltered } from './filter.ts';
import { parseInput } from './input.ts';
import { pxToColWidth } from './metrics.ts';
import { selectRange } from './selection.ts';

interface ClipCell {
  readonly dr: number;
  readonly dc: number;
  /** Row offset in the source when filtered-out rows were left out of the copy (`dr` is then the packed row). */
  readonly srcDr?: number;
  readonly value: CellValue;
  readonly styleId: number;
}

export interface ClipPayload {
  readonly id: string;
  /** The sheet itself, not its index: sheets can move or go between the cut and the paste. */
  readonly sheet: Worksheet;
  readonly source: Range;
  readonly rows: number;
  readonly cols: number;
  readonly cells: readonly ClipCell[];
  readonly merges: readonly Range[];
  /** Source column widths in px at 100% zoom, by offset. */
  readonly colWidths: ReadonlyMap<number, number>;
  readonly cut: boolean;
}

let payload: ClipPayload | null = null;

const MARKER = 'office-kit-xlsx-clip';

/** Copy/cut of a whole column copies only up to the used range, like Excel's clipboard text. */
function boundedSource(ctl: EditorController, range: Range): Range {
  const used = usedRange(ctl.doc.ws);
  if (!used) return { r1: range.r1, c1: range.c1, r2: range.r1, c2: range.c1 };
  return {
    r1: range.r1,
    c1: range.c1,
    r2: range.r2 === MAX_ROW ? Math.max(range.r1, Math.min(range.r2, used.r2)) : range.r2,
    c2: range.c2 === MAX_COL ? Math.max(range.c1, Math.min(range.c2, used.c2)) : range.c2,
  };
}

export function captureSelection(ctl: EditorController, cut: boolean): ClipPayload | null {
  const p = snapshotSelection(ctl, cut);
  if (!p) return null;
  payload = p;
  ctl.clipboard = { sheet: ctl.doc.ws, range: ctl.doc.selection.ranges[0] ?? p.source, cut };
  return p;
}

function snapshotSelection(ctl: EditorController, cut: boolean): ClipPayload | null {
  const sel = ctl.doc.selection;
  if (sel.ranges.length !== 1) {
    ctl.toast = 'multiSelectionCopy';
    return null;
  }
  const full = sel.ranges[0];
  if (!full) return null;
  const range = boundedSource(ctl, full);
  const ws = ctl.doc.ws;
  // Copying a filtered list copies only the rows on show (Excel); a cut keeps the block whole.
  const packed: number[] = [];
  let skipped = false;
  if (!cut && range.r2 - range.r1 < MAX_ROW - 1) {
    for (let r = range.r1; r <= range.r2; r++) {
      if (isRowFiltered(ws, r)) skipped = true;
      else packed.push(r);
    }
  }
  const packedIndex = skipped ? new Map(packed.map((r, i) => [r, i])) : undefined;
  const cells: ClipCell[] = [];
  forEachCellInRange(ws, range, (c) => {
    const dr = c.row - range.r1;
    if (!packedIndex) {
      cells.push({ dr, dc: c.col - range.c1, value: c.value, styleId: c.styleId });
      return;
    }
    const at = packedIndex.get(c.row);
    if (at !== undefined) cells.push({ dr: at, dc: c.col - range.c1, srcDr: dr, value: c.value, styleId: c.styleId });
  });
  // Only merges wholly inside the block travel with it; one sticking out cannot be rebuilt at the destination.
  const merges = packedIndex
    ? []
    : ctl.doc.merges
        .intersecting(range)
        .filter((m) => m.r1 >= range.r1 && m.c1 >= range.c1 && m.r2 <= range.r2 && m.c2 <= range.c2)
        .map((m) => ({ r1: m.r1 - range.r1, c1: m.c1 - range.c1, r2: m.r2 - range.r1, c2: m.c2 - range.c1 }));
  const colWidths = new Map<number, number>();
  for (let c = range.c1; c <= range.c2 && c - range.c1 < 256; c++) colWidths.set(c - range.c1, ctl.doc.cols.sizeOf(c));
  return {
    id: Math.random().toString(36).slice(2),
    sheet: ctl.doc.ws,
    source: range,
    rows: packedIndex ? packed.length : range.r2 - range.r1 + 1,
    cols: range.c2 - range.c1 + 1,
    cells,
    merges,
    colWidths,
    cut,
  };
}

/**
 * Drag the selection's border to `to` (its new top-left): move the cells, or
 * copy them with Option/Ctrl held. Goes through the same path as Cut/Copy +
 * Paste, so formulas pointing at moved cells follow them.
 */
export function dragSelection(ctl: EditorController, to: CellPos, copy: boolean): void {
  const p = snapshotSelection(ctl, !copy);
  if (!p) return;
  if (to.row === p.source.r1 && to.col === p.source.c1) return;
  ctl.doc.setSelection(selectRange({ r1: to.row, c1: to.col, r2: to.row, c2: to.col }));
  pasteInternal(ctl, p, specialFor('all'));
}

function escapeHtml(s: string): string {
  return s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
}

/** TSV + HTML renderings of a payload, for the system clipboard. */
export function serialise(ctl: EditorController, p: ClipPayload): { text: string; html: string } {
  const ws = ctl.doc.ws;
  const wb = ctl.doc.wb;
  const lines: string[] = [];
  const htmlRows: string[] = [];
  for (let r = 0; r < p.rows; r++) {
    const cols: string[] = [];
    const tds: string[] = [];
    for (let c = 0; c < p.cols; c++) {
      const cell = getCellAt(ws, p.source.r1 + r, p.source.c1 + c);
      const text = cell ? getCellDisplayText(wb, cell) : '';
      cols.push(/[\t\n"]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text);
      const css = cell ? Object.entries(cellStyleToCss(wb, cell)).map(([k, v]) => `${k}:${v}`).join(';') : '';
      tds.push(`<td${css ? ` style="${escapeHtml(css)}"` : ''}>${escapeHtml(text).replaceAll('\n', '<br>')}</td>`);
    }
    lines.push(cols.join('\t'));
    htmlRows.push(`<tr>${tds.join('')}</tr>`);
  }
  const html = `<meta name="${MARKER}" content="${p.id}"><table>${htmlRows.join('')}</table>`;
  return { text: lines.join('\n') + '\n', html };
}

export function writeCopyEvent(ctl: EditorController, ev: ClipboardEvent, cut: boolean): void {
  const p = captureSelection(ctl, cut);
  if (!p || !ev.clipboardData) return;
  const { text, html } = serialise(ctl, p);
  ev.clipboardData.setData('text/plain', text);
  ev.clipboardData.setData('text/html', html);
}

/** Ribbon Copy/Cut buttons have no clipboard event; use the async API. */
export async function copyToSystem(ctl: EditorController, cut: boolean): Promise<void> {
  const p = captureSelection(ctl, cut);
  if (!p) return;
  const { text, html } = serialise(ctl, p);
  try {
    await navigator.clipboard.write([
      new ClipboardItem({ 'text/plain': new Blob([text], { type: 'text/plain' }), 'text/html': new Blob([html], { type: 'text/html' }) }),
    ]);
  } catch {
    // Without clipboard permission the internal payload still serves in-app pastes.
  }
}

export async function pasteFromSystem(ctl: EditorController, mode: PasteMode = 'all'): Promise<void> {
  let html = '';
  let text = '';
  try {
    const items = await navigator.clipboard.read();
    for (const item of items) {
      if (item.types.includes('text/html')) html = await (await item.getType('text/html')).text();
      if (item.types.includes('text/plain')) text = await (await item.getType('text/plain')).text();
    }
  } catch {
    // Permission denied: fall back to the internal clipboard.
    if (payload) html = `<meta name="${MARKER}" content="${payload.id}">`;
  }
  pasteData(ctl, html, text, mode);
}

export function pasteFromEvent(ctl: EditorController, ev: ClipboardEvent): void {
  const html = ev.clipboardData?.getData('text/html') ?? '';
  const text = ev.clipboardData?.getData('text/plain') ?? '';
  pasteData(ctl, html, text, 'all');
}

export type PasteMode = 'all' | 'values' | 'formulas' | 'formats' | 'noBorders' | 'transpose' | 'valuesAndFormats' | 'columnWidths';

export interface PasteSpecialOptions {
  readonly what: 'all' | 'formulas' | 'values' | 'formats' | 'comments' | 'validation' | 'allExceptBorders' | 'columnWidths' | 'formulasAndNumberFormats' | 'valuesAndNumberFormats';
  readonly operation: 'none' | 'add' | 'subtract' | 'multiply' | 'divide';
  readonly skipBlanks: boolean;
  readonly transpose: boolean;
}

function pasteData(ctl: EditorController, html: string, text: string, mode: PasteMode): void {
  const id = new RegExp(`<meta name="${MARKER}" content="([^"]+)"`).exec(html)?.[1];
  if (id && payload && payload.id === id) {
    pasteInternal(ctl, payload, specialFor(mode));
    return;
  }
  const grid = html.includes('<table') ? parseHtmlTable(html) : parseTsv(text);
  if (grid.length === 0) return;
  pasteGrid(ctl, grid);
}

function specialFor(mode: PasteMode): PasteSpecialOptions {
  const base = { operation: 'none' as const, skipBlanks: false, transpose: false };
  switch (mode) {
    case 'values':
      return { ...base, what: 'values' };
    case 'formulas':
      return { ...base, what: 'formulas' };
    case 'formats':
      return { ...base, what: 'formats' };
    case 'noBorders':
      return { ...base, what: 'allExceptBorders' };
    case 'transpose':
      return { ...base, what: 'all', transpose: true };
    case 'valuesAndFormats':
      return { ...base, what: 'valuesAndNumberFormats' };
    case 'columnWidths':
      return { ...base, what: 'columnWidths' };
    default:
      return { ...base, what: 'all' };
  }
}

export function pasteSpecial(ctl: EditorController, opts: PasteSpecialOptions): void {
  if (payload) pasteInternal(ctl, payload, opts);
}

/** Destination block(s): Excel tiles the source when the selection is an exact multiple of it. */
function destination(ctl: EditorController, rows: number, cols: number): { origin: CellPos; tilesR: number; tilesC: number } {
  const sel = ctl.doc.selection;
  const r = sel.ranges[sel.activeRange] ?? sel.ranges[0];
  const origin = r ? { row: r.r1, col: r.c1 } : sel.active;
  if (!r) return { origin, tilesR: 1, tilesC: 1 };
  const h = r.r2 - r.r1 + 1;
  const w = r.c2 - r.c1 + 1;
  const tilesR = h % rows === 0 && h > rows && h !== MAX_ROW ? h / rows : 1;
  const tilesC = w % cols === 0 && w > cols && w !== MAX_COL ? w / cols : 1;
  return { origin, tilesR, tilesC };
}

const OP_SYMBOLS = { add: '+', subtract: '-', multiply: '*', divide: '/' } as const;
const NUMERIC_TEXT = /^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\s*$/;

/** A number, or text that reads as one ('5): Paste Special operations convert it. */
function asNumber(v: CellValue): number | undefined {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && NUMERIC_TEXT.test(v)) return Number(v);
  return undefined;
}

function formulaText(v: CellValue): string | undefined {
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' && v.formula ? v.formula : undefined;
}

/**
 * Paste Special ▸ Add / Subtract / Multiply / Divide, as Excel does it: numbers
 * and numeric text combine into a number, a formula on either side becomes
 * `=(target)+source` / `=target+(source)`, and other text, booleans and errors
 * in the paste area are left as they are.
 */
function numericOp(op: PasteSpecialOptions['operation'], target: CellValue, source: CellValue): CellValue {
  if (op === 'none') return source;
  const sym = OP_SYMBOLS[op];
  const targetFormula = formulaText(target);
  const sourceFormula = formulaText(source);
  const a = asNumber(target);
  const b = asNumber(source);
  if (targetFormula !== undefined || (sourceFormula !== undefined && target !== null)) {
    const left = targetFormula !== undefined ? `(${targetFormula})` : a !== undefined ? String(a) : undefined;
    const right = sourceFormula !== undefined ? `(${sourceFormula})` : b !== undefined ? String(b) : undefined;
    if (left === undefined || right === undefined) return target;
    return { kind: 'formula', t: 'normal', formula: `${left}${sym}${right}` };
  }
  if (a === undefined) return target;
  if (b === undefined) return source;
  switch (op) {
    case 'add':
      return a + b;
    case 'subtract':
      return a - b;
    case 'multiply':
      return a * b;
    case 'divide':
      return b === 0 ? { kind: 'error', code: '#DIV/0!' } : a / b;
  }
}

function pasteInternal(ctl: EditorController, p: ClipPayload, opts: PasteSpecialOptions): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const rows = opts.transpose ? p.cols : p.rows;
  const cols = opts.transpose ? p.rows : p.cols;
  const { origin, tilesR, tilesC } = destination(ctl, rows, cols);
  const dest: Range = {
    r1: origin.row,
    c1: origin.col,
    r2: Math.min(MAX_ROW, origin.row + rows * tilesR - 1),
    c2: Math.min(MAX_COL, origin.col + cols * tilesC - 1),
  };
  const sourceSheet = doc.wb.sheets.some((s) => s.sheet === p.sheet) ? p.sheet : undefined;
  const sourceTitle = sourceSheet?.title ?? ws.title;
  const isMove = p.cut && opts.what === 'all' && !opts.transpose;
  // Excel refuses a paste that would cover part of a merged cell.
  const straddles = (m: { minRow: number; minCol: number; maxRow: number; maxCol: number }) =>
    rangesIntersect({ r1: m.minRow, c1: m.minCol, r2: m.maxRow, c2: m.maxCol }, dest) &&
    !(m.minRow >= dest.r1 && m.maxRow <= dest.r2 && m.minCol >= dest.c1 && m.maxCol <= dest.c2);
  if (ws.mergedCells.some(straddles)) {
    ctl.toast = 'mergedCellConflict';
    return;
  }

  doc.transact(p.cut ? 'Cut and Paste' : 'Paste', (tx) => {
    tx.cells(ws, dest);
    tx.sheet(ws, 'mergedCells');
    if (opts.what === 'columnWidths') tx.sheet(ws, 'columnDimensions');
    const moving = isMove && sourceSheet !== undefined;
    if (moving) {
      // A move rewrites references to the moved block everywhere.
      for (const ref of doc.wb.sheets) {
        if (ref.kind !== 'worksheet') continue;
        tx.cells(ref.sheet, { r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL });
      }
    }

    // Dissolve merges overlapping the destination.
    ws.mergedCells = ws.mergedCells.filter((m) => !rangesIntersect({ r1: m.minRow, c1: m.minCol, r2: m.maxRow, c2: m.maxCol }, dest));

    const byPos = new Map<string, ClipCell>();
    for (const c of p.cells) byPos.set(`${c.dr},${c.dc}`, c);

    if (opts.what === 'columnWidths') {
      for (let tc = 0; tc < tilesC; tc++) {
        for (let c = 0; c < cols; c++) {
          const px = p.colWidths.get(c);
          const col = origin.col + tc * cols + c;
          if (px !== undefined && col <= MAX_COL) setColumnDimension(ws, col, { width: pxToColWidth(px), customWidth: true });
        }
      }
    } else if (moving && sourceSheet) {
      // Cut: clear the source, then write cells unchanged (no relative shift).
      const srcCells = p.cells;
      if (sourceSheet !== ws) tx.cells(sourceSheet, p.source);
      // The block's merges move with it (they are re-created at the destination below).
      tx.sheet(sourceSheet, 'mergedCells');
      for (const m of sourceSheet.mergedCells.slice()) {
        if (m.minRow >= p.source.r1 && m.maxRow <= p.source.r2 && m.minCol >= p.source.c1 && m.maxCol <= p.source.c2) unmergeCells(sourceSheet, m);
      }
      deleteCellsInRange(sourceSheet, p.source);
      deleteCellsInRange(ws, dest);
      for (const c of srcCells) {
        const row = origin.row + c.dr;
        const col = origin.col + c.dc;
        if (row > MAX_ROW || col > MAX_COL) continue;
        putCell(ws, row, col, c.value, c.styleId);
      }
      const dRow = origin.row - p.source.r1;
      const dCol = origin.col - p.source.c1;
      for (const ref of doc.wb.sheets) {
        if (ref.kind !== 'worksheet') continue;
        for (const rowMap of ref.sheet.rows.values()) {
          for (const cell of rowMap.values()) {
            const v = cell.value;
            if (v === null || typeof v !== 'object' || v instanceof Date || v.kind !== 'formula' || !v.formula) continue;
            const next = adjustFormulaForMove(v.formula, ref.sheet.title, { sheet: sourceTitle, range: p.source, toSheet: ws.title, dRow, dCol });
            if (next !== v.formula) cell.value = { ...v, formula: next };
          }
        }
      }
      // Names that point into the block follow it, as Excel does.
      const move = { sheet: sourceTitle, range: p.source, toSheet: ws.title, dRow, dCol };
      const names = doc.wb.definedNames.map((dn) => {
        const host = dn.scope === undefined ? sourceTitle : (doc.wb.sheets[dn.scope]?.sheet.title ?? sourceTitle);
        const value = adjustFormulaForMove(dn.value, host, move);
        return value === dn.value ? dn : { ...dn, value };
      });
      if (names.some((dn, i) => dn !== doc.wb.definedNames[i])) {
        // Formulas that use a moved name never changed text, so only a full recalc reaches them.
        tx.structural = true;
        tx.workbook('definedNames');
        doc.wb.definedNames = names;
      }
    } else {
      for (let tr = 0; tr < tilesR; tr++) {
        for (let tc = 0; tc < tilesC; tc++) {
          for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
              const src = opts.transpose ? byPos.get(`${c},${r}`) : byPos.get(`${r},${c}`);
              const row = origin.row + tr * rows + r;
              const col = origin.col + tc * cols + c;
              if (row > MAX_ROW || col > MAX_COL) continue;
              if (opts.skipBlanks && (!src || src.value === null)) continue;
              const srcRow = src?.srcDr ?? (opts.transpose ? c : r);
              writePasted(ctl, row, col, src, opts, {
                dRow: row - (p.source.r1 + srcRow),
                dCol: col - (p.source.c1 + (opts.transpose ? r : c)),
              });
            }
          }
        }
      }
    }

    if (!opts.transpose && (opts.what === 'all' || opts.what === 'formats' || opts.what === 'allExceptBorders')) {
      for (let tr = 0; tr < tilesR; tr++) {
        for (let tc = 0; tc < tilesC; tc++) {
          for (const m of p.merges) {
            const r1 = origin.row + tr * rows + m.r1;
            const c1 = origin.col + tc * cols + m.c1;
            const r2 = origin.row + tr * rows + m.r2;
            const c2 = origin.col + tc * cols + m.c2;
            if (r2 <= MAX_ROW && c2 <= MAX_COL) mergeCells(ws, toBoundaries({ r1, c1, r2, c2 }));
          }
        }
      }
    }
  });
  if (p.cut && p === payload) {
    payload = null;
    ctl.clipboard = null;
  }
  doc.setSelection(selectRange(dest, origin));
}

function putCell(ws: import('@office-kit/xlsx/worksheet').Worksheet, row: number, col: number, value: CellValue, styleId: number): void {
  let rowMap = ws.rows.get(row);
  if (!rowMap) {
    rowMap = new Map();
    ws.rows.set(row, rowMap);
  }
  rowMap.set(col, makeCell(row, col, value, styleId));
}

function writePasted(
  ctl: EditorController,
  row: number,
  col: number,
  src: ClipCell | undefined,
  opts: PasteSpecialOptions,
  shift: { dRow: number; dCol: number },
): void {
  const ws = ctl.doc.ws;
  const existing = getCellAt(ws, row, col);
  const wantsValue = opts.what !== 'formats' && opts.what !== 'comments' && opts.what !== 'validation' && opts.what !== 'columnWidths';
  const wantsStyle = opts.what === 'all' || opts.what === 'formats' || opts.what === 'allExceptBorders';
  let value: CellValue = existing?.value ?? null;
  if (wantsValue) {
    let v: CellValue = src?.value ?? null;
    if (v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula') {
      if (opts.what === 'values' || opts.what === 'valuesAndNumberFormats') {
        v = v.cachedValueType === 'error' ? { kind: 'error', code: String(v.cachedValue) as `#${string}` } : (v.cachedValue ?? null);
      } else if (v.formula) {
        v = { kind: 'formula', t: 'normal', formula: translateFormula(v.formula, shift.dRow, shift.dCol) };
      }
    }
    value = numericOp(opts.operation, existing?.value ?? null, v);
  }
  const styleId = wantsStyle ? (src?.styleId ?? 0) : (existing?.styleId ?? 0);
  if (value === null && styleId === 0) {
    if (existing) {
      ws.rows.get(row)?.delete(col);
      if (ws.rows.get(row)?.size === 0) ws.rows.delete(row);
    }
    return;
  }
  if (existing) {
    existing.value = value;
    existing.styleId = styleId;
  } else putCell(ws, row, col, value, styleId);
}

/** Paste external data (TSV / HTML table) as values, inferring numbers and dates like typing. */
function pasteGrid(ctl: EditorController, grid: string[][]): void {
  const rows = grid.length;
  const cols = Math.max(...grid.map((r) => r.length));
  const { origin } = destination(ctl, rows, cols);
  const ws = ctl.doc.ws;
  const dest: Range = { r1: origin.row, c1: origin.col, r2: Math.min(MAX_ROW, origin.row + rows - 1), c2: Math.min(MAX_COL, origin.col + cols - 1) };
  ctl.doc.transact('Paste', (tx) => {
    tx.cells(ws, dest);
    grid.forEach((line, r) => {
      line.forEach((text, c) => {
        const row = origin.row + r;
        const col = origin.col + c;
        if (row > MAX_ROW || col > MAX_COL) return;
        const parsed = parseInput(text, { dateOrder: ctl.dateOrder(), date1904: ctl.doc.wb.date1904 });
        const existing = getCellAt(ws, row, col);
        if (parsed.value === null) {
          if (existing) existing.value = null;
          return;
        }
        if (existing) existing.value = parsed.value;
        else putCell(ws, row, col, parsed.value, ctl.defaultStyleAt(row, col));
      });
    });
  });
  ctl.doc.setSelection(selectRange(dest, origin));
}

/** RFC 4180-style TSV with quoted fields that may contain tabs and newlines. */
export function parseTsv(text: string): string[][] {
  const out: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  let i = 0;
  const src = text.replace(/\r\n?/g, '\n');
  while (i < src.length) {
    const ch = src[i] ?? '';
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === '') {
      quoted = true;
      i++;
      continue;
    }
    if (ch === '\t') {
      row.push(field);
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      out.push(row);
      row = [];
      field = '';
    } else field += ch;
    i++;
  }
  if (field !== '' || row.length > 0) {
    row.push(field);
    out.push(row);
  }
  return out;
}

export function parseHtmlTable(html: string): string[][] {
  if (typeof DOMParser === 'undefined') return [];
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return [];
  const out: string[][] = [];
  // rowspan/colspan occupy following positions; track cells claimed from above.
  const claimed = new Map<string, true>();
  Array.from(table.rows).forEach((tr, r) => {
    const line: string[] = out[r] ?? [];
    out[r] = line;
    let c = 0;
    for (const td of Array.from(tr.cells)) {
      while (claimed.has(`${r},${c}`)) c++;
      const text = (td.innerText || td.textContent || '').replace(/ /g, ' ');
      line[c] = text.replace(/\n$/, '');
      for (let dr = 0; dr < td.rowSpan; dr++) for (let dc = 0; dc < td.colSpan; dc++) if (dr || dc) claimed.set(`${r + dr},${c + dc}`, true);
      c += td.colSpan;
    }
  });
  return out.map((line) => Array.from(line, (v) => v ?? ''));
}
