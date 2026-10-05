// Editing operations shared by the keyboard, ribbon, menus and dialogs. Each
// runs as one undoable transaction on the editor and declares exactly the
// cells or fields it touches.

import type { Cell, CellValue } from '@office-kit/xlsx/cell';
import { makeCell } from '@office-kit/xlsx/cell';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import {
  getColumnDimension,
  getFreezePanes,
  getRowDimension,
  mergeCells,
  setColumnDimension,
  setFreezePanes,
  setRowDimension,
  unmergeCells,
} from '@office-kit/xlsx/worksheet';
import type { CellPos, Range } from './address.ts';
import { MAX_COL, MAX_ROW, parseRangeAddress, rangesIntersect, toBoundaries } from './address.ts';
import { deleteCellsInRange, forEachCellInRange, getCellAt } from './cells.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';
import type { Transaction } from './history.ts';
import { pxToColWidth } from './metrics.ts';
import { applyStyle, transformStyle, type StylePatch } from './format.ts';
import { isRowFiltered } from './filter.ts';
import { parseInput, type DateOrder } from './input.ts';
import { selectRange } from './selection.ts';

export interface InputOptions {
  readonly dateOrder: DateOrder;
  /**
   * Width in px the cell's displayed text needs. When given, a number or date
   * that would show as #### widens a column whose width was never set — what
   * Excel does as you type.
   */
  readonly measure?: (cell: Cell) => number | undefined;
}

function defaultStyleAt(ws: Worksheet, row: number, col: number): number {
  return getRowDimension(ws, row)?.style ?? getColumnDimension(ws, col)?.style ?? 0;
}

/**
 * Commit typed text into `targets` (one cell, or every cell of the selection
 * for Ctrl/Cmd+Enter). Formulas are re-anchored per target by `translate`,
 * which the formula engine provides.
 */
const ALL_EDGES = { top: true, bottom: true, left: true, right: true };

export function commitInput(
  editor: SpreadsheetEditor,
  at: CellPos,
  text: string,
  opts: InputOptions,
  fill?: { ranges: readonly Range[]; translate: (formula: string, dRow: number, dCol: number) => string },
): void {
  const ws = editor.ws;
  const wb = editor.wb;
  const parsed = parseInput(text, { dateOrder: opts.dateOrder, date1904: wb.date1904, sheetTitles: wb.sheets.map((s) => s.sheet.title) });
  const targets: Range[] = fill ? fill.ranges.slice() : [{ r1: at.row, c1: at.col, r2: at.row, c2: at.col }];
  editor.transact('Typing', (tx) => {
    for (const range of targets) {
      tx.cells(ws, range);
      for (let r = range.r1; r <= range.r2; r++) {
        // Ctrl+Enter over a filtered list fills only the rows on show.
        if (fill && isRowFiltered(ws, r)) continue;
        for (let c = range.c1; c <= range.c2; c++) {
          let value: CellValue = parsed.value;
          if (fill && value !== null && typeof value === 'object' && !(value instanceof Date) && value.kind === 'formula') {
            const formula = fill.translate(value.formula, r - at.row, c - at.col);
            value = { kind: 'formula', t: 'normal', formula };
          }
          writeValue(editor, ws, r, c, value, parsed.impliedFormat);
          // Typing a line break (Alt+Enter) turns on Wrap Text, as in Excel.
          if (typeof value === 'string' && value.includes('\n')) {
            const cell = getCellAt(ws, r, c);
            if (cell && !editor.styles.get(cell.styleId).wrap) cell.styleId = transformStyle(wb, cell.styleId, { alignment: { wrapText: true } }, ALL_EDGES);
          }
        }
      }
    }
    if (!fill && opts.measure && typeof parsed.value === 'number') widenToFit(editor, tx, at, opts.measure);
  });
}

function widenToFit(editor: SpreadsheetEditor, tx: Transaction, at: CellPos, measure: (cell: Cell) => number | undefined): void {
  const ws = editor.ws;
  const cell = getCellAt(ws, at.row, at.col);
  const dim = getColumnDimension(ws, at.col);
  if (!cell || dim?.customWidth || editor.merges.at(at.row, at.col)) return;
  const need = measure(cell);
  if (need === undefined || need <= editor.cols.sizeOf(at.col)) return;
  tx.sheet(ws, 'columnDimensions');
  setColumnDimension(ws, at.col, { width: pxToColWidth(Math.ceil(need)) });
}

export function writeValue(editor: SpreadsheetEditor, ws: Worksheet, row: number, col: number, value: CellValue, impliedFormat: string | undefined): void {
  const wb = editor.wb;
  let cell = getCellAt(ws, row, col);
  if (value === null) {
    if (cell) cell.value = null;
    return;
  }
  if (!cell) {
    cell = makeCell(row, col, null, defaultStyleAt(ws, row, col));
    let rowMap = ws.rows.get(row);
    if (!rowMap) {
      rowMap = new Map();
      ws.rows.set(row, rowMap);
    }
    rowMap.set(col, cell);
  }
  cell.value = value;
  if (impliedFormat) {
    // Only a General cell picks up the format implied by the typed text.
    const current = editor.styles.get(cell.styleId).numFmt;
    if (current === 'General') {
      cell.styleId = transformStyle(wb, cell.styleId, { numFmt: impliedFormat }, ALL_EDGES);
    }
  }
}

export type ClearKind = 'contents' | 'formats' | 'all' | 'comments' | 'hyperlinks' | 'removeHyperlinks';

export function clearRanges(editor: SpreadsheetEditor, ranges: readonly Range[], kind: ClearKind): void {
  const ws = editor.ws;
  editor.transact(kind === 'contents' ? 'Clear Contents' : 'Clear', (tx) => {
    for (const range of ranges) {
      tx.cells(ws, range);
      if (kind === 'all') {
        tx.sheet(ws, 'hyperlinks', 'legacyComments', 'threadedComments');
        deleteCellsInRange(ws, range);
        removeAnnotations(ws, range, true, true);
        continue;
      }
      if (kind === 'removeHyperlinks') {
        // Unlike Clear Hyperlinks, Remove Hyperlinks also drops the blue
        // underlined link style; only the linked cells are reset.
        tx.sheet(ws, 'hyperlinks');
        for (const link of ws.hyperlinks) {
          const at = parseRangeAddress(link.ref)?.range;
          if (!at || !rangesIntersect(at, range)) continue;
          const overlap = { r1: Math.max(at.r1, range.r1), c1: Math.max(at.c1, range.c1), r2: Math.min(at.r2, range.r2), c2: Math.min(at.c2, range.c2) };
          forEachCellInRange(ws, overlap, (cell) => {
            cell.styleId = 0;
          });
        }
        removeAnnotations(ws, range, false, true);
        continue;
      }
      if (kind === 'comments' || kind === 'hyperlinks') {
        tx.sheet(ws, 'hyperlinks', 'legacyComments', 'threadedComments');
        removeAnnotations(ws, range, kind === 'comments', kind === 'hyperlinks');
        continue;
      }
      const empty: Array<{ row: number; col: number }> = [];
      forEachCellInRange(ws, range, (cell) => {
        if (kind === 'contents') cell.value = null;
        else cell.styleId = 0;
        if (cell.value === null && cell.styleId === 0 && cell.hyperlinkId === undefined && cell.commentId === undefined) empty.push(cell);
      });
      for (const { row, col } of empty) {
        const rowMap = ws.rows.get(row);
        rowMap?.delete(col);
        if (rowMap?.size === 0) ws.rows.delete(row);
      }
    }
  });
}

function removeAnnotations(ws: Worksheet, range: Range, comments: boolean, links: boolean): void {
  const inside = (ref: string): boolean => {
    const m = /^([A-Z]+)(\d+)/.exec(ref.replaceAll('$', ''));
    if (!m?.[1] || !m[2]) return false;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
    const row = Number(m[2]);
    return row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2;
  };
  if (comments) {
    // Excel's Clear Comments and Notes removes both kinds.
    ws.legacyComments = ws.legacyComments.filter((c) => !inside(c.ref));
    ws.threadedComments = (ws.threadedComments ?? []).filter((c) => !inside(c.ref));
  }
  if (links) ws.hyperlinks = ws.hyperlinks.filter((h) => !inside(h.ref));
}

export function formatRanges(editor: SpreadsheetEditor, ranges: readonly Range[], patch: StylePatch, label = 'Format Cells'): void {
  const ws = editor.ws;
  editor.transact(label, (tx) => {
    for (const range of ranges) {
      tx.cells(ws, range);
      if (range.r1 === 1 && range.r2 === MAX_ROW) tx.sheet(ws, 'columnDimensions');
      if (range.c1 === 1 && range.c2 === MAX_COL) tx.sheet(ws, 'rowDimensions');
      applyStyle(editor.wb, ws, range, patch);
    }
  });
}

// ---- merge ------------------------------------------------------------------

export type MergeMode = 'mergeCenter' | 'mergeAcross' | 'merge' | 'unmerge';

function mergeParts(range: Range, mode: MergeMode): Range[] {
  return mode === 'mergeAcross' ? Array.from({ length: range.r2 - range.r1 + 1 }, (_, i) => ({ ...range, r1: range.r1 + i, r2: range.r1 + i })) : [range];
}

/** Whether merging would discard a value: some merged block holds more than one. */
export function mergeDiscardsValues(ws: Worksheet, ranges: readonly Range[], mode: MergeMode): boolean {
  if (mode === 'unmerge') return false;
  return ranges.some((range) =>
    mergeParts(range, mode).some((part) => {
      let filled = 0;
      forEachCellInRange(ws, part, (cell) => {
        if (cell.value !== null && cell.value !== '') filled++;
      });
      return filled > 1;
    }),
  );
}

export function mergeRanges(editor: SpreadsheetEditor, ranges: readonly Range[], mode: MergeMode): void {
  const ws = editor.ws;
  editor.transact(mode === 'unmerge' ? 'Unmerge Cells' : 'Merge Cells', (tx) => {
    tx.sheet(ws, 'mergedCells');
    for (const range of ranges) {
      tx.cells(ws, range);
      // Any merge overlapping the target is dissolved first, as Excel does.
      for (const existing of ws.mergedCells.slice()) {
        const r = { r1: existing.minRow, c1: existing.minCol, r2: existing.maxRow, c2: existing.maxCol };
        if (rangesIntersect(r, range)) unmergeCells(ws, existing);
      }
      if (mode === 'unmerge') continue;
      for (const part of mergeParts(range, mode)) {
        if (part.r1 === part.r2 && part.c1 === part.c2) continue;
        // Excel keeps the upper-left-most value, moved into the anchor, and discards the rest.
        let first: CellValue = null;
        forEachCellInRange(ws, part, (cell) => {
          if (first === null && cell.value !== null) first = cell.value;
          cell.value = null;
        });
        const anchor = getCellAt(ws, part.r1, part.c1);
        if (first !== null) {
          if (anchor) anchor.value = first;
          else writeValue(editor, ws, part.r1, part.c1, first, undefined);
        }
        mergeCells(ws, toBoundaries(part));
      }
      if (mode === 'mergeCenter') applyStyle(editor.wb, ws, { r1: range.r1, c1: range.c1, r2: range.r1, c2: range.c1 }, { alignment: { horizontal: 'center' } });
    }
  });
  if (mode !== 'unmerge') {
    const r = ranges[0];
    if (r) editor.setSelection(selectRange(r));
  }
}

// ---- rows & columns ---------------------------------------------------------

export function setColumnWidths(editor: SpreadsheetEditor, cols: readonly number[], width: number | undefined): void {
  const ws = editor.ws;
  editor.transact('Column Width', (tx) => {
    tx.sheet(ws, 'columnDimensions');
    for (const c of cols) {
      const dim = getColumnDimension(ws, c);
      if (width === undefined) setColumnDimension(ws, c, { ...(dim?.style !== undefined ? { style: dim.style } : {}) });
      else setColumnDimension(ws, c, { width, customWidth: true, hidden: width === 0 });
    }
  });
}

export function setRowHeights(editor: SpreadsheetEditor, rows: readonly number[], height: number | undefined): void {
  const ws = editor.ws;
  editor.transact('Row Height', (tx) => {
    tx.sheet(ws, 'rowDimensions');
    for (const r of rows) {
      if (height === undefined) {
        const dim = getRowDimension(ws, r);
        if (dim) {
          delete dim.height;
          delete dim.customHeight;
        }
      } else setRowDimension(ws, r, { height, customHeight: true, hidden: height === 0 });
    }
  });
}

export function setHidden(editor: SpreadsheetEditor, axis: 'row' | 'col', indices: readonly number[], hidden: boolean): void {
  const ws = editor.ws;
  editor.transact(hidden ? 'Hide' : 'Unhide', (tx) => {
    tx.sheet(ws, axis === 'row' ? 'rowDimensions' : 'columnDimensions');
    for (const i of indices) {
      if (axis === 'row') {
        const dim = setRowDimension(ws, i, { hidden });
        // Unhiding a row that was hidden by giving it zero height restores the default height.
        if (!hidden && dim.height === 0) delete dim.height;
      } else {
        const dim = setColumnDimension(ws, i, { hidden });
        if (!hidden && dim.width === 0) delete dim.width;
      }
    }
  });
}

// ---- view -------------------------------------------------------------------

export function frozenCounts(ws: Worksheet): { rows: number; cols: number } {
  const ref = getFreezePanes(ws);
  if (!ref) return { rows: 0, cols: 0 };
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m?.[1] || !m[2]) return { rows: 0, cols: 0 };
  let col = 0;
  for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
  return { rows: Number(m[2]) - 1, cols: col - 1 };
}

export function freeze(editor: SpreadsheetEditor, rows: number, cols: number): void {
  const ws = editor.ws;
  editor.transact(rows + cols === 0 ? 'Unfreeze Panes' : 'Freeze Panes', (tx) => {
    tx.sheet(ws, 'views');
    setFreezePanes(ws, rows + cols === 0 ? undefined : { rows, cols });
  });
}
