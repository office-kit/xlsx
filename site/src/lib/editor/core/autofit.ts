// AutoFit row heights / column widths from the displayed text (double-click
// on a header edge, Home ▸ Format ▸ AutoFit), and the automatic row fit
// Excel applies after every edit. Text is measured with a canvas
// in the cell's own font, so the result tracks what the grid paints.

import type { Cell } from '@office-kit/xlsx/cell';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import { setColumnDimension } from '@office-kit/xlsx/worksheet';
import { wrapLines } from '../grid/paint.ts';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import type { AxisIndex } from './axis.ts';
import type { EditorController } from './controller.svelte.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';
import type { Transaction } from './history.ts';
import { MergeIndex } from './merges.ts';
import { buildColumnAxis, defaultRowPx, PX_PER_PT, pxToColWidth, pxToPt } from './metrics.ts';
import { canvasFont } from './render-style.ts';

let measureCtx: CanvasRenderingContext2D | null = null;

function measurer(): CanvasRenderingContext2D | null {
  if (!measureCtx && typeof document !== 'undefined') measureCtx = document.createElement('canvas').getContext('2d');
  return measureCtx;
}

/** Pixel width a cell's displayed text needs at 100% zoom, padding included. */
export function cellTextWidth(ctl: EditorController, cell: Cell): number | undefined {
  const ctx = measurer();
  if (!ctx) return undefined;
  const style = ctl.doc.styles.get(cell.styleId);
  ctx.font = canvasFont(style, 1);
  const text = getCellDisplayText(ctl.doc.wb, cell);
  return Math.max(...text.split('\n').map((line) => ctx.measureText(line).width)) + 8 + style.indent * 9;
}

/** Double-click on a column edge: width of the widest displayed text in the column. */
export function autofitColumns(ctl: EditorController, cols: readonly number[]): void {
  const ws = ctl.doc.ws;
  const widths = new Map<number, number>();
  for (const rowMap of ws.rows.values()) {
    for (const c of cols) {
      const cell = rowMap.get(c);
      if (!cell || cell.value === null || ctl.doc.merges.at(cell.row, cell.col)) continue;
      if (ctl.doc.styles.get(cell.styleId).wrap) continue;
      const w = cellTextWidth(ctl, cell) ?? 0;
      widths.set(c, Math.max(widths.get(c) ?? 0, w));
    }
  }
  if (widths.size === 0) return;
  ctl.doc.transact('AutoFit Column Width', (tx) => {
    tx.sheet(ws, 'columnDimensions');
    for (const [c, px] of widths) setColumnDimension(ws, c, { width: pxToColWidth(Math.ceil(px)), customWidth: true, bestFit: true });
  });
}

/**
 * Height in px a row's content needs at 100% zoom, or 0 when it has none.
 * A line is as tall as the default row is for the Normal font, scaled by the
 * cell's font size (Excel's measured heights are proportional), so
 * Normal-font rows keep the sheet's default height; a wrapped or multi-line cell needs one line per displayed line. Cells merged
 * across rows do not count, as in Excel.
 */
function rowContentPx(doc: SpreadsheetEditor, ws: Worksheet, row: number, merges: MergeIndex, cols: () => AxisIndex): number {
  const linePx = defaultRowPx(ws) / doc.styles.get(0).fontPx;
  let h = 0;
  for (const cell of ws.rows.get(row)?.values() ?? []) {
    const merge = merges.at(row, cell.col);
    if (merge && merge.r1 !== merge.r2) continue;
    const style = doc.styles.get(cell.styleId);
    let fontPx = style.fontPx;
    // A rich-text run in a larger size makes its line taller.
    if (cell.value !== null && typeof cell.value === 'object' && 'kind' in cell.value && cell.value.kind === 'rich-text') {
      for (const run of cell.value.runs) if (run.font?.sz !== undefined) fontPx = Math.max(fontPx, run.font.sz * PX_PER_PT);
    }
    // Excel shows an unwrapped cell on one line even when its text has line breaks.
    let lines = 1;
    const ctx = style.wrap && cell.value !== null ? measurer() : null;
    if (ctx) {
      ctx.font = canvasFont(style, 1);
      lines = wrapLines(getCellDisplayText(doc.wb, cell), Math.max(1, cols().sizeOf(cell.col) - 6), (t) => ctx.measureText(t).width).length;
    }
    h = Math.max(h, lines * fontPx * linePx);
  }
  return Math.round(h);
}

/** Set each row's height to its content's, or back to the default when the content fits in that. */
function fitRows(doc: SpreadsheetEditor, tx: Transaction, ws: Worksheet, rows: Iterable<number>): void {
  const merges = new MergeIndex(ws);
  let axis: AxisIndex | undefined;
  const cols = () => (axis ??= buildColumnAxis(ws));
  const defaultPx = defaultRowPx(ws);
  let declared = false;
  for (const r of rows) {
    const px = rowContentPx(doc, ws, r, merges, cols);
    const dim = ws.rowDimensions.get(r);
    // Excel never fits a row below the default height, however small the font.
    const target = px <= defaultPx ? undefined : pxToPt(px);
    if (dim?.height === target) continue;
    if (!declared) {
      tx.sheet(ws, 'rowDimensions');
      declared = true;
    }
    if (target === undefined) {
      if (!dim) continue;
      const { height: _h, customHeight: _c, ...rest } = dim;
      if (Object.keys(rest).length === 0) ws.rowDimensions.delete(r);
      else ws.rowDimensions.set(r, rest);
    } else ws.rowDimensions.set(r, { ...dim, height: target });
  }
}

/**
 * Excel keeps rows whose height was never set by hand fitted to their
 * content: a larger font, wrapped text or a pasted block grows the row, and
 * clearing it shrinks the row back. Runs inside the edit's transaction so the
 * height change undoes with it.
 */
export function fitAutoRows(doc: SpreadsheetEditor, tx: Transaction): void {
  if (tx.structural) return;
  const byWs = new Map<Worksheet, Set<number>>();
  for (const part of tx.parts) {
    if (part.kind !== 'cells') continue;
    const { ws, range } = part;
    let rows = byWs.get(ws);
    if (!rows) byWs.set(ws, (rows = new Set()));
    // Whole-column ranges reach every row, so walk only the rows that exist.
    if (range.r2 - range.r1 + 1 > ws.rows.size) {
      for (const r of ws.rows.keys()) if (r >= range.r1 && r <= range.r2) rows.add(r);
    } else for (let r = range.r1; r <= range.r2; r++) rows.add(r);
  }
  for (const [ws, rows] of byWs) {
    const auto = [...rows].filter((r) => {
      const dim = ws.rowDimensions.get(r);
      return !dim?.customHeight && !dim?.hidden;
    });
    if (auto.length > 0) fitRows(doc, tx, ws, auto);
  }
}

/** Double-click on a row edge: fit each row to its content. */
export function autofitRows(ctl: EditorController, rows: readonly number[]): void {
  const ws = ctl.doc.ws;
  ctl.doc.transact('AutoFit Row Height', (tx) => {
    tx.sheet(ws, 'rowDimensions');
    for (const r of rows) {
      const dim = ws.rowDimensions.get(r);
      if (dim?.customHeight) ws.rowDimensions.set(r, { ...dim, customHeight: false });
    }
    fitRows(ctl.doc, tx, ws, rows);
  });
}
