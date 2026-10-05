// AutoFit row heights / column widths from the displayed text (double-click
// on a header edge, Home ▸ Format ▸ AutoFit). Text is measured with a canvas
// in the cell's own font, so the result tracks what the grid paints.

import type { Cell } from '@office-kit/xlsx/cell';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import { setColumnDimension } from '@office-kit/xlsx/worksheet';
import { wrapLines } from '../grid/paint.ts';
import type { EditorController } from './controller.svelte.ts';
import { pxToColWidth, pxToPt } from './metrics.ts';
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

/** Double-click on a row edge: tallest wrapped/multi-line content (or font size) in the row. */
export function autofitRows(ctl: EditorController, rows: readonly number[]): void {
  const ctx = measurer();
  if (!ctx) return;
  const ws = ctl.doc.ws;
  const heights: number[] = [];
  for (const r of rows) {
    let h = 0;
    for (const cell of ws.rows.get(r)?.values() ?? []) {
      if (cell.value === null) continue;
      const style = ctl.doc.styles.get(cell.styleId);
      ctx.font = canvasFont(style, 1);
      const text = getCellDisplayText(ctl.doc.wb, cell);
      const width = ctl.doc.cols.sizeOf(cell.col) - 6;
      const lines = style.wrap ? wrapLines(text, Math.max(1, width), (s) => ctx.measureText(s).width).length : text.split('\n').length;
      h = Math.max(h, lines * style.fontPx * 1.25 + 4);
    }
    heights.push(h);
  }
  ctl.doc.transact('AutoFit Row Height', (tx) => {
    tx.sheet(ws, 'rowDimensions');
    rows.forEach((r, i) => {
      const h = heights[i] ?? 0;
      const dim = ws.rowDimensions.get(r);
      if (h === 0) {
        if (dim) {
          delete dim.height;
          delete dim.customHeight;
        }
        return;
      }
      ws.rowDimensions.set(r, { ...dim, height: pxToPt(Math.ceil(h)) });
    });
  });
}
