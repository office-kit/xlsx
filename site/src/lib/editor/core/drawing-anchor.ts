// Drawing anchors ↔ sheet pixels. Anchors are in EMU relative to cell
// markers; the grid works in content pixels at 100% zoom, through the same
// axis indices the painter uses.

import type { DrawingAnchor } from '@office-kit/xlsx/drawing';
import type { Axis } from './axis.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';

/** DrawingML stores positions in EMU; at 96 dpi one pixel is 9525 EMU. */
export const EMU_PER_PX = 9525;

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

function markerPx(axis: Axis, index0: number, offEmu: number): number {
  return axis.offsetOf(index0 + 1) + offEmu / EMU_PER_PX;
}

/** Anchor → sheet-content pixels at 100% zoom. */
export function anchorRect(a: DrawingAnchor, cols: Axis, rows: Axis): Rect {
  switch (a.kind) {
    case 'absolute':
      return { x: a.pos.x / EMU_PER_PX, y: a.pos.y / EMU_PER_PX, w: a.ext.cx / EMU_PER_PX, h: a.ext.cy / EMU_PER_PX };
    case 'oneCell': {
      const x = markerPx(cols, a.from.col, a.from.colOff);
      const y = markerPx(rows, a.from.row, a.from.rowOff);
      return { x, y, w: a.ext.cx / EMU_PER_PX, h: a.ext.cy / EMU_PER_PX };
    }
    case 'twoCell': {
      const x = markerPx(cols, a.from.col, a.from.colOff);
      const y = markerPx(rows, a.from.row, a.from.rowOff);
      return { x, y, w: markerPx(cols, a.to.col, a.to.colOff) - x, h: markerPx(rows, a.to.row, a.to.rowOff) - y };
    }
  }
}

function markerAt(axis: Axis, px: number): { index: number; off: number } {
  const i = axis.indexAt(Math.max(0, px));
  return { index: i - 1, off: Math.max(0, Math.round((px - axis.offsetOf(i)) * EMU_PER_PX)) };
}

/** The same kind of anchor as `a`, moved/resized to `r` (content px). */
export function anchorFor(a: DrawingAnchor, r: Rect, cols: Axis, rows: Axis): DrawingAnchor {
  const ext = { cx: Math.round(r.w * EMU_PER_PX), cy: Math.round(r.h * EMU_PER_PX) };
  if (a.kind === 'absolute') return { kind: 'absolute', pos: { x: Math.round(r.x * EMU_PER_PX), y: Math.round(r.y * EMU_PER_PX) }, ext };
  const c = markerAt(cols, r.x);
  const rw = markerAt(rows, r.y);
  const from = { col: c.index, colOff: c.off, row: rw.index, rowOff: rw.off };
  if (a.kind === 'oneCell') return { kind: 'oneCell', from, ext };
  const c2 = markerAt(cols, r.x + r.w);
  const r2 = markerAt(rows, r.y + r.h);
  return { ...a, from, to: { col: c2.index, colOff: c2.off, row: r2.index, rowOff: r2.off } };
}

/** The active sheet's drawing item at `index` in content pixels, or undefined when there is none. */
export function drawingRect(doc: SpreadsheetEditor, index: number): Rect | undefined {
  const item = doc.ws.drawing?.items[index];
  return item ? anchorRect(item.anchor, doc.cols, doc.rows) : undefined;
}

/** Format ▸ Size: resize the item at `index` from its top-left corner. One undo step. */
export function resizeDrawing(doc: SpreadsheetEditor, index: number, width: number, height: number): void {
  const ws = doc.ws;
  const item = ws.drawing?.items[index];
  if (!item) return;
  const r = anchorRect(item.anchor, doc.cols, doc.rows);
  const next = { ...r, w: Math.max(1, width), h: Math.max(1, height) };
  doc.transact('Resize Object', (tx) => {
    tx.sheet(ws, 'drawing');
    const live = ws.drawing?.items[index];
    if (live) live.anchor = anchorFor(live.anchor, next, doc.cols, doc.rows);
  });
}
