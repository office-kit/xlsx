// Format Painter: pick up the formatting of the selection, then paint it onto
// the next selection the user makes (or every selection until Escape when
// double-clicked, Excel's "sticky" mode).

import type { Range } from './address.ts';
import { getCellAt, usedRange } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import { currentRange } from './selection.ts';

class FormatPainter {
  active = $state(false);
  sticky = false;
  #source: { styles: number[][]; rows: number; cols: number } | null = null;

  toggle(ctl: EditorController, sticky: boolean): void {
    if (this.active && !sticky) {
      this.active = false;
      return;
    }
    const r = currentRange(ctl.doc.selection);
    const rows = Math.min(r.r2 - r.r1 + 1, 500);
    const cols = Math.min(r.c2 - r.c1 + 1, 200);
    const styles: number[][] = [];
    for (let i = 0; i < rows; i++) {
      const line: number[] = [];
      for (let j = 0; j < cols; j++) line.push(getCellAt(ctl.doc.ws, r.r1 + i, r.c1 + j)?.styleId ?? ctl.defaultStyleAt(r.r1 + i, r.c1 + j));
      styles.push(line);
    }
    this.#source = { styles, rows, cols };
    this.active = true;
    this.sticky = sticky;
  }

  /** Called when a selection gesture ends. */
  apply(ctl: EditorController, target: Range): void {
    const src = this.#source;
    if (!this.active || !src) return;
    const ws = ctl.doc.ws;
    // Whole rows/columns only need painting as far as the sheet is used.
    const used = usedRange(ws);
    const area: Range = {
      r1: target.r1,
      c1: target.c1,
      r2: target.r1 === target.r2 ? target.r1 + src.rows - 1 : Math.min(target.r2, Math.max(used?.r2 ?? 1, target.r1 + src.rows - 1)),
      c2: target.c1 === target.c2 ? target.c1 + src.cols - 1 : Math.min(target.c2, Math.max(used?.c2 ?? 1, target.c1 + src.cols - 1)),
    };
    ctl.doc.transact('Format Painter', (tx) => {
      tx.cells(ws, area);
      for (let r = area.r1; r <= area.r2; r++) {
        for (let c = area.c1; c <= area.c2; c++) {
          const styleId = src.styles[(r - area.r1) % src.rows]?.[(c - area.c1) % src.cols] ?? 0;
          const cell = getCellAt(ws, r, c);
          if (cell) cell.styleId = styleId;
          else if (styleId !== 0) {
            let rowMap = ws.rows.get(r);
            if (!rowMap) {
              rowMap = new Map();
              ws.rows.set(r, rowMap);
            }
            rowMap.set(c, { row: r, col: c, value: null, styleId });
          }
        }
      }
    });
    if (!this.sticky) this.active = false;
  }

  cancel(): void {
    this.active = false;
  }
}

export const formatPainter = new FormatPainter();
