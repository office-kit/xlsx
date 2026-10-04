// The fill-handle drag.

import type { CellPos, Range } from '../core/address.ts';
import { MAX_COL, MAX_ROW } from '../core/address.ts';
import { autoFill } from '../core/actions.ts';
import type { EditorController } from '../core/controller.svelte.ts';

let target: Range | null = null;

export function startFillDrag(ctl: EditorController): void {
  target = null;
  ctl.dragPreview = null;
}

/** The fill target follows the pointer along whichever axis it left the source on. */
export function previewFill(ctl: EditorController, source: Range, pos: CellPos): void {
  const below = pos.row - source.r2;
  const above = source.r1 - pos.row;
  const right = pos.col - source.c2;
  const left = source.c1 - pos.col;
  const vertical = Math.max(below, above) >= Math.max(right, left);
  let next: Range;
  if (vertical && below > 0) next = { ...source, r2: Math.min(MAX_ROW, pos.row) };
  else if (vertical && above > 0) next = { ...source, r1: Math.max(1, pos.row) };
  else if (!vertical && right > 0) next = { ...source, c2: Math.min(MAX_COL, pos.col) };
  else if (!vertical && left > 0) next = { ...source, c1: Math.max(1, pos.col) };
  else {
    // Dragging back inside the source marks the cells that will be cleared.
    next = source.r2 - source.r1 >= source.c2 - source.c1 ? { ...source, r2: Math.max(source.r1, pos.row) } : { ...source, c2: Math.max(source.c1, pos.col) };
  }
  target = next;
  ctl.dragPreview = next;
}

/** `toggle` (Option/Ctrl held) flips Excel's guess: a lone value counts on, a series copies. */
export function finishFillDrag(ctl: EditorController, source: Range, toggle: boolean): void {
  const t = target;
  target = null;
  ctl.dragPreview = null;
  if (!t) return;
  const single = source.r1 === source.r2 && source.c1 === source.c2;
  const mode = toggle ? (single ? 'series' : 'copy') : 'auto';
  autoFill(ctl, source, t, mode);
  if (t.r2 >= source.r2 && t.c2 >= source.c2 && (t.r2 > source.r2 || t.c2 > source.c2)) {
    ctl.fillOptions = { source, target: t, mode, version: ctl.doc.version };
  }
}
