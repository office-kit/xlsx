// SheetView / Pane / Selection. (`Worksheet.views`) + the openpyxl reference at
// `worksheet/views.py`.
//
// **Stage 1**: SheetView with the most-used field subset (tabSelected, view,
// workbookViewId, showGridLines, zoomScale, topLeftCell, pane, selection).
// Reader / writer cover the round-trip; `setFreezePanes` builds the pane from
// an "A1"-style top-left ref. Per-pane multi-selection blocks aren't widespread
// in real-world fixtures, so stage-1 stores a single Selection.

import { coordinateToTuple, tupleToCoordinate } from '../utils/coordinate.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';

export type PaneType = 'bottomRight' | 'topRight' | 'bottomLeft' | 'topLeft';
export type PaneState = 'split' | 'frozen' | 'frozenSplit';
export type SheetViewMode = 'normal' | 'pageBreakPreview' | 'pageLayout';

export interface Pane {
  /** Column-axis split. Number of columns frozen on the left. */
  xSplit?: number;
  /** Row-axis split. Number of rows frozen on top. */
  ySplit?: number;
  /** First visible cell of the bottom-right pane. */
  topLeftCell?: string;
  /** Which pane is active when the sheet is opened. Defaults to `bottomRight` for full freezes. */
  activePane?: PaneType;
  state: PaneState;
}

export interface Selection {
  pane?: PaneType;
  activeCell?: string;
  sqref?: string;
}

export interface SheetView {
  /** Index into the parent workbook's bookViews list. Defaults to 0. */
  workbookViewId: number;
  tabSelected?: boolean;
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  showFormulas?: boolean;
  showZeros?: boolean;
  rightToLeft?: boolean;
  /** Outline (grouping) symbols beside the headers; Excel's Ctrl+8 / Cmd+8 toggle. */
  showOutlineSymbols?: boolean;
  view?: SheetViewMode;
  topLeftCell?: string;
  zoomScale?: number;
  zoomScaleNormal?: number;
  pane?: Pane;
  selection?: Selection;
}

/** Build a SheetView with sensible defaults. */
export function makeSheetView(opts: Partial<SheetView> = {}): SheetView {
  return {
    workbookViewId: opts.workbookViewId ?? 0,
    ...(opts.tabSelected !== undefined ? { tabSelected: opts.tabSelected } : {}),
    ...(opts.showGridLines !== undefined ? { showGridLines: opts.showGridLines } : {}),
    ...(opts.showRowColHeaders !== undefined ? { showRowColHeaders: opts.showRowColHeaders } : {}),
    ...(opts.showFormulas !== undefined ? { showFormulas: opts.showFormulas } : {}),
    ...(opts.showZeros !== undefined ? { showZeros: opts.showZeros } : {}),
    ...(opts.rightToLeft !== undefined ? { rightToLeft: opts.rightToLeft } : {}),
    ...(opts.showOutlineSymbols !== undefined ? { showOutlineSymbols: opts.showOutlineSymbols } : {}),
    ...(opts.view !== undefined ? { view: opts.view } : {}),
    ...(opts.topLeftCell !== undefined ? { topLeftCell: opts.topLeftCell } : {}),
    ...(opts.zoomScale !== undefined ? { zoomScale: opts.zoomScale } : {}),
    ...(opts.zoomScaleNormal !== undefined ? { zoomScaleNormal: opts.zoomScaleNormal } : {}),
    ...(opts.pane ? { pane: opts.pane } : {}),
    ...(opts.selection ? { selection: opts.selection } : {}),
  };
}

/** How many rows / columns a freeze holds in place, counted from the top-left. */
export interface FreezeCounts {
  rows: number;
  cols: number;
}

/**
 * Build a frozen Pane from a top-left coordinate. Per Excel semantics:
 * - "B2" → freeze 1 row + 1 col → xSplit=1, ySplit=1, activePane='bottomRight'
 * - "A2" → freeze 1 row only → ySplit=1, activePane='bottomLeft'
 * - "B1" → freeze 1 col only → xSplit=1, activePane='topRight'
 * - "A1" → no freeze; throws (caller should clear `ws.views[].pane`).
 */
export function makeFreezePane(topLeftRef: string): Pane {
  const { col, row } = coordinateToTuple(topLeftRef);
  if (col === 1 && row === 1) {
    throw new OpenXmlSchemaError('makeFreezePane: "A1" is not a valid freeze ref (no rows or columns to freeze)');
  }
  const xSplit = col - 1;
  const ySplit = row - 1;
  let activePane: PaneType;
  if (xSplit > 0 && ySplit > 0) activePane = 'bottomRight';
  else if (ySplit > 0) activePane = 'bottomLeft';
  else activePane = 'topRight';
  const pane: Pane = {
    state: 'frozen',
    topLeftCell: topLeftRef,
    activePane,
  };
  if (xSplit > 0) pane.xSplit = xSplit;
  if (ySplit > 0) pane.ySplit = ySplit;
  return pane;
}

/** Inverse of {@link makeFreezePane}. Returns the top-left ref of the bottomRight pane, or undefined. */
export function freezePaneRef(view: SheetView): string | undefined {
  const pane = view.pane;
  if (!pane || pane.state !== 'frozen') return undefined;
  if (pane.topLeftCell) return pane.topLeftCell;
  const xSplit = pane.xSplit ?? 0;
  const ySplit = pane.ySplit ?? 0;
  return tupleToCoordinate(xSplit + 1, ySplit + 1);
}
