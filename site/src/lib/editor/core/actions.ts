// User-level operations reachable from the ribbon, menus, keyboard and
// dialogs. They only depend on the controller's public state, so every entry
// point runs exactly the same code.

import type { CellValue } from '@office-kit/xlsx/cell';
import { makeCell } from '@office-kit/xlsx/cell';
import type { Font, HorizontalAlignment, Side, VerticalAlignment } from '@office-kit/xlsx/styles';
import { classifyDateFormat, getCellFont, getCellProtection, isDateFormat, makeColor } from '@office-kit/xlsx/styles';
import { addWorksheet, moveSheet, removeSheet, renameSheet, setSheetState } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import type { PageSetup } from '@office-kit/xlsx/worksheet';
import {
  setColumnDimension,
  setRowDimension,
  setSheetTabColor,
  makeSheetView,
  setSheetViewMode as setViewMode,
} from '@office-kit/xlsx/worksheet';
import { columnIndexFromLetter, columnLetterFromIndex, coordinateFromString } from '@office-kit/xlsx/utils';
import { addImageAt, loadImage } from '@office-kit/xlsx/drawing';
import { deleteSheetInFormula, formulaReferences, fromStorageFormula, renameSheetInFormula, translateFormula } from '../calc/index.ts';
import type { Range } from './address.ts';
import { MAX_COL, MAX_ROW, inRange, quoteSheetName, rangeAddress } from './address.ts';
import { refCell } from './comments.ts';
import { forEachCellInRange, getCellAt, isBlank } from './cells.ts';
import { validateValue } from './validation.ts';
import { flashFill } from './flash-fill.ts';
import { groupLines, isCollapsed, toggleRun } from './outline.ts';
import { clearRanges, formatRanges, freeze, mergeDiscardsValues, mergeRanges, setHidden, type ClearKind, type MergeMode } from './commands.ts';
import { autofitColumns, autofitRows } from './autofit.ts';
import { isRowFiltered } from './filter.ts';
import { copyToSystem, pasteFromSystem, type PasteMode } from './clipboard.ts';
import type { EditorController } from './controller.svelte.ts';
import { extendSeries, type SeriesMode, type SeriesSeed } from './fill.ts';
import { borderPatch, type BorderPreset, type StylePatch } from './format.ts';
import { pxToColWidth, pxToPt } from './metrics.ts';
import { currentRange } from './selection.ts';
import { validateName } from './names.ts';
import { applyStructuralEdit, declareStructural, structuralEdit } from './structure.ts';
import { recentFunctions } from './recent-functions.svelte.ts';

// ---- history ------------------------------------------------------------------

export function undo(ctl: EditorController): void {
  if (ctl.edit) ctl.cancelEdit();
  ctl.doc.undo();
}

export function redo(ctl: EditorController): void {
  if (ctl.doc.canRedo) ctl.doc.redo();
  else ctl.repeatable?.();
}

/** Record `fn` as the action F4 / Cmd+Y repeats, then run it. */
function repeatable(ctl: EditorController, fn: () => void): void {
  ctl.repeatable = fn;
  fn();
}

// ---- clipboard ------------------------------------------------------------------

export function copy(ctl: EditorController): void {
  void copyToSystem(ctl, false);
}

export function cut(ctl: EditorController): void {
  void copyToSystem(ctl, true);
}

export function paste(ctl: EditorController, mode: PasteMode = 'all'): void {
  void pasteFromSystem(ctl, mode);
}

// ---- clearing & formatting --------------------------------------------------------

export function clear(ctl: EditorController, kind: ClearKind): void {
  if ((kind === 'contents' || kind === 'all') && ctl.splitsArray(ctl.doc.selection.ranges)) return;
  clearRanges(ctl.doc, ctl.doc.selection.ranges, kind);
}

export function format(ctl: EditorController, patch: StylePatch, label?: string): void {
  repeatable(ctl, () => formatRanges(ctl.doc, ctl.doc.selection.ranges, patch, label));
}

/** Font of the active cell — what toggle buttons reflect. */
export function activeFont(ctl: EditorController): Font {
  const { row, col } = ctl.doc.selection.active;
  const cell = getCellAt(ctl.doc.ws, row, col) ?? makeCell(row, col, null, ctl.defaultStyleAt(row, col));
  return getCellFont(ctl.doc.wb, cell);
}

export function activeStyle(ctl: EditorController) {
  const { row, col } = ctl.doc.selection.active;
  const cell = getCellAt(ctl.doc.ws, row, col);
  return ctl.doc.styles.get(cell?.styleId ?? ctl.defaultStyleAt(row, col));
}

export function toggleBold(ctl: EditorController): void {
  format(ctl, { font: { bold: !activeFont(ctl).bold } }, 'Bold');
}

export function toggleItalic(ctl: EditorController): void {
  format(ctl, { font: { italic: !activeFont(ctl).italic } }, 'Italic');
}

export function toggleUnderline(ctl: EditorController, kind: 'single' | 'double' = 'single'): void {
  const cur = activeFont(ctl).underline;
  format(ctl, { font: { underline: cur === kind ? undefined : kind } }, 'Underline');
}

export function toggleStrike(ctl: EditorController): void {
  format(ctl, { font: { strike: !activeFont(ctl).strike } }, 'Strikethrough');
}

export function setFontName(ctl: EditorController, name: string): void {
  format(ctl, { font: { name, scheme: undefined } }, 'Font');
}

export function setFontSize(ctl: EditorController, size: number): void {
  format(ctl, { font: { size } }, 'Font Size');
}

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

export function stepFontSize(ctl: EditorController, direction: 1 | -1): void {
  const cur = activeFont(ctl).size ?? 11;
  const next = direction > 0 ? (FONT_SIZES.find((s) => s > cur) ?? cur + 10) : ([...FONT_SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1));
  setFontSize(ctl, next);
}

export function setFontColor(ctl: EditorController, rgb: string | null): void {
  format(ctl, { font: { color: rgb ? makeColor({ rgb }) : makeColor({ theme: 1 }) } }, 'Font Color');
}

export function setFillColor(ctl: EditorController, rgb: string | null): void {
  format(ctl, { fill: rgb ? { kind: 'pattern', patternType: 'solid', fgColor: makeColor({ rgb }) } : null }, 'Fill Color');
}

export function setHAlign(ctl: EditorController, horizontal: HorizontalAlignment | undefined): void {
  const cur = activeStyle(ctl).hAlign;
  format(ctl, { alignment: { horizontal: cur === horizontal ? undefined : horizontal } }, 'Align');
}

export function setVAlign(ctl: EditorController, vertical: VerticalAlignment): void {
  format(ctl, { alignment: { vertical } }, 'Align');
}

export function toggleWrap(ctl: EditorController): void {
  format(ctl, { alignment: { wrapText: !activeStyle(ctl).wrap } }, 'Wrap Text');
}

export function toggleShrink(ctl: EditorController): void {
  format(ctl, { alignment: { shrinkToFit: !activeStyle(ctl).shrink } }, 'Shrink Text to Fit');
}

export function setRotation(ctl: EditorController, textRotation: number): void {
  format(ctl, { alignment: { textRotation } }, 'Orientation');
}

export function stepIndent(ctl: EditorController, direction: 1 | -1): void {
  const s = activeStyle(ctl);
  const indent = Math.max(0, s.indent + direction);
  format(ctl, { alignment: { indent: indent || undefined, horizontal: indent > 0 && s.hAlign === 'general' ? 'left' : s.hAlign === 'general' ? undefined : s.hAlign } }, 'Indent');
}

export function setNumberFormat(ctl: EditorController, code: string): void {
  format(ctl, { numFmt: code }, 'Number Format');
}

/** Increase/Decrease Decimal: adjust the decimals of the active cell's format (General gets the value's own). */
export function stepDecimals(ctl: EditorController, direction: 1 | -1): void {
  const s = activeStyle(ctl);
  let code = s.numFmt;
  if (code === 'General') {
    const { row, col } = ctl.doc.selection.active;
    const v = getCellAt(ctl.doc.ws, row, col)?.value;
    const n = typeof v === 'number' ? v : 0;
    const decimals = (String(n).split('.')[1] ?? '').length;
    code = decimals > 0 ? `0.${'0'.repeat(decimals)}` : '0';
  }
  const next = code.replace(/0(\.0*)?(?![^"]*")/, (m) => {
    const decimals = m.length > 1 ? m.length - 2 : 0;
    const n = Math.max(0, decimals + direction);
    return n === 0 ? '0' : `0.${'0'.repeat(n)}`;
  });
  setNumberFormat(ctl, next);
}

export function applyBorder(ctl: EditorController, preset: BorderPreset, side: Side = { style: 'thin' }): void {
  format(ctl, { border: borderPatch(preset, side) }, 'Borders');
}

export function merge(ctl: EditorController, mode: MergeMode): void {
  const run = () => repeatable(ctl, () => mergeRanges(ctl.doc, ctl.doc.selection.ranges, mode));
  if (mergeDiscardsValues(ctl.doc.ws, ctl.doc.selection.ranges, mode)) {
    ctl.dialog = { kind: 'alert', props: { message: 'mergeDiscardsValues', onConfirm: run } };
    return;
  }
  run();
}

// ---- fill -------------------------------------------------------------------------

/**
 * Fill `target` from `source` (fill-handle drag, Cmd+D / Cmd+R). When the
 * target is inside the source (dragging back), the cells it no longer covers
 * are cleared.
 */
/**
 * AutoFill `target` from `source` (which it contains). `formats` copies only
 * the formatting; `values` fills the series but keeps the target's formats
 * (Excel's "Fill Without Formatting").
 */
export type FillMode = SeriesMode | 'formats' | 'values';

export function autoFill(ctl: EditorController, source: Range, target: Range, mode: FillMode = 'auto'): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  if (target.r1 === source.r1 && target.r2 === source.r2 && target.c1 === source.c1 && target.c2 === source.c2) return;
  const shrinking = target.r2 < source.r2 || target.c2 < source.c2;
  if (shrinking && target.r1 === source.r1 && target.c1 === source.c1) {
    const cleared: Range = target.r2 < source.r2 ? { ...source, r1: target.r2 + 1 } : { ...source, c1: target.c2 + 1 };
    clearRanges(doc, [cleared], 'contents');
    ctl.selectRange(target);
    return;
  }
  const isDate = (styleId: number) => isDateFormat(doc.styles.get(styleId).numFmt);
  const isTime = (styleId: number) => classifyDateFormat(doc.styles.get(styleId).numFmt) === 'time';
  const fillCtx = { translate: translateFormula, isDate, isTime, date1904: doc.wb.date1904 };
  const seriesMode: SeriesMode = mode === 'formats' ? 'copy' : mode === 'values' ? 'auto' : mode;
  doc.transact('AutoFill', (tx) => {
    tx.cells(ws, target);
    const write = (row: number, col: number, produced: SeriesSeed) => {
      // Ctrl+D / Ctrl+R over a filtered list fill only the rows on show.
      if (mode === 'copy' && isRowFiltered(ws, row)) return;
      const existing = getCellAt(ws, row, col);
      const seed: SeriesSeed =
        mode === 'formats'
          ? { value: existing?.value ?? null, styleId: produced.styleId }
          : mode === 'values'
            ? { value: produced.value, styleId: existing?.styleId ?? ctl.defaultStyleAt(row, col) }
            : produced;
      let rowMap = ws.rows.get(row);
      if (seed.value === null && seed.styleId === 0) {
        rowMap?.delete(col);
        return;
      }
      if (!rowMap) {
        rowMap = new Map();
        ws.rows.set(row, rowMap);
      }
      rowMap.set(col, makeCell(row, col, seed.value, seed.styleId));
    };
    const seedAt = (row: number, col: number): SeriesSeed => {
      const c = getCellAt(ws, row, col);
      return { value: c?.value ?? null, styleId: c?.styleId ?? 0 };
    };
    if (target.r2 > source.r2 || target.r1 < source.r1) {
      const down = target.r2 > source.r2;
      const count = down ? target.r2 - source.r2 : source.r1 - target.r1;
      for (let c = source.c1; c <= source.c2; c++) {
        const seeds: SeriesSeed[] = [];
        for (let r = source.r1; r <= source.r2; r++) seeds.push(seedAt(r, c));
        const out = extendSeries(seeds, count, down ? 1 : -1, 'row', fillCtx, seriesMode);
        out.forEach((seed, i) => write(down ? source.r2 + 1 + i : source.r1 - 1 - i, c, seed));
      }
    } else {
      const right = target.c2 > source.c2;
      const count = right ? target.c2 - source.c2 : source.c1 - target.c1;
      for (let r = source.r1; r <= source.r2; r++) {
        const seeds: SeriesSeed[] = [];
        for (let c = source.c1; c <= source.c2; c++) seeds.push(seedAt(r, c));
        const out = extendSeries(seeds, count, right ? 1 : -1, 'col', fillCtx, seriesMode);
        out.forEach((seed, i) => write(r, right ? source.c2 + 1 + i : source.c1 - 1 - i, seed));
      }
    }
  });
  ctl.selectRange(target, { row: source.r1, col: source.c1 });
}

/** Cmd+D / Cmd+R: copy the top row / left column of the selection across it. */
export function fillFrom(ctl: EditorController, direction: 'down' | 'right' | 'up' | 'left'): void {
  repeatable(ctl, () => {
    const r = currentRange(ctl.doc.selection);
    // `autoFill` selects the filled range around the source; keep the cell the
    // user was on active, or the next entry would overwrite the source.
    const active = ctl.doc.selection.active;
    const single = r.r1 === r.r2 && r.c1 === r.c2;
    // A single cell fills from its neighbour (the cell above / to the left).
    if (direction === 'down') {
      const src = single ? { ...r, r1: r.r1 - 1, r2: r.r1 - 1 } : { ...r, r2: r.r1 };
      if (src.r1 < 1) return;
      autoFill(ctl, src, { ...r, r1: src.r1 }, 'copy');
    } else if (direction === 'right') {
      const src = single ? { ...r, c1: r.c1 - 1, c2: r.c1 - 1 } : { ...r, c2: r.c1 };
      if (src.c1 < 1) return;
      autoFill(ctl, src, { ...r, c1: src.c1 }, 'copy');
    } else if (direction === 'up') {
      autoFill(ctl, { ...r, r1: r.r2 }, r, 'copy');
    } else {
      autoFill(ctl, { ...r, c1: r.c2 }, r, 'copy');
    }
    ctl.selectRange(r, active);
  });
}

/** Double-clicking the fill handle fills down as far as the adjacent column has data. */
export function fillHandleDoubleClick(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const src = currentRange(ctl.doc.selection);
  const probe = (col: number): number => {
    let r = src.r2;
    while (r < MAX_ROW && !isBlank(getCellAt(ws, r + 1, col))) r++;
    return r;
  };
  const left = src.c1 > 1 ? probe(src.c1 - 1) : src.r2;
  const right = src.c2 < MAX_COL ? probe(src.c2 + 1) : src.r2;
  const last = Math.max(left, right);
  if (last > src.r2) autoFill(ctl, src, { ...src, r2: last });
}

// ---- rows / columns / cells -----------------------------------------------------------

function selectedLines(ctl: EditorController, axis: 'row' | 'col'): Array<[number, number]> {
  const spans = ctl.doc.selection.ranges.map((r) => (axis === 'row' ? [r.r1, r.r2] : [r.c1, r.c2]) as [number, number]);
  spans.sort((a, b) => a[0] - b[0]);
  const merged: Array<[number, number]> = [];
  for (const s of spans) {
    const last = merged.at(-1);
    if (last && s[0] <= last[1] + 1) last[1] = Math.max(last[1], s[1]);
    else merged.push([s[0], s[1]]);
  }
  return merged;
}

export function insertLines(ctl: EditorController, axis: 'row' | 'col'): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const spans = selectedLines(ctl, axis);
  repeatable(ctl, () =>
    doc.transact(axis === 'row' ? 'Insert Rows' : 'Insert Columns', (tx) => {
      // Snapshot the workbook once, not once per selected span; insert from
      // the bottom/right so earlier indices stay valid.
      declareStructural(tx, doc.wb, ws);
      for (const [lo, hi] of spans.slice().reverse()) applyStructuralEdit(doc.wb, ws, { axis, at: lo, count: hi - lo + 1 });
    }),
  );
}

export function deleteLines(ctl: EditorController, axis: 'row' | 'col'): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const spans = selectedLines(ctl, axis);
  doc.transact(axis === 'row' ? 'Delete Rows' : 'Delete Columns', (tx) => {
    declareStructural(tx, doc.wb, ws);
    for (const [lo, hi] of spans.slice().reverse()) applyStructuralEdit(doc.wb, ws, { axis, at: lo, count: -(hi - lo + 1) });
  });
  const r = currentRange(doc.selection);
  const keep = axis === 'row' ? { ...r, r2: r.r1 } : { ...r, c2: r.c1 };
  ctl.selectRange(axis === 'row' ? { ...keep, c1: r.c1, c2: r.c2 } : keep);
}

/** Insert Cells ▸ shift down / right; Delete Cells ▸ shift up / left. */
export function shiftCells(ctl: EditorController, mode: 'down' | 'right' | 'up' | 'left'): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const r = currentRange(doc.selection);
  doc.transact(mode === 'down' || mode === 'right' ? 'Insert Cells' : 'Delete Cells', (tx) => {
    if (mode === 'down' || mode === 'up') {
      structuralEdit(tx, doc.wb, ws, { axis: 'row', at: r.r1, count: mode === 'down' ? r.r2 - r.r1 + 1 : -(r.r2 - r.r1 + 1), band: { from: r.c1, to: r.c2 } });
    } else {
      structuralEdit(tx, doc.wb, ws, { axis: 'col', at: r.c1, count: mode === 'right' ? r.c2 - r.c1 + 1 : -(r.c2 - r.c1 + 1), band: { from: r.r1, to: r.r2 } });
    }
  });
}

/** Ctrl+Shift+= / Ctrl+-: whole-row/column selections insert/delete lines directly; otherwise ask. */
export function insertCellsSmart(ctl: EditorController): void {
  const r = currentRange(ctl.doc.selection);
  if (r.c1 === 1 && r.c2 === MAX_COL) insertLines(ctl, 'row');
  else if (r.r1 === 1 && r.r2 === MAX_ROW) insertLines(ctl, 'col');
  else ctl.dialog = { kind: 'insertCells' };
}

export function deleteCellsSmart(ctl: EditorController): void {
  const r = currentRange(ctl.doc.selection);
  if (r.c1 === 1 && r.c2 === MAX_COL) deleteLines(ctl, 'row');
  else if (r.r1 === 1 && r.r2 === MAX_ROW) deleteLines(ctl, 'col');
  else ctl.dialog = { kind: 'deleteCells' };
}

export function hideLines(ctl: EditorController, axis: 'row' | 'col', hidden: boolean): void {
  const indices: number[] = [];
  for (const [lo, hi] of selectedLines(ctl, axis)) {
    // Unhiding a whole-sheet selection only touches lines that carry dimensions.
    if (hi - lo > 20_000) {
      const dims = axis === 'row' ? [...ctl.doc.ws.rowDimensions.keys()] : [...ctl.doc.ws.columnDimensions.values()].flatMap((d) => Array.from({ length: d.max - d.min + 1 }, (_, i) => d.min + i));
      for (const i of dims) if (i >= lo && i <= hi) indices.push(i);
      continue;
    }
    for (let i = lo; i <= hi; i++) indices.push(i);
  }
  setHidden(ctl.doc, axis, indices, hidden);
}

export function resizeColumnsPx(ctl: EditorController, cols: readonly number[], px: number): void {
  const ws = ctl.doc.ws;
  ctl.doc.transact('Column Width', (tx) => {
    tx.sheet(ws, 'columnDimensions');
    for (const c of cols) setColumnDimension(ws, c, px === 0 ? { hidden: true } : { width: pxToColWidth(px), customWidth: true, hidden: false });
  });
}

export function resizeRowsPx(ctl: EditorController, rows: readonly number[], px: number): void {
  const ws = ctl.doc.ws;
  ctl.doc.transact('Row Height', (tx) => {
    tx.sheet(ws, 'rowDimensions');
    for (const r of rows) setRowDimension(ws, r, px === 0 ? { hidden: true } : { height: pxToPt(px), customHeight: true, hidden: false });
  });
}

// ---- view ---------------------------------------------------------------------------

export function freezePanes(ctl: EditorController, mode: 'panes' | 'topRow' | 'firstColumn' | 'unfreeze'): void {
  const f = ctl.frozen;
  if (mode === 'unfreeze' || (mode === 'panes' && f.rows + f.cols > 0)) {
    freeze(ctl.doc, 0, 0);
    return;
  }
  if (mode === 'topRow') freeze(ctl.doc, 1, 0);
  else if (mode === 'firstColumn') freeze(ctl.doc, 0, 1);
  else {
    const { row, col } = ctl.doc.selection.active;
    freeze(ctl.doc, row - 1, col - 1);
  }
}

/** View toggles (gridlines, headings, formulas) are view state: not undoable, but saved with the file. */
export function setViewFlag(ctl: EditorController, flag: 'showGridLines' | 'showRowColHeaders' | 'showFormulas' | 'showZeros' | 'showOutlineSymbols', on: boolean): void {
  const ws = ctl.doc.ws;
  const view = ws.views[0] ?? makeSheetView();
  if (ws.views.length === 0) ws.views.push(view);
  view[flag] = on;
  ctl.doc.dirty = true;
  ctl.doc.version++;
}

// ---- sheets --------------------------------------------------------------------------

function uniqueSheetName(ctl: EditorController, base: string): string {
  const names = new Set(ctl.doc.wb.sheets.map((s) => s.sheet.title.toLowerCase()));
  let n = ctl.doc.wb.sheets.length + 1;
  let name = `${base}${n}`;
  while (names.has(name.toLowerCase())) name = `${base}${++n}`;
  return name;
}

export function insertSheet(ctl: EditorController, at?: number): void {
  const doc = ctl.doc;
  const index = at ?? doc.activeSheetIndex + 1;
  doc.transact('Insert Sheet', (tx) => {
    tx.structural = true;
    tx.workbook('sheets', 'activeSheetIndex', 'definedNames');
    addWorksheet(doc.wb, uniqueSheetName(ctl, 'Sheet'), { index });
  });
  doc.activateSheet(index);
}

export function deleteSheet(ctl: EditorController, index: number): void {
  const doc = ctl.doc;
  const visible = doc.wb.sheets.filter((s) => s.state === 'visible' && s.kind === 'worksheet');
  if (visible.length <= 1) {
    ctl.dialog = { kind: 'alert', props: { message: 'lastSheet' } };
    return;
  }
  const title = doc.wb.sheets[index]?.sheet.title;
  if (title === undefined) return;
  const next = Math.max(0, index >= doc.wb.sheets.length - 1 ? index - 1 : index);
  const order = doc.wb.sheets.map((s) => s.sheet.title);
  doc.transact('Delete Sheet', (tx) => {
    tx.structural = true;
    tx.workbook('sheets', 'activeSheetIndex', 'definedNames');
    removeSheet(doc.wb, title);
    // A PivotTable reading from the deleted sheet can no longer be rebuilt; its values stay as plain cells.
    for (const s of doc.wb.sheets) {
      if (s.kind !== 'worksheet' || !s.sheet.pivotTables?.some((pt) => pt.source.sheet.toLowerCase() === title.toLowerCase())) continue;
      tx.sheet(s.sheet, 'pivotTables');
      s.sheet.pivotTables = s.sheet.pivotTables.filter((pt) => pt.source.sheet.toLowerCase() !== title.toLowerCase());
    }
    for (const s of doc.wb.sheets) {
      if (s.kind !== 'worksheet') continue;
      let touched = false;
      for (const rowMap of s.sheet.rows.values()) {
        for (const cell of rowMap.values()) {
          const v = cell.value;
          if (v === null || typeof v !== 'object' || v instanceof Date || v.kind !== 'formula' || !v.formula) continue;
          const f = deleteSheetInFormula(v.formula, title, order);
          if (f === v.formula) continue;
          if (!touched) {
            tx.cells(s.sheet, { r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL });
            touched = true;
          }
          cell.value = { ...v, formula: f };
        }
      }
    }
    doc.wb.definedNames = doc.wb.definedNames.map((dn) => {
      const value = deleteSheetInFormula(dn.value, title, order);
      return value === dn.value ? dn : { ...dn, value };
    });
  });
  doc.activeSheetIndex = Math.min(next, doc.wb.sheets.length - 1);
  doc.layoutVersion++;
}

export function renameSheetAt(ctl: EditorController, index: number, name: string): string | undefined {
  const doc = ctl.doc;
  const ref = doc.wb.sheets[index];
  if (!ref) return undefined;
  const trimmed = name.trim();
  if (!trimmed || trimmed.length > 31 || /[\\/?*[\]:]/.test(trimmed) || trimmed.startsWith("'") || trimmed.endsWith("'")) return 'invalidSheetName';
  if (doc.wb.sheets.some((s, i) => i !== index && s.sheet.title.toLowerCase() === trimmed.toLowerCase())) return 'duplicateSheetName';
  const old = ref.sheet.title;
  if (old === trimmed) return undefined;
  doc.transact('Rename Sheet', (tx) => {
    tx.structural = true;
    tx.title(ref.sheet as Worksheet);
    tx.workbook('definedNames');
    for (const s of doc.wb.sheets) if (s.kind === 'worksheet') tx.cells(s.sheet, { r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL });
    // renameSheet re-points PivotTable sources at the new title.
    for (const s of doc.wb.sheets) if (s.kind === 'worksheet' && s.sheet.pivotTables) tx.sheet(s.sheet, 'pivotTables');
    renameSheet(doc.wb, old, trimmed);
    for (const s of doc.wb.sheets) {
      if (s.kind !== 'worksheet') continue;
      for (const rowMap of s.sheet.rows.values()) {
        for (const cell of rowMap.values()) {
          const v = cell.value;
          if (v === null || typeof v !== 'object' || v instanceof Date || v.kind !== 'formula' || !v.formula) continue;
          const f = renameSheetInFormula(v.formula, old, trimmed);
          if (f !== v.formula) cell.value = { ...v, formula: f };
        }
      }
    }
    for (const dn of doc.wb.definedNames) dn.value = renameSheetInFormula(dn.value, old, trimmed);
  });
  return undefined;
}

export function moveSheetTo(ctl: EditorController, from: number, to: number): void {
  const doc = ctl.doc;
  const title = doc.wb.sheets[from]?.sheet.title;
  if (title === undefined || from === to) return;
  doc.transact('Move Sheet', (tx) => {
    tx.structural = true;
    tx.workbook('sheets', 'activeSheetIndex', 'definedNames');
    moveSheet(doc.wb, title, to);
  });
  doc.activeSheetIndex = Math.max(0, Math.min(to, doc.wb.sheets.length - 1));
  doc.layoutVersion++;
}

/** Duplicate a worksheet (Move or Copy ▸ Create a copy). Cells, styles and sheet settings are cloned. */
export function duplicateSheet(ctl: EditorController, index: number, to: number): void {
  const doc = ctl.doc;
  const ref = doc.wb.sheets[index];
  if (ref?.kind !== 'worksheet') return;
  const src = ref.sheet;
  let name = `${src.title} (2)`;
  for (let n = 3; doc.wb.sheets.some((s) => s.sheet.title.toLowerCase() === name.toLowerCase()); n++) name = `${src.title} (${n})`;
  doc.transact('Copy Sheet', (tx) => {
    tx.structural = true;
    tx.workbook('sheets', 'activeSheetIndex', 'definedNames');
    const sheet = addWorksheet(doc.wb, name, { index: to });
    const { rows: _rows, title: _title, ...rest } = src;
    Object.assign(sheet, structuredClone(rest));
    for (const [r, rowMap] of src.rows) {
      const m = new Map<number, ReturnType<typeof makeCell>>();
      for (const [c, cell] of rowMap) {
        const nc = makeCell(r, c, cell.value, cell.styleId);
        if (cell.hyperlinkId !== undefined) nc.hyperlinkId = cell.hyperlinkId;
        if (cell.commentId !== undefined) nc.commentId = cell.commentId;
        m.set(c, nc);
      }
      sheet.rows.set(r, m);
    }
    // Tables must have workbook-unique names; drop them from the copy rather than collide.
    sheet.tables = [];
    renewThreadIds(sheet);
  });
  doc.activateSheet(to);
}

/**
 * Excel drops a thread whose id, or a mention whose id, also appears on
 * another sheet, so a copied sheet's threads get fresh ids (replies follow
 * their root) the way the library's own sheet duplication does. Authors keep
 * pointing at the same `Workbook.persons` entries.
 */
function renewThreadIds(ws: Worksheet): void {
  const guid = (): string => `{${crypto.randomUUID().toUpperCase()}}`;
  const threads = ws.threadedComments ?? [];
  const ids = new Map(threads.map((c) => [c.id, guid()]));
  for (const c of threads) {
    c.id = ids.get(c.id) ?? c.id;
    if (c.parentId !== undefined) c.parentId = ids.get(c.parentId) ?? c.parentId;
    for (const m of c.mentions ?? []) m.mentionId = guid();
  }
}

export function setSheetHidden(ctl: EditorController, index: number, hidden: boolean): void {
  const doc = ctl.doc;
  const ref = doc.wb.sheets[index];
  if (!ref) return;
  if (hidden && doc.wb.sheets.filter((s) => s.state === 'visible').length <= 1) {
    ctl.dialog = { kind: 'alert', props: { message: 'lastSheet' } };
    return;
  }
  doc.transact(hidden ? 'Hide Sheet' : 'Unhide Sheet', (tx) => {
    tx.workbook('sheets');
    setSheetState(doc.wb, ref.sheet.title, hidden ? 'hidden' : 'visible');
  });
  if (hidden && index === doc.activeSheetIndex) {
    const next = doc.wb.sheets.findIndex((s) => s.state === 'visible' && s.kind === 'worksheet');
    if (next >= 0) doc.activateSheet(next);
  } else if (!hidden) doc.activateSheet(index);
}

export function setTabColor(ctl: EditorController, index: number, rgb: string | null): void {
  const doc = ctl.doc;
  const ref = doc.wb.sheets[index];
  if (ref?.kind !== 'worksheet') return;
  doc.transact('Tab Color', (tx) => {
    tx.sheet(ref.sheet, 'sheetProperties');
    if (rgb) setSheetTabColor(ref.sheet, rgb);
    else if (ref.sheet.sheetProperties) {
      const { tabColor: _removed, ...rest } = ref.sheet.sheetProperties;
      ref.sheet.sheetProperties = rest;
    }
  });
}

// ---- data -----------------------------------------------------------------------------

/** Alt+= / AutoSum: SUM (or another function) over the numbers above or to the left. */
export function autoSum(ctl: EditorController, fn = 'SUM'): void {
  const ws = ctl.doc.ws;
  const sel = ctl.doc.selection;
  const r = currentRange(sel);
  const isNum = (row: number, col: number) => {
    const v = getCellAt(ws, row, col)?.value;
    if (typeof v === 'number') return true;
    return v !== null && v !== undefined && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' && typeof v.cachedValue === 'number';
  };
  if (r.r1 === r.r2 && r.c1 === r.c2) {
    const { row, col } = sel.active;
    let top = row - 1;
    while (top >= 1 && isNum(top, col)) top--;
    let range: Range | undefined;
    if (top < row - 1) range = { r1: top + 1, c1: col, r2: row - 1, c2: col };
    else {
      let left = col - 1;
      while (left >= 1 && isNum(row, left)) left--;
      if (left < col - 1) range = { r1: row, c1: left + 1, r2: row, c2: col - 1 };
    }
    const arg = range ? rangeAddress(range) : '';
    ctl.startEdit(`=${fn}(${arg})`);
    const e = ctl.edit;
    if (e) {
      // Excel leaves the proposed range selected inside the parentheses.
      e.selStart = fn.length + 2;
      e.selEnd = fn.length + 2 + arg.length;
      if (range) e.point = { start: fn.length + 2, end: fn.length + 2 + arg.length, anchor: { row: range.r1, col: range.c1 }, cursor: { row: range.r2, col: range.c2 } };
      e.mode = 'point';
    }
    return;
  }
  // A multi-cell selection gets totals below each column (or right of each row when one row).
  const doc = ctl.doc;
  const vertical = r.r2 > r.r1 || r.c1 === r.c2;
  doc.transact('AutoSum', (tx) => {
    if (vertical) {
      const target: Range = { r1: r.r2 + 1, c1: r.c1, r2: r.r2 + 1, c2: r.c2 };
      tx.cells(ws, target);
      for (let c = r.c1; c <= r.c2; c++) putFormula(ws, r.r2 + 1, c, `${fn}(${rangeAddress({ r1: r.r1, c1: c, r2: r.r2, c2: c })})`);
    } else {
      const target: Range = { r1: r.r1, c1: r.c2 + 1, r2: r.r2, c2: r.c2 + 1 };
      tx.cells(ws, target);
      for (let row = r.r1; row <= r.r2; row++) putFormula(ws, row, r.c2 + 1, `${fn}(${rangeAddress({ r1: row, c1: r.c1, r2: row, c2: r.c2 })})`);
    }
  });
}

function putFormula(ws: Worksheet, row: number, col: number, formula: string): void {
  const value: CellValue = { kind: 'formula', t: 'normal', formula };
  const cell = getCellAt(ws, row, col);
  if (cell) cell.value = value;
  else {
    let rowMap = ws.rows.get(row);
    if (!rowMap) {
      rowMap = new Map();
      ws.rows.set(row, rowMap);
    }
    rowMap.set(col, makeCell(row, col, value));
  }
}

/** Ctrl+; / Ctrl+Shift+; stamps as typed text, so they go through normal input parsing. */
export function nowText(kind: 'date' | 'time', ymd: boolean): string {
  const d = new Date();
  if (kind === 'time') return `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`;
  return ymd ? `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}` : `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

/** Home ▸ Format ▸ AutoFit Row Height / Column Width over the selected lines. */
export function autofitSelection(ctl: EditorController, axis: 'row' | 'col'): void {
  const indices: number[] = [];
  for (const [lo, hi] of selectedLines(ctl, axis)) {
    for (let i = lo; i <= Math.min(hi, lo + 5000); i++) indices.push(i);
  }
  if (axis === 'row') autofitRows(ctl, indices);
  else autofitColumns(ctl, indices);
}

/** Home ▸ Format ▸ Lock Cell toggles the Protection ▸ Locked flag. */
export function toggleLocked(ctl: EditorController): void {
  const { row, col } = ctl.doc.selection.active;
  const cell = getCellAt(ctl.doc.ws, row, col);
  const locked = cell ? getCellProtection(ctl.doc.wb, cell).locked !== false : true;
  format(ctl, { protection: { locked: !locked } }, 'Lock Cell');
}

// ---- view & page layout ---------------------------------------------------------

export function colName(col: number): string {
  return columnLetterFromIndex(col);
}

export function setSheetViewMode(ctl: EditorController, mode: 'normal' | 'pageBreakPreview' | 'pageLayout'): void {
  setViewMode(ctl.doc.ws, mode);
  ctl.doc.dirty = true;
  ctl.doc.version++;
  // Each view lays the sheet out differently, so scroll back to the active cell.
  ctl.doc.setScroll(0, 0);
  const { row, col } = ctl.doc.selection.active;
  ctl.reveal(row, col);
}

/** View ▸ Zoom to Selection: fit the selected range into the window. */
export function zoomToSelection(ctl: EditorController): void {
  const r = currentRange(ctl.doc.selection);
  const geo = ctl.geometry;
  const w = ctl.doc.cols.offsetOf(Math.min(r.c2, MAX_COL) + 1) - ctl.doc.cols.offsetOf(r.c1);
  const h = ctl.doc.rows.offsetOf(Math.min(r.r2, MAX_ROW) + 1) - ctl.doc.rows.offsetOf(r.r1);
  const zoom = Math.min((geo.width - geo.headerW - 20) / Math.max(1, w), (geo.height - geo.headerH - 20) / Math.max(1, h));
  ctl.doc.setZoom(Math.floor(Math.min(4, Math.max(0.1, zoom)) * 100) / 100);
  ctl.doc.setScroll(ctl.doc.cols.offsetOf(r.c1) * ctl.doc.zoom - geo.frozenW, ctl.doc.rows.offsetOf(r.r1) * ctl.doc.zoom - geo.frozenH);
}

function pageSetupOf(ws: Worksheet): PageSetup {
  ws.pageSetup ??= {};
  return ws.pageSetup;
}

function editPage(ctl: EditorController, label: string, fn: (ws: Worksheet) => void): void {
  const ws = ctl.doc.ws;
  ctl.doc.transact(label, (tx) => {
    tx.sheet(ws, 'pageSetup', 'pageMargins', 'printOptions', 'rowBreaks', 'colBreaks', 'sheetProperties');
    fn(ws);
  });
}

const MARGIN_PRESETS = {
  normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
} as const;

export function setMargins(ctl: EditorController, preset: keyof typeof MARGIN_PRESETS): void {
  editPage(ctl, 'Margins', (ws) => {
    ws.pageMargins = { ...MARGIN_PRESETS[preset] };
  });
}

export function setOrientation(ctl: EditorController, orientation: 'portrait' | 'landscape'): void {
  editPage(ctl, 'Orientation', (ws) => {
    pageSetupOf(ws).orientation = orientation;
  });
}

export function setPaperSize(ctl: EditorController, paperSize: number): void {
  editPage(ctl, 'Size', (ws) => {
    pageSetupOf(ws).paperSize = paperSize;
  });
}

export function setFitTo(ctl: EditorController, axis: 'width' | 'height', value: string): void {
  editPage(ctl, 'Scale to Fit', (ws) => {
    const ps = pageSetupOf(ws);
    const n = value === 'auto' ? 0 : Number(value);
    if (axis === 'width') ps.fitToWidth = n;
    else ps.fitToHeight = n;
    const fit = (ps.fitToWidth ?? 1) !== 0 || (ps.fitToHeight ?? 1) !== 0;
    ws.sheetProperties = { ...ws.sheetProperties, pageSetUpPr: { ...ws.sheetProperties?.pageSetUpPr, fitToPage: fit } };
  });
}

export function setPrintScale(ctl: EditorController, scale: number): void {
  if (!Number.isFinite(scale) || scale < 10 || scale > 400) return;
  editPage(ctl, 'Scale', (ws) => {
    pageSetupOf(ws).scale = Math.round(scale);
    ws.sheetProperties = { ...ws.sheetProperties, pageSetUpPr: { ...ws.sheetProperties?.pageSetUpPr, fitToPage: false } };
  });
}

export function setPrintOption(ctl: EditorController, key: 'gridLines' | 'headings', on: boolean): void {
  editPage(ctl, 'Sheet Options', (ws) => {
    ws.printOptions = { ...ws.printOptions, [key]: on };
  });
}

const PRINT_AREA = '_xlnm.Print_Area';

export function setPrintArea(ctl: EditorController, set: boolean): void {
  const doc = ctl.doc;
  const scope = doc.activeSheetIndex;
  const refs = doc.selection.ranges.map((r) => `${quoteSheetName(doc.ws.title)}!${rangeAddress(r, true)}`).join(',');
  doc.transact(set ? 'Set Print Area' : 'Clear Print Area', (tx) => {
    tx.workbook('definedNames');
    const others = doc.wb.definedNames.filter((d) => !(d.name === PRINT_AREA && d.scope === scope));
    doc.wb.definedNames = set ? [...others, { name: PRINT_AREA, value: refs, scope }] : others;
  });
}

export function insertPageBreak(ctl: EditorController): void {
  const { row, col } = ctl.doc.selection.active;
  editPage(ctl, 'Insert Page Break', (ws) => {
    if (row > 1) ws.rowBreaks = [...ws.rowBreaks.filter((b) => b.id !== row - 1), { id: row - 1, max: MAX_COL - 1, man: true }];
    if (col > 1) ws.colBreaks = [...ws.colBreaks.filter((b) => b.id !== col - 1), { id: col - 1, max: MAX_ROW - 1, man: true }];
  });
}

/** Removes the breaks on the active cell's top and left edges — the ones Insert Page Break would add there. */
export function removePageBreak(ctl: EditorController): void {
  const { row, col } = ctl.doc.selection.active;
  const ws = ctl.doc.ws;
  if (!ws.rowBreaks.some((b) => b.id === row - 1) && !ws.colBreaks.some((b) => b.id === col - 1)) return;
  editPage(ctl, 'Remove Page Break', (sheet) => {
    sheet.rowBreaks = sheet.rowBreaks.filter((b) => b.id !== row - 1);
    sheet.colBreaks = sheet.colBreaks.filter((b) => b.id !== col - 1);
  });
}

export function resetPageBreaks(ctl: EditorController): void {
  editPage(ctl, 'Reset Page Breaks', (ws) => {
    ws.rowBreaks = [];
    ws.colBreaks = [];
  });
}

// ---- formulas tab -----------------------------------------------------------------

/** Function Library menus: start a formula with the function, as Excel's Function Arguments flow does. */
export function insertFunctionCall(ctl: EditorController, name: string): void {
  recentFunctions.use(name);
  const e = ctl.edit;
  if (e && e.text.startsWith('=')) {
    const ins = `${name}(`;
    ctl.setEditText(e.text.slice(0, e.selStart) + ins + e.text.slice(e.selEnd), e.selStart + ins.length);
    return;
  }
  ctl.startEdit(`=${name}()`);
  if (ctl.edit) ctl.edit.selStart = ctl.edit.selEnd = name.length + 2;
}

export function insertNameIntoFormula(ctl: EditorController, name: string): void {
  const e = ctl.edit;
  if (e) {
    ctl.setEditText(e.text.slice(0, e.selStart) + name + e.text.slice(e.selEnd), e.selStart + name.length);
    return;
  }
  ctl.startEdit(`=${name}`);
}

/** Create from Selection: names from the top row (or left column) label each column (row). */
export function createNamesFromSelection(ctl: EditorController): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const r = currentRange(doc.selection);
  const names: Array<{ name: string; value: string }> = [];
  const clean = (s: string) => s.trim().replace(/[^A-Za-z0-9_.À-￿]+/g, '_').replace(/^(\d)/, '_$1');
  const useTop = r.r2 > r.r1;
  if (useTop) {
    for (let c = r.c1; c <= r.c2; c++) {
      const label = getCellAt(ws, r.r1, c)?.value;
      if (typeof label !== 'string' || !label.trim()) continue;
      names.push({ name: clean(label), value: `${quoteSheetName(ws.title)}!${rangeAddress({ r1: r.r1 + 1, c1: c, r2: r.r2, c2: c }, true)}` });
    }
  } else {
    for (let row = r.r1; row <= r.r2; row++) {
      const label = getCellAt(ws, row, r.c1)?.value;
      if (typeof label !== 'string' || !label.trim()) continue;
      names.push({ name: clean(label), value: `${quoteSheetName(ws.title)}!${rangeAddress({ r1: row, c1: r.c1 + 1, r2: row, c2: r.c2 }, true)}` });
    }
  }
  // A label that reads as a cell reference (`Q1`, `R`, `C`) or is otherwise
  // not a legal name would make a file Excel rejects, so it gets a leading `_`.
  const taken = new Set(doc.wb.definedNames.filter((d) => d.scope === undefined).map((d) => d.name.toLowerCase()));
  const fresh: Array<{ name: string; value: string }> = [];
  for (const n of names) {
    const name = validateName(n.name) ? `_${n.name}` : n.name;
    if (validateName(name) || taken.has(name.toLowerCase())) continue;
    taken.add(name.toLowerCase());
    fresh.push({ name, value: n.value });
  }
  if (fresh.length === 0) return;
  doc.transact('Create Names', (tx) => {
    tx.structural = true;
    tx.workbook('definedNames');
    doc.wb.definedNames = [...doc.wb.definedNames, ...fresh];
  });
}

export function setCalcMode(ctl: EditorController, mode: 'auto' | 'manual'): void {
  const wb = ctl.doc.wb;
  ctl.doc.transact('Calculation Options', (tx) => {
    tx.workbook('calcProperties');
    wb.calcProperties = { ...wb.calcProperties, calcMode: mode };
  });
}

/** Trace Precedents / Dependents: outline the related cells on the grid (arrows in Excel). */
export function traceCells(ctl: EditorController, kind: 'precedents' | 'dependents' | 'clear'): void {
  if (kind === 'clear') {
    ctl.traces = [];
    return;
  }
  const ws = ctl.doc.ws;
  const { row, col } = ctl.doc.selection.active;
  const out: Range[] = [];
  if (kind === 'precedents') {
    const v = getCellAt(ws, row, col)?.value;
    if (v && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' && v.formula) {
      for (const ref of formulaReferences(v.formula)) if (!ref.sheet || ref.sheet === ws.title) out.push(ref.range);
    }
  } else {
    for (const rowMap of ws.rows.values()) {
      for (const cell of rowMap.values()) {
        const v = cell.value;
        if (!v || typeof v !== 'object' || v instanceof Date || v.kind !== 'formula' || !v.formula) continue;
        const hit = formulaReferences(v.formula).some((ref) => (!ref.sheet || ref.sheet === ws.title) && row >= ref.range.r1 && row <= ref.range.r2 && col >= ref.range.c1 && col <= ref.range.c2);
        if (hit) out.push({ r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });
      }
    }
  }
  if (out.length === 0) ctl.toast = kind === 'precedents' ? 'noPrecedents' : 'noDependents';
  ctl.traces = out.map((range) => ({ from: { row, col }, range, kind }));
}

/** Error Checking: jump to the next cell showing an error value. */
export function errorChecking(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const a = ctl.doc.selection.active;
  const errors: Array<{ row: number; col: number }> = [];
  for (const rowMap of ws.rows.values()) {
    for (const cell of rowMap.values()) {
      const v = cell.value;
      const isErr = v !== null && typeof v === 'object' && !(v instanceof Date) && (v.kind === 'error' || (v.kind === 'formula' && v.cachedValueType === 'error'));
      if (isErr) errors.push({ row: cell.row, col: cell.col });
    }
  }
  errors.sort((x, y) => x.row - y.row || x.col - y.col);
  const next = errors.find((p) => p.row > a.row || (p.row === a.row && p.col > a.col)) ?? errors[0];
  if (!next) {
    ctl.toast = 'noErrors';
    return;
  }
  ctl.selectCell(next);
  ctl.reveal(next.row, next.col);
}

// ---- data tab -------------------------------------------------------------------


/** Data ▸ Group / Ungroup: columns when whole columns are selected, rows otherwise. */
export function groupSelection(ctl: EditorController, group: boolean): void {
  const r = currentRange(ctl.doc.selection);
  if (r.r1 === 1 && r.r2 === MAX_ROW) groupLines(ctl.doc, 'col', r.c1, r.c2, group);
  else groupLines(ctl.doc, 'row', r.r1, r.r2, group);
}

/** Data ▸ Show / Hide Detail on the innermost group touching the active cell. */
export function showDetail(ctl: EditorController, show: boolean): void {
  const ws = ctl.doc.ws;
  const r = currentRange(ctl.doc.selection);
  const axis = r.r1 === 1 && r.r2 === MAX_ROW ? 'col' : 'row';
  const at = axis === 'row' ? ctl.doc.selection.active.row : ctl.doc.selection.active.col;
  const outline = axis === 'row' ? ctl.outline.rows : ctl.outline.cols;
  const run = outline.runs
    .filter((x) => (at >= x.start && at <= x.end) || at === x.summary)
    .sort((x, y) => y.level - x.level)
    .find((x) => isCollapsed(ws, axis, x) === show);
  if (run) toggleRun(ctl.doc, axis, outline, run);
}

/** Circle Invalid Data: outline cells that break their data-validation rule. */
export function circleInvalid(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const out: Range[] = [];
  for (const dv of ws.dataValidations) {
    for (const b of dv.sqref.ranges) {
      forEachCellInRange(ws, { r1: b.minRow, c1: b.minCol, r2: b.maxRow, c2: b.maxCol }, (cell) => {
        if (!validateValue(ctl, dv, cell.value, cell.row, cell.col)) out.push({ r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });
      });
    }
  }
  ctl.invalidCircles = out;
  if (out.length === 0) ctl.toast = 'noInvalidData';
}

export function clearInvalidCircles(ctl: EditorController): void {
  ctl.invalidCircles = [];
}

// ---- review tab -----------------------------------------------------------------

function noteCells(ws: Worksheet): Array<{ row: number; col: number }> {
  return ws.legacyComments
    .map((c) => coordinateFromString(c.ref.replaceAll('$', '').split(':')[0] ?? 'A1'))
    .map((p) => ({ row: p.row, col: columnIndexFromLetter(p.column) }))
    .sort((a, b) => a.row - b.row || a.col - b.col);
}

export function stepNote(ctl: EditorController, step: 1 | -1): void {
  const notes = noteCells(ctl.doc.ws);
  if (notes.length === 0) return;
  const a = ctl.doc.selection.active;
  const next = step === 1 ? (notes.find((p) => p.row > a.row || (p.row === a.row && p.col > a.col)) ?? notes[0]) : ([...notes].reverse().find((p) => p.row < a.row || (p.row === a.row && p.col < a.col)) ?? notes.at(-1));
  if (!next) return;
  ctl.selectCell(next);
  ctl.reveal(next.row, next.col);
}

export function deleteNotes(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const ranges = ctl.doc.selection.ranges;
  const kept = ws.legacyComments.filter((c) => {
    const p = refCell(c.ref);
    return !ranges.some((r) => inRange(r, p.row, p.col));
  });
  if (kept.length === ws.legacyComments.length) return;
  // Only notes: threaded comments have their own Delete.
  ctl.doc.transact('Delete Note', (tx) => {
    tx.sheet(ws, 'legacyComments');
    ws.legacyComments = kept;
  });
}

/** Review ▸ Unprotect Sheet: asks for the password when the protection has one. */
export function unprotectSheet(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  if (ws.sheetProtection?.hashValue) {
    ctl.openDialog('unprotect', { target: 'sheet' });
    return;
  }
  ctl.doc.transact('Unprotect Sheet', (tx) => {
    tx.sheet(ws, 'sheetProtection');
    delete ws.sheetProtection;
  });
}

/** Review ▸ Protect Workbook toggles: protect through the dialog, unprotect (asking for the password if set). */
export function toggleWorkbookProtection(ctl: EditorController): void {
  const wb = ctl.doc.wb;
  if (!wb.workbookProtection) {
    ctl.openDialog('protectWorkbook');
    return;
  }
  if (wb.workbookProtection.workbookHashValue) {
    ctl.openDialog('unprotect', { target: 'workbook' });
    return;
  }
  ctl.doc.transact('Unprotect Workbook', (tx) => {
    tx.workbook('workbookProtection');
    delete wb.workbookProtection;
  });
}


/** Insert ▸ Pictures: embed the image at the active cell at its natural size (capped to the window). */
export async function insertPicture(ctl: EditorController, file: File): Promise<void> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  let image;
  try {
    image = loadImage(bytes);
  } catch {
    ctl.dialog = { kind: 'alert', props: { message: 'unsupportedImage' } };
    return;
  }
  const scale = Math.min(1, 640 / Math.max(image.width, 1), 480 / Math.max(image.height, 1));
  const { row, col } = ctl.doc.selection.active;
  const ws = ctl.doc.ws;
  ctl.doc.transact('Insert Picture', (tx) => {
    tx.sheet(ws, 'drawing');
    addImageAt(ws, `${columnLetterFromIndex(col)}${row}`, image, { widthPx: Math.round(image.width * scale), heightPx: Math.round(image.height * scale) });
  });
}

/** Data ▸ Flash Fill (Cmd/Ctrl+E) on the active column, down to the end of the adjacent data. */
export function flashFillActive(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const { col } = ctl.doc.selection.active;
  let top = ctl.doc.selection.active.row;
  while (top > 1 && !isBlank(getCellAt(ws, top - 1, col - 1))) top--;
  let bottom = ctl.doc.selection.active.row;
  while (bottom < MAX_ROW && !isBlank(getCellAt(ws, bottom + 1, col - 1))) bottom++;
  // Skip a header row: Flash Fill starts at the first example.
  const values = flashFill(ws, col, top, bottom);
  if (!values || values.size === 0) {
    ctl.dialog = { kind: 'alert', props: { message: 'flashFillFailed' } };
    return;
  }
  ctl.doc.transact('Flash Fill', (tx) => {
    tx.cells(ws, { r1: top, c1: col, r2: bottom, c2: col });
    for (const [row, v] of values) {
      const cell = getCellAt(ws, row, col);
      if (cell) cell.value = v;
      else {
        let rowMap = ws.rows.get(row);
        if (!rowMap) {
          rowMap = new Map();
          ws.rows.set(row, rowMap);
        }
        rowMap.set(col, makeCell(row, col, v, ctl.defaultStyleAt(row, col)));
      }
    }
  });
  ctl.toast = 'flashFillDone';
}

/**
 * Cmd/Ctrl+' copies the formula above into the active cell unchanged (no
 * reference shift); Ctrl+Shift+" copies the value above. Both open the cell
 * for editing, as Excel does.
 */
export function copyFromAbove(ctl: EditorController, what: 'formula' | 'value'): void {
  const { row, col } = ctl.doc.selection.active;
  if (row <= 1) return;
  const above = getCellAt(ctl.doc.ws, row - 1, col);
  if (!above || isBlank(above)) return;
  const v = above.value;
  const isFormula = v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula';
  ctl.startEdit(what === 'formula' && isFormula ? `=${fromStorageFormula(v.formula)}` : ctl.displayText(above));
}
