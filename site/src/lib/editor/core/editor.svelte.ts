// The editor's single source of truth: the live workbook, per-sheet view
// state (selection, scroll, zoom), undo history and the formula engine.
//
// The library model is a large mutable object graph (a Map per row), so it is
// held in `$state.raw` and never deep-proxied; every committed mutation bumps
// `version`, and derived values read that counter. Layout-affecting edits bump
// `layoutVersion` as well, so the row/column pixel indices are only rebuilt
// when sizes actually change rather than on every keystroke.

import type { Cell } from '@office-kit/xlsx/cell';
import type { Workbook } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import { CalcEngine, type CellRef } from '../calc/index.ts';
import type { CellPos, Range } from './address.ts';
import { clampPos, rangeOf } from './address.ts';
import { getCellAt, isBlank } from './cells.ts';
import { syncTableHeaders } from './tables.ts';
import { changedCells, EditRefusedError, History, Transaction, type HistoryStep, type TransactionGuard } from './history.ts';
import { MergeIndex } from './merges.ts';
import { buildColumnAxis, buildRowAxis } from './metrics.ts';
import { StyleResolver } from './render-style.ts';
import { currentRange, singleCell, type Selection } from './selection.ts';
import type { AxisIndex } from './axis.ts';

export interface SheetView {
  selection: Selection;
  scrollX: number;
  scrollY: number;
  zoom: number;
}

/** What undo/redo restores alongside the model: where the user was. */
interface ViewToken {
  readonly sheet: number;
  readonly selection: Selection;
}

const LAYOUT_FIELDS = new Set(['columnDimensions', 'rowDimensions', 'defaultColumnWidth', 'defaultRowHeight', 'baseColWidth']);

export function blankWorkbook(): Workbook {
  const wb = createWorkbook();
  addWorksheet(wb, 'Sheet1');
  return wb;
}

export class SpreadsheetEditor {
  wb = $state.raw<Workbook>(blankWorkbook());
  fileName = $state('Book1.xlsx');
  dirty = $state(false);
  /** Bumped on every committed model change. */
  version = $state(0);
  /** Bumped when row heights, column widths, hidden flags or merges change. */
  layoutVersion = $state(0);
  activeSheetIndex = $state(0);
  #views = new Map<Worksheet, SheetView>();
  /** Reactive mirror of the active sheet's view; written through `setSelection` etc. */
  selection = $state.raw<Selection>(singleCell({ row: 1, col: 1 }));
  scrollX = $state(0);
  scrollY = $state(0);
  zoom = $state(1);

  styles = $derived.by(() => new StyleResolver(this.wb));
  sheets = $derived.by(() => {
    void this.version;
    return this.wb.sheets.map((s) => s.sheet);
  });

  ws = $derived.by<Worksheet>(() => {
    void this.version;
    const ref = this.wb.sheets[this.activeSheetIndex];
    if (ref?.kind === 'worksheet') return ref.sheet;
    const first = this.wb.sheets.find((s) => s.kind === 'worksheet');
    if (first?.kind !== 'worksheet') throw new Error('workbook has no worksheet');
    return first.sheet;
  });

  cols = $derived.by<AxisIndex>(() => {
    void this.layoutVersion;
    return buildColumnAxis(this.ws);
  });

  rows = $derived.by<AxisIndex>(() => {
    void this.layoutVersion;
    return buildRowAxis(this.ws);
  });

  merges = $derived.by<MergeIndex>(() => {
    void this.layoutVersion;
    return new MergeIndex(this.ws);
  });

  calc: CalcEngine;
  readonly history: History<ViewToken>;
  canUndo = $state(false);
  canRedo = $state(false);

  constructor(wb?: Workbook) {
    if (wb) this.wb = wb;
    this.history = new History<ViewToken>(() => this.wb);
    this.calc = new CalcEngine(this.wb);
    this.calc.recalculateAll();
    this.activeSheetIndex = this.#initialSheetIndex();
    this.#loadView();
  }

  #initialSheetIndex(): number {
    const i = this.wb.activeSheetIndex;
    return this.wb.sheets[i]?.kind === 'worksheet' ? i : Math.max(0, this.wb.sheets.findIndex((s) => s.kind === 'worksheet'));
  }

  // ---- document lifecycle -------------------------------------------------

  replaceWorkbook(wb: Workbook, fileName: string): void {
    this.wb = wb;
    this.fileName = fileName;
    this.dirty = false;
    this.#views = new Map();
    this.history.clear();
    this.#syncHistoryFlags();
    this.calc = new CalcEngine(wb);
    this.calc.recalculateAll();
    this.activeSheetIndex = this.#initialSheetIndex();
    this.version++;
    this.layoutVersion++;
    this.#loadView();
  }

  async open(file: File): Promise<void> {
    const wb = await loadWorkbook(fromArrayBuffer(await file.arrayBuffer()));
    this.replaceWorkbook(wb, file.name);
  }

  async toBytes(): Promise<Uint8Array> {
    this.wb.activeSheetIndex = this.activeSheetIndex;
    syncTableHeaders(this.wb, (ws, row, col) => {
      const cell = getCellAt(ws, row, col);
      return cell && !isBlank(cell) ? getCellDisplayText(this.wb, cell).trim() : '';
    });
    return workbookToBytes(this.wb);
  }

  // ---- sheet & view -------------------------------------------------------

  #viewFor(ws: Worksheet): SheetView {
    let v = this.#views.get(ws);
    if (!v) {
      const sv = ws.views[0];
      const activeRef = sv?.selection?.activeCell;
      const pos = activeRef ? parseA1(activeRef) : undefined;
      v = {
        selection: singleCell(pos ?? { row: 1, col: 1 }),
        scrollX: 0,
        scrollY: 0,
        zoom: sv?.zoomScale ? sv.zoomScale / 100 : 1,
      };
      this.#views.set(ws, v);
    }
    return v;
  }

  #loadView(): void {
    const v = this.#viewFor(this.ws);
    this.selection = v.selection;
    this.scrollX = v.scrollX;
    this.scrollY = v.scrollY;
    this.zoom = v.zoom;
  }

  #saveView(): void {
    const v = this.#viewFor(this.ws);
    v.selection = this.selection;
    v.scrollX = this.scrollX;
    v.scrollY = this.scrollY;
    v.zoom = this.zoom;
  }

  activateSheet(index: number): void {
    if (index === this.activeSheetIndex || this.wb.sheets[index]?.kind !== 'worksheet') return;
    this.#saveView();
    this.activeSheetIndex = index;
    this.layoutVersion++;
    this.#loadView();
  }

  setSelection(sel: Selection): void {
    this.selection = sel;
  }

  setScroll(x: number, y: number): void {
    this.scrollX = Math.max(0, x);
    this.scrollY = Math.max(0, y);
  }

  setZoom(zoom: number): void {
    this.zoom = Math.min(4, Math.max(0.1, zoom));
  }

  get active(): CellPos {
    return this.selection.active;
  }

  get activeRangeValue(): Range {
    return currentRange(this.selection);
  }

  cellAt(row: number, col: number): Cell | undefined {
    return getCellAt(this.ws, row, col);
  }

  // ---- mutation -----------------------------------------------------------

  /**
   * Run `fn` as one undoable step. `fn` declares what it touches on the
   * transaction before mutating; the step records those parts before and
   * after, and the formula engine recalculates only the touched cells'
   * dependents unless the step was structural.
   */
  /** Refuses edits the sheet's protection forbids; set by the controller. */
  guard: TransactionGuard | undefined = undefined;
  /** Told why an edit was refused, so the UI can say so. */
  onRefused: ((reason: EditRefusedError['reason']) => void) | undefined = undefined;

  /** Returns the callback's result, or undefined when the guard refused the edit (nothing changed). */
  transact<T>(label: string, fn: (tx: Transaction) => T): T | undefined {
    const tx = new Transaction(this.wb, this.guard);
    const viewBefore: ViewToken = { sheet: this.activeSheetIndex, selection: this.selection };
    let result: T;
    try {
      result = fn(tx);
    } catch (err) {
      if (!(err instanceof EditRefusedError)) throw err;
      tx.rollback();
      this.onRefused?.(err.reason);
      return undefined;
    }
    if (tx.parts.length === 0) return result;
    const step: HistoryStep<ViewToken> = {
      label,
      before: tx.before,
      after: tx.finish(),
      viewBefore,
      viewAfter: { sheet: this.activeSheetIndex, selection: this.selection },
      structural: tx.structural,
    };
    this.history.push(step);
    this.#afterChange(step);
    return result;
  }

  #afterChange(step: HistoryStep<ViewToken>): void {
    if (step.structural) {
      this.calc.invalidateAll();
      this.calc.recalculateAll();
    } else {
      const changed: CellRef[] = changedCells(step).map(({ ws, row, col }) => ({ sheet: ws.title, row, col }));
      if (changed.length > 0) this.calc.update(changed);
      if (step.before.some((s) => s.part.kind === 'sheet' && s.part.fields.includes('rowDimensions'))) this.calc.recalculateSubtotals();
    }
    if (step.structural || step.before.some((s) => s.part.kind === 'sheet' && s.part.fields.some((f) => LAYOUT_FIELDS.has(f) || f === 'mergedCells'))) {
      this.layoutVersion++;
    }
    this.dirty = true;
    this.version++;
    this.#syncHistoryFlags();
  }

  #syncHistoryFlags(): void {
    this.canUndo = this.history.canUndo;
    this.canRedo = this.history.canRedo;
  }

  undo(): void {
    const step = this.history.undo();
    if (!step) return;
    this.#restoreView(step.viewBefore);
    this.#afterChange(step);
  }

  redo(): void {
    const step = this.history.redo();
    if (!step) return;
    this.#restoreView(step.viewAfter);
    this.#afterChange(step);
  }

  #restoreView(token: ViewToken): void {
    if (token.sheet !== this.activeSheetIndex && this.wb.sheets[token.sheet]?.kind === 'worksheet') {
      this.#saveView();
      this.activeSheetIndex = token.sheet;
      this.#loadView();
    }
    this.selection = token.selection;
  }

  /** Point an engine at the model after `replaceWorkbook`-free edits that bypass `transact` (load-time fixes). */
  recalculate(): void {
    this.calc.invalidateAll();
    this.calc.recalculateAll();
    this.version++;
  }
}

function parseA1(ref: string): CellPos | undefined {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)/i.exec(ref);
  if (!m?.[1] || !m[2]) return undefined;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return clampPos(Number(m[2]), col);
}

export { rangeOf };
