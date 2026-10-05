// UI state and user-level actions on top of the document store: the in-cell
// editor (Excel's Enter / Edit / Point modes), keyboard navigation, the
// clipboard marquee, dialogs and context menus.

import type { Cell } from '@office-kit/xlsx/cell';
import { getCellDisplayText, isDateFormat } from '@office-kit/xlsx/styles';
import { getColumnDimension, getRowDimension, type DataValidation, type DataValidationErrorStyle, type Worksheet } from '@office-kit/xlsx/worksheet';
import { formulaReferences, parseFormula, toggleReferenceAt, translateFormula } from '../calc/index.ts';
import { i18n, type MessageKey } from '../i18n/i18n.svelte.ts';
import type { CellPos, Range } from './address.ts';
import { cellAddress, clampPos, MAX_COL, MAX_ROW, quoteSheetName, rangeAddress, rangeOf } from './address.ts';
import { getCellAt } from './cells.ts';
import { commitInput, frozenCounts, type ClearKind } from './commands.ts';
import { clear, fillHandleDoubleClick, nowText, resizeColumnsPx, resizeRowsPx } from './actions.ts';
import { pickListItems, validationAt } from './data.ts';
import { toggleAutoFilter } from './filter.ts';
import { cellTextWidth } from './autofit.ts';
import { buildOutline } from './outline.ts';
import { buildPageLayout } from './page-layout.ts';
import type { FillMode } from './actions.ts';
import { addName, selectionRefText } from './names.ts';
import { validateValue } from './validation.ts';
import { isCellLocked, protectionGuard } from './protection.ts';
import { applyCompletion, completionAt, rewriteReference, type Completion } from './formula-assist.ts';
import { SpreadsheetEditor } from './editor.svelte.ts';
import { editTextFor, parseInput } from './input.ts';
import { currentRegion, dataEdge, lastUsedCell } from './navigation.ts';
import { addRange, currentRange, cycleActive, extendTo, isMultiCell, selectRange, singleCell, wholeColumns, wholeRows, type Selection } from './selection.ts';
import { GridGeometry } from '../grid/geometry.ts';
import { sparklineIndex } from './sparklines.ts';
import { pivotAt } from './pivot.ts';
import type { ShapeTool } from './shapes.ts';

export type EditMode = 'enter' | 'edit' | 'point';

export interface EditState {
  readonly row: number;
  readonly col: number;
  /** Sheet the edit started on; point mode may browse other sheets. */
  readonly sheetIndex: number;
  text: string;
  selStart: number;
  selEnd: number;
  mode: EditMode;
  /** Where the caret lives: the in-cell editor or the formula bar. */
  source: 'cell' | 'bar';
  /** The reference span being driven by arrow keys / mouse in point mode. */
  point: { start: number; end: number; anchor: CellPos; cursor: CellPos } | null;
}

export type DialogKind =
  | 'formatCells'
  | 'find'
  | 'replace'
  | 'goto'
  | 'gotoSpecial'
  | 'pasteSpecial'
  | 'insertCells'
  | 'deleteCells'
  | 'columnWidth'
  | 'rowHeight'
  | 'standardWidth'
  | 'renameSheet'
  | 'moveCopySheet'
  | 'sort'
  | 'dataValidation'
  | 'conditionalFormatting'
  | 'nameManager'
  | 'defineName'
  | 'insertFunction'
  | 'hyperlink'
  | 'note'
  | 'createTable'
  | 'pageSetup'
  | 'zoom'
  | 'unhideSheet'
  | 'insertChart'
  | 'protectSheet'
  | 'protectWorkbook'
  | 'evaluateFormula'
  | 'symbol'
  | 'workbookStatistics'
  | 'unprotect'
  | 'removeDuplicates'
  | 'textToColumns'
  | 'series'
  | 'chartSelectData'
  | 'moveChart'
  | 'altText'
  | 'createSparklines'
  | 'createPivotTable'
  | 'resizeTable'
  | 'subtotal'
  | 'consolidate'
  | 'advancedFilter'
  | 'goalSeek'
  | 'scenarioManager'
  | 'dataTable'
  | 'alert';

export interface DialogState {
  readonly kind: DialogKind;
  readonly props?: Record<string, unknown>;
}

export interface EditorHost {
  open(): void;
  save(): void;
  saveAs(): void;
  newWorkbook(): void;
  print(): void;
}

export interface MenuState {
  readonly x: number;
  readonly y: number;
  readonly kind: 'cell' | 'rowHeader' | 'colHeader' | 'sheetTab';
  readonly sheetIndex?: number;
}

const REF_COLORS = ['#3B6FD8', '#D13438', '#7A43B6', '#0F8A3B', '#B05E00', '#C239B3', '#038387'];

/** Characters after which arrow keys/clicks insert a reference instead of committing. */
const POINTABLE = /[=(,+\-*/^&<>:;%{ ]$/;

/** The corner of the active range opposite the anchor: what Shift+arrow / Shift+Page move. */
function farCorner(sel: Selection): CellPos {
  const range = currentRange(sel);
  return { row: sel.anchor.row === range.r1 ? range.r2 : range.r1, col: sel.anchor.col === range.c1 ? range.c2 : range.c1 };
}

export class EditorController {
  readonly doc: SpreadsheetEditor;
  edit = $state<EditState | null>(null);
  viewportW = $state(800);
  viewportH = $state(600);
  dialog = $state<DialogState | null>(null);
  menu = $state<MenuState | null>(null);
  /** Source of the last Copy/Cut, shown with marching ants until consumed. */
  clipboard = $state.raw<{ sheet: Worksheet; range: Range; cut: boolean } | null>(null);
  toast = $state<MessageKey | null>(null);
  /** Column a run of Tab presses started from; Enter returns there. */
  tabStartCol: number | null = null;
  /** F4 / Cmd+Y "repeat last action". */
  repeatable: (() => void) | null = null;
  ribbonTab = $state('home');
  ribbonCollapsed = $state(false);
  showFormulaBar = $state(true);
  formulaBarExpanded = $state(false);
  /** Fill-handle / move drag preview outline. */
  dragPreview = $state.raw<Range | null>(null);
  /** Index into the active sheet's `drawing.items` of the selected chart or picture; drives the contextual ribbon tabs. */
  selectedDrawing = $state<number | null>(null);
  /** Insert ▸ Shapes / Text Box armed: the next drag (or click) on the grid draws this. */
  shapeTool = $state.raw<ShapeTool | null>(null);
  /** Drawing index of the shape whose text is being typed in place (double-click, or a new text box). */
  shapeTextEdit = $state<number | null>(null);
  /** Trace Precedents/Dependents outlines. */
  traces = $state.raw<Array<{ from: CellPos; range: Range; kind: 'precedents' | 'dependents' }>>([]);
  /** Auto Fill Options button after a fill-handle drag; valid while the document is unchanged since. */
  fillOptions = $state.raw<{ source: Range; target: Range; mode: FillMode; version: number } | null>(null);
  /** Formulas ▸ Watch Window: open state and the watched cells (per session, as in Excel). */
  watchWindow = $state(false);
  watches = $state.raw<Array<{ sheet: Worksheet; row: number; col: number }>>([]);
  /**
   * Bumped to hand keyboard focus back to the grid after a control that
   * changes neither the edit nor a dialog (Name Box, font boxes, number
   * format list), so typing goes into the cell as in Excel instead of into
   * <body> or the control.
   */
  gridFocusRequest = $state(0);
  /** View ▸ Focus Cell highlight colour; a per-user view preference, so not saved in the file. */
  focusCell = $state<string | null>(null);
  /** Formula AutoComplete: highlighted row, and the text it was dismissed at (Escape). */
  completionIndex = $state(0);
  completionClosedFor = $state<string | null>(null);
  readonly completion = $derived.by((): Completion | undefined => {
    const e = this.edit;
    if (!e || e.selStart !== e.selEnd || e.text === this.completionClosedFor) return undefined;
    const names = this.doc.wb.definedNames.filter((d) => !d.hidden && !d.name.startsWith('_xlnm.')).map((d) => d.name);
    return completionAt(e.text, e.selEnd, names);
  });

  /** Arrow keys pick, Tab inserts, Escape closes the AutoComplete list. Returns true when the key was used. */
  handleCompletionKey(ev: KeyboardEvent): boolean {
    const c = this.completion;
    const e = this.edit;
    if (!c || !e) return false;
    const n = c.items.length;
    if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') {
      this.completionIndex = (this.completionIndex + (ev.key === 'ArrowDown' ? 1 : n - 1)) % n;
    } else if (ev.key === 'Tab') {
      this.acceptCompletion(this.completionIndex);
    } else if (ev.key === 'Escape') {
      this.completionClosedFor = e.text;
    } else return false;
    ev.preventDefault();
    return true;
  }

  acceptCompletion(index: number): void {
    const c = this.completion;
    const e = this.edit;
    const item = c?.items[index];
    if (!c || !e || !item) return;
    const next = applyCompletion(e.text, c, item);
    this.setEditText(next.text, next.caret);
    if (e.mode === 'enter' || e.mode === 'point') e.mode = 'enter';
  }

  /** Data-validation error alert for the entry being committed. */
  validationPrompt = $state.raw<{ style: DataValidationErrorStyle; title?: string; message?: string } | null>(null);
  /** Cell under the mouse (for note pop-ups). */
  hoverCell = $state.raw<CellPos | null>(null);
  /** Cell whose threaded-comment card is pinned open (a hover only previews it). */
  commentCard = $state.raw<CellPos | null>(null);
  /** Review ▸ Show Comments: the docked pane listing every thread. */
  commentsPane = $state(false);
  /**
   * Index of the chart sheet on screen, or null when the grid is. The document
   * keeps its last worksheet active underneath, since every edit targets one.
   */
  chartsheet = $state<number | null>(null);
  /** `chartsheet` while it still names a chart sheet; a deleted or moved one falls back to the grid. */
  shownChartsheet = $derived.by(() => {
    void this.doc.version;
    const i = this.chartsheet;
    return i !== null && this.doc.wb.sheets[i]?.kind === 'chartsheet' ? i : null;
  });
  /** View ▸ Navigation pane. */
  navigationPane = $state(false);
  /** Open AutoFilter drop-down: the header cell whose button was clicked. */
  filterMenu = $state.raw<CellPos | null>(null);
  /** Circle Invalid Data marks. */
  invalidCircles = $state.raw<Range[]>([]);
  /** File operations provided by the hosting page (open/save dialogs, printing). */
  host: EditorHost | null = null;
  /** Drop-down list under the active cell (data-validation list / AutoComplete pick list). */
  pickList = $state<{ row: number; col: number; items: string[] } | null>(null);
  findState = $state({ query: '', replace: '', matchCase: false, wholeCell: false, byColumns: false, inWorkbook: false, lookIn: 'formulas' as 'formulas' | 'values' });

  constructor(doc?: SpreadsheetEditor) {
    this.doc = doc ?? new SpreadsheetEditor();
    const d = this.doc;
    d.guard = protectionGuard(() => d.wb);
    d.onRefused = (reason) => {
      this.dialog = { kind: 'alert', props: { message: reason === 'protectedCell' ? 'protectedCellAlert' : 'protectedStructureAlert' } };
    };
  }

  // ---- derived view ---------------------------------------------------------

  /** PivotTable Analyze ▸ Field List: the PivotTable Fields pane shows while the active cell is in a pivot. */
  pivotFieldList = $state(true);
  /** The PivotTable under the active cell; drives the field pane and the PivotTable tabs. */
  activePivot = $derived.by(() => {
    void this.doc.version;
    const { row, col } = this.doc.selection.active;
    return pivotAt(this.doc.wb, this.doc.ws, row, col);
  });

  /** Sparklines on the active sheet keyed `row:col`; drives the grid painter and the Sparkline tab. */
  sparklines = $derived.by(() => {
    void this.doc.version;
    return sparklineIndex(this.doc.ws);
  });

  frozen = $derived.by(() => {
    void this.doc.version;
    return frozenCounts(this.doc.ws);
  });

  sheetView = $derived.by(() => {
    void this.doc.version;
    // View flags are mutated in place; a fresh object lets the flags derived from it update.
    const view = this.doc.ws.views[0];
    return view ? { ...view } : undefined;
  });

  /** Page Layout view's paged axes, or undefined in the other views. */
  pageLayout = $derived.by(() => {
    void this.doc.layoutVersion;
    return this.sheetView?.view === 'pageLayout' ? buildPageLayout(this.doc) : undefined;
  });

  outline = $derived.by(() => {
    void this.doc.layoutVersion;
    return { rows: buildOutline(this.doc.ws, 'row'), cols: buildOutline(this.doc.ws, 'col') };
  });

  showGridlines = $derived(this.sheetView?.showGridLines !== false);
  showHeaders = $derived(this.sheetView?.showRowColHeaders !== false);
  showFormulas = $derived(this.sheetView?.showFormulas === true);
  showOutlineSymbols = $derived(this.sheetView?.showOutlineSymbols !== false);

  geometry = $derived.by(
    () =>
      new GridGeometry({
        cols: this.pageLayout?.cols ?? this.doc.cols,
        rows: this.pageLayout?.rows ?? this.doc.rows,
        zoom: this.doc.zoom,
        scrollX: this.doc.scrollX,
        scrollY: this.doc.scrollY,
        // Page Layout view has no frozen panes (Excel ignores them there too).
        frozenRows: this.pageLayout ? 0 : this.frozen.rows,
        frozenCols: this.pageLayout ? 0 : this.frozen.cols,
        width: this.viewportW,
        height: this.viewportH,
        showHeaders: this.showHeaders,
        rowOutlineLevels: this.showOutlineSymbols ? this.outline.rows.maxLevel : 0,
        colOutlineLevels: this.showOutlineSymbols ? this.outline.cols.maxLevel : 0,
      }),
  );

  /** Coloured references of the formula being edited, for the grid outlines and the editor text. */
  editRefs = $derived.by(() => {
    const e = this.edit;
    if (!e || !e.text.startsWith('=')) return [];
    const sheetName = this.doc.wb.sheets[this.doc.activeSheetIndex]?.sheet.title ?? '';
    const editSheet = this.doc.wb.sheets[e.sheetIndex]?.sheet.title ?? '';
    let refs: ReturnType<typeof formulaReferences>;
    try {
      refs = formulaReferences(e.text.slice(1));
    } catch {
      return [];
    }
    return refs.map((ref, i) => ({
      start: ref.start + 1,
      end: ref.end + 1,
      color: REF_COLORS[i % REF_COLORS.length] ?? '#3B6FD8',
      range: { r1: ref.range.r1, c1: ref.range.c1, r2: Math.min(ref.range.r2, MAX_ROW), c2: Math.min(ref.range.c2, MAX_COL) },
      visible: (ref.sheet ?? editSheet) === sheetName,
    }));
  });

  // ---- helpers ------------------------------------------------------------

  get ws() {
    return this.doc.ws;
  }

  cell(row: number, col: number): Cell | undefined {
    return getCellAt(this.doc.ws, row, col);
  }

  defaultStyleAt = (row: number, col: number): number =>
    getRowDimension(this.doc.ws, row)?.style ?? getColumnDimension(this.doc.ws, col)?.style ?? 0;

  dateOrder(): 'mdy' | 'ymd' {
    return i18n.locale === 'ja' ? 'ymd' : 'mdy';
  }

  /** Bring a cell into view in the scrolling pane. */
  reveal(row: number, col: number): void {
    const { x, y } = this.geometry.scrollToReveal(row, col);
    if (x !== this.doc.scrollX || y !== this.doc.scrollY) this.doc.setScroll(x, y);
  }

  // ---- selection ------------------------------------------------------------

  /** Normalise a position onto a visible cell, snapping into merges. */
  #snap(pos: CellPos): CellPos {
    const merge = this.doc.merges.at(pos.row, pos.col);
    return merge ? { row: merge.r1, col: merge.c1 } : pos;
  }

  selectCell(pos: CellPos, opts: { extend?: boolean; add?: boolean } = {}): void {
    const p = clampPos(pos.row, pos.col);
    const sel = this.doc.selection;
    if (opts.extend) {
      const next = extendTo(sel, p);
      const r = this.doc.merges.expand(currentRange(next));
      const ranges = next.ranges.slice();
      ranges[next.activeRange] = r;
      this.doc.setSelection({ ...next, ranges });
    } else if (opts.add) {
      this.doc.setSelection(addRange(sel, this.#snap(p)));
    } else {
      const s = this.#snap(p);
      const merge = this.doc.merges.at(s.row, s.col);
      this.doc.setSelection(merge ? { ...singleCell(s), ranges: [merge] } : singleCell(s));
    }
    this.tabStartCol = null;
  }

  selectRange(range: Range, active?: CellPos): void {
    this.doc.setSelection(selectRange(range, active));
  }

  /** Arrow-key movement, with Shift extending and Ctrl/Cmd jumping to data edges. */
  move(dRow: -1 | 0 | 1, dCol: -1 | 0 | 1, opts: { extend?: boolean; jump?: boolean } = {}): void {
    const sel = this.doc.selection;
    const from: CellPos = opts.extend ? farCorner(sel) : sel.active;
    let to: CellPos;
    if (opts.jump) {
      to = dataEdge(this.doc.ws, from, dRow, dCol, (i) => (dRow !== 0 ? this.doc.rows : this.doc.cols).isHidden(i));
    } else {
      // Step out of a merged block from its far side.
      const merge = this.doc.merges.at(from.row, from.col);
      const baseRow = merge ? (dRow > 0 ? merge.r2 : merge.r1) : from.row;
      const baseCol = merge ? (dCol > 0 ? merge.c2 : merge.c1) : from.col;
      to = {
        row: dRow === 0 ? from.row : this.doc.rows.nextVisible(baseRow, dRow),
        col: dCol === 0 ? from.col : this.doc.cols.nextVisible(baseCol, dCol),
      };
    }
    if (opts.extend) this.selectCell(to, { extend: true });
    else this.selectCell(to);
    this.reveal(to.row, to.col);
  }

  /** More than one cell is selected; a selected merged cell counts as one. */
  #spansSeveralCells(sel: Selection): boolean {
    if (!isMultiCell(sel)) return false;
    const merge = this.doc.merges.at(sel.active.row, sel.active.col);
    const r = sel.ranges[0];
    return !(sel.ranges.length === 1 && merge && r && merge.r1 === r.r1 && merge.c1 === r.c1 && merge.r2 === r.r2 && merge.c2 === r.c2);
  }

  /** Enter / Tab movement: cycles within a multi-cell selection, otherwise moves one cell. */
  advance(direction: 'down' | 'up' | 'right' | 'left'): void {
    const sel = this.doc.selection;
    if (this.#spansSeveralCells(sel)) {
      let next = cycleActive(sel, direction);
      // Excel stops only on a merged cell's top-left cell, never inside it.
      for (let guard = 0; guard < 16_384; guard++) {
        const merge = this.doc.merges.at(next.active.row, next.active.col);
        if (!merge || (merge.r1 === next.active.row && merge.c1 === next.active.col)) break;
        next = cycleActive(next, direction);
      }
      this.doc.setSelection(next);
      this.reveal(next.active.row, next.active.col);
      return;
    }
    if (direction === 'right' || direction === 'left') {
      if (this.tabStartCol === null) this.tabStartCol = sel.active.col;
      const startCol = this.tabStartCol;
      this.move(0, direction === 'right' ? 1 : -1);
      this.tabStartCol = startCol;
      return;
    }
    if (direction === 'down' && this.tabStartCol !== null) {
      const col = this.tabStartCol;
      const row = this.doc.rows.nextVisible(sel.active.row, 1);
      this.selectCell({ row, col });
      this.reveal(row, col);
      return;
    }
    this.move(direction === 'down' ? 1 : -1, 0);
  }

  pageMove(direction: 1 | -1, horizontal: boolean, extend: boolean): void {
    const geo = this.geometry;
    // Shift+Page keeps extending from the moving corner, page after page.
    const from = extend ? farCorner(this.doc.selection) : this.doc.selection.active;
    if (horizontal) {
      const [c1, c2] = geo.mainCols();
      const span = Math.max(1, c2 - c1);
      const col = Math.min(MAX_COL, Math.max(1, from.col + direction * span));
      this.doc.setScroll(Math.max(0, this.doc.scrollX + direction * (geo.colX(c2) - geo.colX(c1))), this.doc.scrollY);
      this.selectCell({ row: from.row, col }, { extend });
      return;
    }
    const [r1, r2] = geo.mainRows();
    const span = Math.max(1, r2 - r1);
    const row = Math.min(MAX_ROW, Math.max(1, from.row + direction * span));
    this.doc.setScroll(this.doc.scrollX, Math.max(0, this.doc.scrollY + direction * (geo.rowY(r2) - geo.rowY(r1))));
    this.selectCell({ row, col: from.col }, { extend });
  }

  home(ctrl: boolean, extend: boolean): void {
    const sel = this.doc.selection;
    // Home lands on the first column (and Ctrl+Home the first row) that is shown, as in Excel.
    const firstShown = (axis: { isHidden(i: number): boolean }, from: number, limit: number): number => {
      let i = from;
      while (i < limit && axis.isHidden(i)) i++;
      return i;
    };
    const col = firstShown(this.doc.cols, ctrl ? this.frozen.cols + 1 : 1, MAX_COL);
    const target = ctrl ? { row: firstShown(this.doc.rows, this.frozen.rows + 1, MAX_ROW), col } : { row: sel.active.row, col };
    this.selectCell(target, { extend });
    this.reveal(target.row, target.col);
  }

  end(extend: boolean): void {
    const target = lastUsedCell(this.doc.ws);
    this.selectCell(target, { extend });
    this.reveal(target.row, target.col);
  }

  /** Cmd/Ctrl+A: current region first, the whole sheet on the second press. */
  selectAll(): void {
    const sel = this.doc.selection;
    const region = currentRegion(this.doc.ws, sel.active);
    const cur = currentRange(sel);
    const isRegion = region.r1 === cur.r1 && region.r2 === cur.r2 && region.c1 === cur.c1 && region.c2 === cur.c2;
    const single = region.r1 === region.r2 && region.c1 === region.c2;
    if (!isRegion && !single) this.selectRange(region, sel.active);
    else this.selectRange({ r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL }, sel.active);
  }

  /** Ctrl+. : move the active cell clockwise to the next corner of the selection. */
  nextCorner(): void {
    const sel = this.doc.selection;
    const r = currentRange(sel);
    const { row, col } = sel.active;
    const corners: CellPos[] = [
      { row: r.r1, col: r.c1 },
      { row: r.r1, col: r.c2 },
      { row: r.r2, col: r.c2 },
      { row: r.r2, col: r.c1 },
    ];
    const i = corners.findIndex((p) => p.row === row && p.col === col);
    const next = corners[(i + 1) % corners.length] ?? corners[0];
    if (!next) return;
    this.doc.setSelection({ ...sel, active: next });
    this.reveal(next.row, next.col);
  }

  selectEntire(axis: 'rows' | 'cols'): void {
    const sel = this.doc.selection;
    const r = currentRange(sel);
    const range = axis === 'cols' ? wholeColumns(r.c1, r.c2) : wholeRows(r.r1, r.r2);
    this.doc.setSelection({ ...sel, ranges: [range], activeRange: 0 });
  }

  // ---- in-cell editing ------------------------------------------------------

  /** Begin editing the active cell. `initial` replaces its content (typing); otherwise the content is loaded (F2). */
  startEdit(initial?: string, source: 'cell' | 'bar' = 'cell'): void {
    const { row, col } = this.doc.selection.active;
    // Excel refuses to even open the editor on a locked cell of a protected sheet.
    if (this.doc.ws.sheetProtection?.sheet && isCellLocked(this.doc.wb, this.doc.ws, row, col)) {
      this.doc.onRefused?.('protectedCell');
      return;
    }
    let text: string;
    let mode: EditMode;
    if (initial !== undefined) {
      text = initial;
      mode = 'enter';
    } else {
      const cell = this.cell(row, col);
      const style = this.doc.styles.get(cell?.styleId ?? this.defaultStyleAt(row, col));
      const formatted = cell ? this.displayText(cell) : '';
      text = cell ? editTextFor(cell.value, formatted, isDateFormat(style.numFmt)) : '';
      mode = 'edit';
    }
    this.edit = { row, col, sheetIndex: this.doc.activeSheetIndex, text, selStart: text.length, selEnd: text.length, mode, source, point: null };
    this.reveal(row, col);
  }

  displayText(cell: Cell): string {
    return getCellDisplayText(this.doc.wb, cell);
  }

  setEditText(text: string, selStart: number, selEnd: number = selStart): void {
    const e = this.edit;
    if (!e) return;
    if (e.text !== text) this.completionIndex = 0;
    e.text = text;
    e.selStart = selStart;
    e.selEnd = selEnd;
    e.point = null;
  }

  /**
   * Whether arrow keys / clicks should insert a reference at the caret (Excel's
   * Point mode). In Edit mode (F2, or typing in the formula bar) the arrow keys
   * move the caret instead, but a click on a cell or sheet tab still points, as
   * in Excel.
   */
  canPoint(byMouse = false): boolean {
    const e = this.edit;
    if (!e || !e.text.startsWith('=') || (e.mode === 'edit' && !byMouse)) return false;
    if (e.point) return true;
    return POINTABLE.test(e.text.slice(0, e.selStart));
  }

  /** Insert or move the point-mode reference to `pos` (extending from the anchor when `extend`). */
  pointTo(pos: CellPos, extend = false): void {
    const e = this.edit;
    if (!e) return;
    const anchor = extend && e.point ? e.point.anchor : pos;
    const range = rangeOf(anchor, pos);
    const onOtherSheet = this.doc.activeSheetIndex !== e.sheetIndex;
    const prefix = onOtherSheet ? `${quoteSheetName(this.doc.ws.title)}!` : '';
    const refText = prefix + (range.r1 === range.r2 && range.c1 === range.c2 ? cellAddress(range.r1, range.c1) : rangeAddress(range));
    const start = e.point ? e.point.start : e.selStart;
    const end = e.point ? e.point.end : e.selEnd;
    e.text = e.text.slice(0, start) + refText + e.text.slice(end);
    e.point = { start, end: start + refText.length, anchor, cursor: pos };
    e.selStart = e.selEnd = start + refText.length;
    e.mode = 'point';
    this.selectCell(pos, { extend: false });
    if (extend) this.doc.setSelection({ ...this.doc.selection, ranges: [range] });
    this.reveal(pos.row, pos.col);
  }

  /** Range finder: retarget the `index`-th coloured reference of the edited formula. */
  retargetEditRef(index: number, range: Range): void {
    const e = this.edit;
    const ref = this.editRefs[index];
    if (!e || !ref) return;
    const text = rewriteReference(e.text.slice(ref.start, ref.end), range);
    if (!text) return;
    this.setEditText(e.text.slice(0, ref.start) + text + e.text.slice(ref.end), ref.start + text.length);
  }

  pointMove(dRow: number, dCol: number, extend: boolean): void {
    const e = this.edit;
    if (!e) return;
    const from = e.point?.cursor ?? { row: e.row, col: e.col };
    const to = clampPos(from.row + dRow, from.col + dCol);
    this.pointTo(to, extend);
  }

  /** F4 while editing: cycle the anchoring of the reference at the caret. */
  toggleAbsolute(): void {
    const e = this.edit;
    if (!e || !e.text.startsWith('=')) return;
    const res = toggleReferenceAt(e.text.slice(1), Math.max(0, e.selStart - 1));
    e.text = `=${res.text}`;
    e.selStart = e.selEnd = res.caret + 1;
    e.point = null;
  }

  cancelEdit(): void {
    const e = this.edit;
    this.edit = null;
    if (e && e.sheetIndex !== this.doc.activeSheetIndex) this.doc.activateSheet(e.sheetIndex);
    if (e) this.selectCell({ row: e.row, col: e.col });
  }

  /**
   * Commit the edit. Returns false (keeping the editor open) when the formula
   * does not parse — Excel refuses to leave the cell in that case.
   */
  commitEdit(opts: { fillSelection?: boolean; skipValidation?: boolean } = {}): boolean {
    const e = this.edit;
    if (!e) return true;
    let text = e.text;
    if (text.startsWith('=') && text.length > 1) {
      text = autoCloseParens(text);
      try {
        parseFormula(text.slice(1));
      } catch (err) {
        this.dialog = { kind: 'alert', props: { message: 'formulaError', detail: err instanceof Error ? err.message : String(err) } };
        return false;
      }
    }
    if (!opts.skipValidation) {
      const rule = this.#violatedRule(e.sheetIndex, e.row, e.col, text);
      if (rule) {
        this.validationPrompt = {
          style: rule.errorStyle ?? 'stop',
          ...(rule.errorTitle ? { title: rule.errorTitle } : {}),
          ...(rule.error ? { message: rule.error } : {}),
        };
        return false;
      }
    }
    if (e.sheetIndex !== this.doc.activeSheetIndex) this.doc.activateSheet(e.sheetIndex);
    const at = { row: e.row, col: e.col };
    const selection = this.doc.selection;
    const several = this.#spansSeveralCells(selection);
    const fill = opts.fillSelection && several ? { ranges: selection.ranges, translate: translateFormula } : undefined;
    if (this.splitsArray(fill?.ranges ?? [{ r1: at.row, c1: at.col, r2: at.row, c2: at.col }])) return false;
    this.edit = null;
    commitInput(this.doc, at, text, { dateOrder: this.dateOrder(), measure: (cell) => cellTextWidth(this, cell) }, fill);
    if (!fill && !several) {
      // Re-selecting the edited cell must not end a Tab run: Enter returns to its first column.
      const tabStartCol = this.tabStartCol;
      this.selectCell(at);
      this.tabStartCol = tabStartCol;
    }
    // F4 / Ctrl+Y repeats the entry into the active cell (Excel's "Repeat Typing").
    this.repeatable = fill
      ? null
      : () => commitInput(this.doc, this.doc.selection.active, text, { dateOrder: this.dateOrder(), measure: (cell) => cellTextWidth(this, cell) });
    return true;
  }

  /** Excel refuses to change part of a legacy array; says so and returns true when `ranges` would. */
  splitsArray(ranges: readonly Range[]): boolean {
    const sheet = this.doc.ws.title;
    if (!ranges.some((r) => this.doc.calc.splitsFixedArray(sheet, r))) return false;
    this.dialog = { kind: 'alert', props: { message: 'partOfArray' } };
    return true;
  }

  /** The data-validation rule `text` would break at (row, col), if its error alert is on. */
  #violatedRule(sheetIndex: number, row: number, col: number, text: string): DataValidation | undefined {
    const ref = this.doc.wb.sheets[sheetIndex];
    if (ref?.kind !== 'worksheet') return undefined;
    const dv = validationAt(ref.sheet, row, col);
    if (!dv || dv.showErrorMessage === false) return undefined;
    const parsed = parseInput(text, { dateOrder: this.dateOrder(), date1904: this.doc.wb.date1904 });
    let value = parsed.value;
    if (value !== null && typeof value === 'object' && !(value instanceof Date) && value.kind === 'formula') {
      // A formula is checked by its result, as Excel does.
      const result = this.doc.calc.evaluate(value.formula, ref.sheet.title, row, col);
      value =
        typeof result === 'number' || typeof result === 'string' || typeof result === 'boolean'
          ? result
          : result !== null && result.kind === 'error'
            ? { kind: 'error', code: result.code }
            : null;
    }
    return validateValue(this, dv, value, row, col) ? undefined : dv;
  }

  /** Answer to the validation alert: commit anyway, keep editing, or drop the entry. */
  resolveValidation(choice: 'commit' | 'retry' | 'cancel'): void {
    this.validationPrompt = null;
    if (choice === 'commit') this.commitEdit({ skipValidation: true });
    else if (choice === 'cancel') this.cancelEdit();
  }

  // ---- thin wrappers used by the grid -----------------------------------------

  openDialog(kind: DialogKind, props?: Record<string, unknown>): void {
    if (this.edit && !this.commitEdit()) return;
    this.menu = null;
    this.dialog = props ? { kind, props } : { kind };
  }

  /** Name Box: typing an unused name defines it for the current selection. */
  defineNameForSelection(name: string): void {
    const err = addName(this, { name, value: selectionRefText(this) });
    if (err) this.dialog = { kind: 'alert', props: { message: err } };
  }

  closeDialog(): void {
    this.dialog = null;
  }

  toggleFilter(): void {
    toggleAutoFilter(this);
  }

  /** Option/Alt+Down: the data-validation list for the cell, else the column's distinct texts. */
  openPickList(): void {
    const { row, col } = this.doc.selection.active;
    const items = pickListItems(this, row, col);
    if (items.length > 0) this.pickList = { row, col, items };
  }

  resizeColumns(cols: readonly number[], px: number): void {
    resizeColumnsPx(this, cols, px);
  }

  resizeRows(rows: readonly number[], px: number): void {
    resizeRowsPx(this, rows, px);
  }

  fillHandleDoubleClick(): void {
    fillHandleDoubleClick(this);
  }

  clearSelection(kind: ClearKind): void {
    clear(this, kind);
  }

  nowText(kind: 'date' | 'time'): string {
    return nowText(kind, this.dateOrder() === 'ymd');
  }

  /** Ctrl/Cmd+PageDown / PageUp: next/previous visible worksheet. */
  /** Bring sheet `index` on screen: a worksheet becomes the active one, a chart sheet is shown over it. */
  showSheet(index: number): void {
    const ref = this.doc.wb.sheets[index];
    if (ref?.kind === 'chartsheet') {
      this.chartsheet = index;
      return;
    }
    this.chartsheet = null;
    this.doc.activateSheet(index);
  }

  switchSheet(step: 1 | -1): void {
    const sheets = this.doc.wb.sheets;
    let i = this.shownChartsheet ?? this.doc.activeSheetIndex;
    for (let n = 0; n < sheets.length; n++) {
      i = (i + step + sheets.length) % sheets.length;
      const ref = sheets[i];
      if (ref?.state === 'visible') {
        if (this.edit && !this.canPoint()) this.commitEdit();
        this.showSheet(i);
        return;
      }
    }
  }

  openContextMenuAtActive(): void {
    const { row, col } = this.doc.selection.active;
    const geo = this.geometry;
    const host = typeof document === 'undefined' ? null : document.querySelector('.grid-host');
    const rect = host?.getBoundingClientRect();
    this.menu = { x: (rect?.left ?? 0) + geo.colX(col) + geo.colW(col) / 2, y: (rect?.top ?? 0) + geo.rowY(row) + geo.rowH(row), kind: 'cell' };
  }
}

/** Excel adds the closing parentheses a formula is missing when it is committed. */
function autoCloseParens(text: string): string {
  let depth = 0;
  let quoted = false;
  for (const ch of text) {
    if (ch === '"') quoted = !quoted;
    else if (!quoted && ch === '(') depth++;
    else if (!quoted && ch === ')') depth--;
  }
  return depth > 0 && !quoted ? text + ')'.repeat(depth) : text;
}
