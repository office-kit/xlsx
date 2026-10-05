<script lang="ts">
  // The worksheet surface: a canvas painted from the visible window, a native
  // scroll container whose spacer grows as the user scrolls (Excel's
  // scrollbars track the used range plus the current view, not 1,048,576
  // rows), and one textarea that both receives keystrokes while navigating
  // and becomes the in-cell editor. Keeping a real text field focused at the
  // active cell is what lets an IME composition start in place.
  import { onMount, untrack } from 'svelte';
  import { headerFooterFor, parseHeader, renderHf } from '../core/header-footer.ts';
  import { getEditor } from '../core/context.ts';
  import { MAX_COL, MAX_ROW, rangeOf, type CellPos, type Range } from '../core/address.ts';
  import { paintGrid, LIGHT_THEME, type CommentMark, type PaperView } from './paint.ts';
  import { commentThreads, refCell } from '../core/comments.ts';
  import { canvasFont } from '../core/render-style.ts';
  import { handleGridKey } from './keyboard.ts';
  import { usedRange } from '../core/cells.ts';
  import { currentRange, selectionContains } from '../core/selection.ts';
  import { formatPainter } from '../core/format-painter.svelte.ts';
  import CellEditorText from './CellEditorText.svelte';
  import { outlineButtonAt, outlineLayout } from './outline-layout.ts';
  import { activeCriteria, filterOwners } from '../core/filter.ts';
  import { paginate, printArea } from '../core/pages.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { showLevel, toggleRun } from '../core/outline.ts';
  import { followHyperlink, hyperlinkAt } from '../core/hyperlinks.ts';
  import GridOverlays from './GridOverlays.svelte';
  import TotalRowButton from './TotalRowButton.svelte';
  import DrawingLayer from './DrawingLayer.svelte';
  import ShapeDrawOverlay from './ShapeDrawOverlay.svelte';
  import FormulaAssist from './FormulaAssist.svelte';
  import { startFillDrag, finishFillDrag, previewFill } from './gestures.ts';
  import { autofitColumns, autofitRows } from '../core/autofit.ts';
  import { dragSelection, pasteFromEvent, writeCopyEvent } from '../core/clipboard.ts';
  import { conditionalOverlay } from '../core/conditional-format.ts';
  import { tableLooks } from '../core/table-style.ts';
  import { sparklineValues } from '../core/sparklines.ts';
  import type { SparklinePaintCell } from './sparkline-paint.ts';

  const ctl = getEditor();
  const doc = ctl.doc;

  let host: HTMLDivElement;
  let scroller: HTMLDivElement;
  let canvas: HTMLCanvasElement;
  let input: HTMLTextAreaElement;
  let dpr = $state(typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1);
  let antsPhase = $state(0);
  let composing = false;
  // The textarea's text while an IME composes. The model is not updated until
  // the composition ends (rewriting the textarea would cancel it), but the
  // overlay must show what is being typed, as Excel does.
  let composingText: string | null = $state(null);

  const geo = $derived(ctl.geometry);

  // Spacer extent: the used range or the current view (whichever is larger),
  // plus one more screen, so the scrollbar keeps growing like Excel's.
  // The used range only changes with the model; kept apart from the
  // selection / scroll reads below so moving around does not rescan the sheet.
  const used = $derived.by(() => {
    void doc.version;
    return usedRange(doc.ws);
  });
  const extent = $derived.by(() => {
    const sel = doc.selection.active;
    const lastRow = Math.min(MAX_ROW, Math.max(used?.r2 ?? 1, sel.row) + 1);
    const lastCol = Math.min(MAX_COL, Math.max(used?.c2 ?? 1, sel.col) + 1);
    const g = geo;
    const w = Math.max(g.headerW + g.cols.offsetOf(lastCol + 1) * doc.zoom, doc.scrollX + g.width * 2);
    const h = Math.max(g.headerH + g.rows.offsetOf(lastRow + 1) * doc.zoom, doc.scrollY + g.height * 2);
    // Browsers cap element sizes around 33M px; the last row still has to be reachable by keyboard.
    return { w: Math.min(w, 30_000_000), h: Math.min(h, 30_000_000) };
  });

  const overlay = $derived.by(() => {
    void doc.version;
    return conditionalOverlay(doc.wb, doc.ws, doc.calc, doc.styles);
  });

  const tableLook = $derived.by(() => {
    void doc.version;
    return tableLooks(doc.ws.tables, doc.styles.palette);
  });

  const filteredRows = $derived.by(() => {
    void doc.version;
    return filterOwners(doc.ws)
      .filter((o) => activeCriteria(o).size > 0)
      .map((o) => [o.range.r1 + 1, o.range.r2] as const);
  });

  const paper = $derived.by((): PaperView | undefined => {
    const layout = ctl.pageLayout;
    if (!layout) return undefined;
    void doc.version;
    const area = printArea(doc);
    // Only pages that hold part of the printed area get a number (and a header / footer).
    const rowPages = area ? layout.rows.pageOf(area.r2) + 1 : 1;
    const colPages = area ? layout.cols.pageOf(area.c2) + 1 : 1;
    const overThenDown = doc.ws.pageSetup?.pageOrder === 'overThenDown';
    const hf = doc.ws.headerFooter;
    const pages = rowPages * colPages;
    const now = new Date();
    const fields = { date: now.toLocaleDateString(), time: now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }), file: doc.fileName, path: '', sheet: doc.ws.title };
    const expand = (text: string | undefined, page: number) => {
      const s = parseHeader(text);
      return { left: renderHf(s.left, fields, page, pages), center: renderHf(s.center, fields, page, pages), right: renderHf(s.right, fields, page, pages) };
    };
    return {
      cols: layout.cols,
      rows: layout.rows,
      metrics: layout.metrics,
      texts: (rp, cp) => {
        if (rp >= rowPages || cp >= colPages) return undefined;
        const page = overThenDown ? rp * colPages + cp + 1 : cp * rowPages + rp + 1;
        return { header: expand(headerFooterFor(hf, page, 'header'), page), footer: expand(headerFooterFor(hf, page, 'footer'), page) };
      },
    };
  });

  const pages = $derived.by(() => {
    void doc.version;
    void doc.layoutVersion;
    return ctl.sheetView?.view === 'pageBreakPreview' ? paginate(doc) : undefined;
  });

  const outlineView = $derived.by(() => {
    void doc.version;
    const o = ctl.outline;
    return ctl.showOutlineSymbols && (o.rows.maxLevel > 0 || o.cols.maxLevel > 0) ? outlineLayout(geo, doc.ws, o.rows, o.cols) : undefined;
  });

  // Values are read on first paint of each sparkline and kept until the document changes.
  const sparklineCells = $derived.by((): SparklinePaintCell[] => {
    void doc.version;
    const sheet = doc.ws.title;
    return [...ctl.sparklines.values()].map((s) => {
      let values: Array<number | null> | undefined;
      return { row: s.row, col: s.col, group: s.group, values: () => (values ??= sparklineValues(doc.calc, sheet, s.sparkline.formula)) };
    });
  });

  const commentMarks = $derived.by((): CommentMark[] => {
    void doc.version;
    const marks: CommentMark[] = [];
    for (const c of doc.ws.legacyComments) marks.push({ ...refCell(c.ref), kind: 'note' });
    for (const th of commentThreads(doc.ws)) marks.push({ row: th.row, col: th.col, kind: th.root.done ? 'resolved' : 'thread' });
    return marks;
  });

  let frame = 0;
  function schedule() {
    if (frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      paint();
    });
  }

  function paint() {
    const ctx = canvas?.getContext('2d');
    if (!ctx) return;
    const w = geo.width;
    const h = geo.height;
    if (canvas.width !== Math.round(w * dpr) || canvas.height !== Math.round(h * dpr)) {
      canvas.width = Math.round(w * dpr);
      canvas.height = Math.round(h * dpr);
    }
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    const e = ctl.edit;
    const editingHere = e && e.sheetIndex === doc.activeSheetIndex ? { row: e.row, col: e.col } : null;
    paintGrid({
      ctx,
      geo,
      wb: doc.wb,
      ws: doc.ws,
      styles: doc.styles,
      merges: doc.merges,
      selection: doc.selection,
      focusCell: ctl.focusCell ?? undefined,
      paper,
      showGridlines: ctl.showGridlines,
      showHeaders: ctl.showHeaders,
      showFormulas: ctl.showFormulas,
      theme: LIGHT_THEME,
      defaultStyleAt: ctl.defaultStyleAt,
      overlayAt: overlay,
      tableLookAt: tableLook,
      refHighlights: ctl.editRefs.filter((r) => r.visible).map((r) => ({ range: r.range, color: r.color })),
      copyRange: ctl.clipboard && ctl.clipboard.sheetIndex === doc.activeSheetIndex ? ctl.clipboard.range : null,
      antsPhase,
      dragPreview: ctl.dragPreview,
      editing: editingHere,
      commentMarks,
      traces: ctl.traces,
      invalidCircles: ctl.invalidCircles,
      outline: outlineView,
      filteredRows,
      pages,
      pageLabel: (n) => t('pageN', { n }),
      sparklines: sparklineCells,
    });
  }

  // Repaint whenever anything the painter reads changes.
  $effect(() => {
    void doc.version;
    void doc.layoutVersion;
    void doc.selection;
    void doc.scrollX;
    void doc.scrollY;
    void doc.zoom;
    void geo;
    void ctl.edit?.text;
    void ctl.edit?.sheetIndex;
    void ctl.clipboard;
    void ctl.dragPreview;
    void antsPhase;
    void overlay;
    void sparklineCells;
    schedule();
  });

  // Keep the native scroller in sync when the model scrolls (keyboard reveal, Go To).
  $effect(() => {
    const x = doc.scrollX;
    const y = doc.scrollY;
    untrack(() => {
      if (!scroller) return;
      if (Math.abs(scroller.scrollLeft - x) > 1) scroller.scrollLeft = x;
      if (Math.abs(scroller.scrollTop - y) > 1) scroller.scrollTop = y;
    });
  });

  // Marching ants for the copy source.
  $effect(() => {
    if (!ctl.clipboard) return;
    const id = setInterval(() => (antsPhase = (antsPhase + 1) % 14), 80);
    return () => clearInterval(id);
  });

  onMount(() => {
    const ro = new ResizeObserver(() => {
      // The scroller's client box excludes its scrollbars, so cells are never revealed underneath them.
      ctl.viewportW = Math.max(100, scroller.clientWidth);
      ctl.viewportH = Math.max(100, scroller.clientHeight);
      dpr = window.devicePixelRatio || 1;
    });
    ro.observe(host);
    focusInput();
    return () => {
      ro.disconnect();
      cancelAnimationFrame(frame);
    };
  });

  function onScroll() {
    doc.setScroll(scroller.scrollLeft, scroller.scrollTop);
  }

  // ---- in-place editor geometry ------------------------------------------------

  const activeRect = $derived.by(() => {
    const e = ctl.edit;
    const pos: CellPos = e && e.sheetIndex === doc.activeSheetIndex ? { row: e.row, col: e.col } : doc.selection.active;
    const merge = doc.merges.at(pos.row, pos.col);
    const r = geo.rectOf(merge ?? { r1: pos.row, c1: pos.col, r2: pos.row, c2: pos.col });
    return r;
  });

  const editStyle = $derived.by(() => {
    const e = ctl.edit;
    if (!e) return null;
    const cell = ctl.cell(e.row, e.col);
    return doc.styles.get(cell?.styleId ?? ctl.defaultStyleAt(e.row, e.col));
  });

  const editorVisible = $derived(!!ctl.edit && ctl.edit.source === 'cell' && ctl.edit.sheetIndex === doc.activeSheetIndex);

  function focusInput() {
    if (!input || ctl.dialog) return;
    if (document.activeElement !== input) input.focus({ preventScroll: true });
  }

  $effect(() => {
    // Return focus to the grid when a dialog/menu closes or the edit ends.
    void ctl.dialog;
    void ctl.edit;
    void ctl.gridFocusRequest;
    if (!ctl.dialog && !ctl.validationPrompt && (!ctl.edit || ctl.edit.source === 'cell')) queueMicrotask(focusInput);
  });

  // Keep the textarea's value and caret in sync with the edit state.
  $effect(() => {
    const e = ctl.edit;
    const text = e?.text ?? '';
    const selStart = e?.selStart ?? 0;
    const selEnd = e?.selEnd ?? 0;
    untrack(() => {
      if (!input || composing) return;
      if (input.value !== text) input.value = text;
      if (e && e.source === 'cell' && (input.selectionStart !== selStart || input.selectionEnd !== selEnd)) {
        input.setSelectionRange(selStart, selEnd);
      }
    });
  });

  function onInput() {
    if (composing) {
      composingText = input.value;
      return;
    }
    const value = input.value;
    if (!ctl.edit) {
      if (value === '') return;
      ctl.startEdit(value);
      return;
    }
    ctl.setEditText(value, input.selectionStart ?? value.length, input.selectionEnd ?? value.length);
  }

  function onCompositionStart() {
    composing = true;
    if (!ctl.edit) ctl.startEdit('');
  }

  function onCompositionEnd() {
    composing = false;
    composingText = null;
    onInput();
  }

  function onSelect() {
    const e = ctl.edit;
    if (!e || composing) return;
    e.selStart = input.selectionStart ?? 0;
    e.selEnd = input.selectionEnd ?? 0;
  }

  function onKeyDown(ev: KeyboardEvent) {
    if (ev.isComposing || composing) return;
    handleGridKey(ctl, ev, input);
  }

  // ---- pointer ----------------------------------------------------------------

  type Drag =
    | { kind: 'select'; extend: boolean; from: CellPos; moved: boolean }
    | { kind: 'move'; grab: { dr: number; dc: number }; size: { rows: number; cols: number }; to: CellPos | null }
    | { kind: 'point'; anchor: CellPos }
    | { kind: 'refMove'; index: number; origin: Range; grab: CellPos }
    | { kind: 'refResize'; index: number; fixed: CellPos }
    | { kind: 'colHeader'; anchor: number }
    | { kind: 'rowHeader'; anchor: number }
    | { kind: 'colResize'; col: number; startX: number; startW: number }
    | { kind: 'rowResize'; row: number; startY: number; startH: number }
    | { kind: 'fill'; source: Range };

  let drag: Drag | null = null;
  let hoverCursor = $state('cell');
  let resizeGuide = $state<{ x?: number; y?: number; label: string } | null>(null);
  let autoScrollTimer = 0;
  let lastPointer = { x: 0, y: 0 };

  function localXY(ev: PointerEvent | MouseEvent) {
    const rect = host.getBoundingClientRect();
    return { x: ev.clientX - rect.left, y: ev.clientY - rect.top };
  }

  function fillHandleHit(x: number, y: number): boolean {
    const only = doc.selection.ranges[0];
    if (doc.selection.ranges.length !== 1 || !only) return false;
    const r = geo.rectOf(only);
    const hx = r.x + r.w - 1;
    const hy = r.y + r.h - 1;
    return Math.abs(x - hx) <= 5 && Math.abs(y - hy) <= 5;
  }

  /** The selection's border (not the fill handle): dragging it moves the cells. */
  function borderHit(x: number, y: number): boolean {
    const sel = doc.selection;
    const range = sel.ranges[0];
    if (sel.ranges.length !== 1 || !range || range.r2 === MAX_ROW || range.c2 === MAX_COL || ctl.edit) return false;
    const r = geo.rectOf(range);
    const near = (a: number, b: number) => Math.abs(a - b) <= 3;
    const inX = x >= r.x - 3 && x <= r.x + r.w + 3;
    const inY = y >= r.y - 3 && y <= r.y + r.h + 3;
    return ((near(x, r.x) || near(x, r.x + r.w - 1)) && inY) || ((near(y, r.y) || near(y, r.y + r.h - 1)) && inX);
  }

  /** A coloured reference box's edge (move) or corner handle (resize) under the pointer. */
  function rangeFinderHit(x: number, y: number): { index: number; range: Range; corner?: CellPos } | undefined {
    const near = (a: number, b: number) => Math.abs(a - b) <= 4;
    for (const [index, ref] of ctl.editRefs.entries()) {
      if (!ref.visible || ref.range.r2 === MAX_ROW || ref.range.c2 === MAX_COL) continue;
      const r = geo.rectOf(ref.range);
      const left = r.x - 1;
      const right = r.x + r.w - 1;
      const top = r.y - 1;
      const bottom = r.y + r.h - 1;
      for (const [cx, cy, row, col] of [
        [left, top, ref.range.r1, ref.range.c1],
        [right, top, ref.range.r1, ref.range.c2],
        [left, bottom, ref.range.r2, ref.range.c1],
        [right, bottom, ref.range.r2, ref.range.c2],
      ] as const) {
        if (near(x, cx) && near(y, cy)) return { index, range: ref.range, corner: { row, col } };
      }
      const inX = x >= left - 3 && x <= right + 3;
      const inY = y >= top - 3 && y <= bottom + 3;
      if (((near(x, left) || near(x, right)) && inY) || ((near(y, top) || near(y, bottom)) && inX)) return { index, range: ref.range };
    }
    return undefined;
  }

  function onPointerMoveHover(ev: PointerEvent) {
    if (drag) return;
    const { x, y } = localXY(ev);
    const hit = geo.hit(x, y);
    const hover = hit.kind === 'cell' ? { row: hit.row, col: hit.col } : null;
    const prev = ctl.hoverCell;
    if (hover?.row !== prev?.row || hover?.col !== prev?.col) ctl.hoverCell = hover;
    if (hit.kind === 'colHeader' && geo.colResizeEdge(x) !== undefined) hoverCursor = 'col-resize';
    else if (hit.kind === 'rowHeader' && geo.rowResizeEdge(y) !== undefined) hoverCursor = 'row-resize';
    else if (hit.kind === 'colHeader') hoverCursor = 'col-select';
    else if (hit.kind === 'rowHeader') hoverCursor = 'row-select';
    else if (hit.kind === 'cell' && ctl.edit && rangeFinderHit(x, y)) hoverCursor = rangeFinderHit(x, y)?.corner ? 'nwse-resize' : 'move';
    else if (hit.kind === 'cell' && fillHandleHit(x, y)) hoverCursor = 'crosshair';
    else if (hit.kind === 'cell' && borderHit(x, y)) hoverCursor = 'move';
    else hoverCursor = 'cell';
  }

  function onPointerDown(ev: PointerEvent) {
    if (ev.button === 2) return; // context menu handles right-click
    if (ev.button !== 0) return;
    const { x, y } = localXY(ev);
    // Clicks on the scrollbars belong to the scroller.
    if (x > scroller.clientWidth || y > scroller.clientHeight) return;
    ev.preventDefault();
    focusInput();
    ctl.menu = null;
    const hit = geo.hit(x, y);
    const additive = ev.metaKey || ev.ctrlKey;
    const extend = ev.shiftKey;

    const finder = ctl.edit && hit.kind === 'cell' ? rangeFinderHit(x, y) : undefined;
    if (finder && hit.kind === 'cell') {
      const r = finder.range;
      drag =
        finder.corner
          ? { kind: 'refResize', index: finder.index, fixed: { row: finder.corner.row === r.r1 ? r.r2 : r.r1, col: finder.corner.col === r.c1 ? r.c2 : r.c1 } }
          : { kind: 'refMove', index: finder.index, origin: r, grab: { row: hit.row, col: hit.col } };
      capture(ev);
      return;
    }
    if (ctl.edit && hit.kind === 'cell' && ctl.canPoint(true)) {
      ctl.pointTo({ row: hit.row, col: hit.col }, extend);
      drag = { kind: 'point', anchor: { row: hit.row, col: hit.col } };
      capture(ev);
      return;
    }
    if (ctl.edit && !ctl.commitEdit()) return;

    if (hit.kind === 'outline') {
      const b = outlineView && outlineButtonAt(outlineView, x, y);
      if (!b) return;
      const axisOutline = b.axis === 'row' ? ctl.outline.rows : ctl.outline.cols;
      if (b.action.kind === 'level') showLevel(doc, b.axis, axisOutline, b.action.level);
      else toggleRun(doc, b.axis, axisOutline, b.action.run);
      return;
    }
    if (hit.kind === 'corner') {
      ctl.selectRange({ r1: 1, c1: 1, r2: MAX_ROW, c2: MAX_COL }, doc.selection.active);
      return;
    }
    if (hit.kind === 'colHeader') {
      const edge = geo.colResizeEdge(x);
      if (edge !== undefined) {
        drag = { kind: 'colResize', col: edge, startX: x, startW: geo.colW(edge) };
        capture(ev);
        return;
      }
      const sel = doc.selection;
      if (extend) {
        const a = sel.anchor.col;
        doc.setSelection({ ...sel, ranges: replaceActive(sel.ranges, sel.activeRange, { r1: 1, r2: MAX_ROW, c1: Math.min(a, hit.col), c2: Math.max(a, hit.col) }) });
      } else {
        const range = { r1: 1, r2: MAX_ROW, c1: hit.col, c2: hit.col };
        const top = { row: geo.frozenRows + 1 > 1 ? 1 : geo.rowAt(geo.headerH + geo.frozenH + 1), col: hit.col };
        if (additive) doc.setSelection({ ranges: [...sel.ranges, range], activeRange: sel.ranges.length, active: top, anchor: top });
        else doc.setSelection({ ranges: [range], activeRange: 0, active: top, anchor: top });
      }
      drag = { kind: 'colHeader', anchor: extend ? doc.selection.anchor.col : hit.col };
      capture(ev);
      return;
    }
    if (hit.kind === 'rowHeader') {
      const edge = geo.rowResizeEdge(y);
      if (edge !== undefined) {
        drag = { kind: 'rowResize', row: edge, startY: y, startH: geo.rowH(edge) };
        capture(ev);
        return;
      }
      const sel = doc.selection;
      if (extend) {
        const a = sel.anchor.row;
        doc.setSelection({ ...sel, ranges: replaceActive(sel.ranges, sel.activeRange, { c1: 1, c2: MAX_COL, r1: Math.min(a, hit.row), r2: Math.max(a, hit.row) }) });
      } else {
        const range = { c1: 1, c2: MAX_COL, r1: hit.row, r2: hit.row };
        const left = { row: hit.row, col: geo.colAt(geo.headerW + geo.frozenW + 1) };
        const pos = geo.frozenCols > 0 ? { row: hit.row, col: 1 } : left;
        if (additive) doc.setSelection({ ranges: [...sel.ranges, range], activeRange: sel.ranges.length, active: pos, anchor: pos });
        else doc.setSelection({ ranges: [range], activeRange: 0, active: pos, anchor: pos });
      }
      drag = { kind: 'rowHeader', anchor: extend ? doc.selection.anchor.row : hit.row };
      capture(ev);
      return;
    }
    const range = doc.selection.ranges[0];
    if (range && !fillHandleHit(x, y) && borderHit(x, y)) {
      drag = {
        kind: 'move',
        grab: { dr: Math.min(Math.max(hit.row, range.r1), range.r2) - range.r1, dc: Math.min(Math.max(hit.col, range.c1), range.c2) - range.c1 },
        size: { rows: range.r2 - range.r1 + 1, cols: range.c2 - range.c1 + 1 },
        to: null,
      };
      capture(ev);
      return;
    }
    if (range && fillHandleHit(x, y)) {
      drag = { kind: 'fill', source: range };
      startFillDrag(ctl);
      capture(ev);
      return;
    }
    ctl.selectCell({ row: hit.row, col: hit.col }, { extend, add: additive && !extend });
    drag = { kind: 'select', extend, from: { row: hit.row, col: hit.col }, moved: false };
    capture(ev);
  }

  function replaceActive(ranges: readonly Range[], i: number, r: Range): Range[] {
    const out = ranges.slice();
    out[i] = r;
    return out;
  }

  function capture(ev: PointerEvent) {
    (ev.currentTarget as HTMLElement).setPointerCapture(ev.pointerId);
  }

  function onPointerMove(ev: PointerEvent) {
    if (!drag) {
      onPointerMoveHover(ev);
      return;
    }
    const { x, y } = localXY(ev);
    lastPointer = { x, y };
    dragTo(x, y);
    maybeAutoScroll(x, y);
  }

  function dragTo(x: number, y: number) {
    if (!drag) return;
    const row = geo.rowAt(Math.max(geo.headerH + 1, Math.min(y, geo.height - 1)));
    const col = geo.colAt(Math.max(geo.headerW + 1, Math.min(x, geo.width - 1)));
    switch (drag.kind) {
      case 'select':
        if (row !== drag.from.row || col !== drag.from.col) drag.moved = true;
        ctl.selectCell({ row, col }, { extend: true });
        break;
      case 'point':
        if (ctl.edit?.point) ctl.edit.point.anchor = drag.anchor;
        ctl.pointTo({ row, col }, true);
        break;
      case 'refMove': {
        const { origin, grab } = drag;
        const dr = Math.max(1 - origin.r1, Math.min(MAX_ROW - origin.r2, row - grab.row));
        const dc = Math.max(1 - origin.c1, Math.min(MAX_COL - origin.c2, col - grab.col));
        ctl.retargetEditRef(drag.index, { r1: origin.r1 + dr, c1: origin.c1 + dc, r2: origin.r2 + dr, c2: origin.c2 + dc });
        break;
      }
      case 'refResize':
        ctl.retargetEditRef(drag.index, rangeOf(drag.fixed, { row, col }));
        break;
      case 'colHeader': {
        const sel = doc.selection;
        doc.setSelection({ ...sel, ranges: replaceActive(sel.ranges, sel.activeRange, { r1: 1, r2: MAX_ROW, c1: Math.min(drag.anchor, col), c2: Math.max(drag.anchor, col) }) });
        break;
      }
      case 'rowHeader': {
        const sel = doc.selection;
        doc.setSelection({ ...sel, ranges: replaceActive(sel.ranges, sel.activeRange, { c1: 1, c2: MAX_COL, r1: Math.min(drag.anchor, row), r2: Math.max(drag.anchor, row) }) });
        break;
      }
      case 'colResize': {
        const w = Math.max(0, drag.startW + (x - drag.startX));
        resizeGuide = { x: geo.colX(drag.col) + w, label: `${(w / doc.zoom).toFixed(0)} px` };
        break;
      }
      case 'rowResize': {
        const h = Math.max(0, drag.startH + (y - drag.startY));
        resizeGuide = { y: geo.rowY(drag.row) + h, label: `${(h / doc.zoom).toFixed(0)} px` };
        break;
      }
      case 'fill':
        previewFill(ctl, drag.source, { row, col });
        break;
      case 'move': {
        const r1 = Math.max(1, Math.min(MAX_ROW - drag.size.rows + 1, row - drag.grab.dr));
        const c1 = Math.max(1, Math.min(MAX_COL - drag.size.cols + 1, col - drag.grab.dc));
        drag.to = { row: r1, col: c1 };
        ctl.dragPreview = { r1, c1, r2: r1 + drag.size.rows - 1, c2: c1 + drag.size.cols - 1 };
        break;
      }
    }
  }

  function maybeAutoScroll(x: number, y: number) {
    if (!drag || drag.kind === 'colResize' || drag.kind === 'rowResize') return;
    const margin = 16;
    const dx = x > geo.width - margin ? 20 : x < geo.headerW + geo.frozenW + 2 && doc.scrollX > 0 ? -20 : 0;
    const dy = y > geo.height - margin ? 20 : y < geo.headerH + geo.frozenH + 2 && doc.scrollY > 0 ? -20 : 0;
    if (!dx && !dy) {
      clearInterval(autoScrollTimer);
      autoScrollTimer = 0;
      return;
    }
    if (autoScrollTimer) return;
    autoScrollTimer = window.setInterval(() => {
      doc.setScroll(doc.scrollX + dx, doc.scrollY + dy);
      dragTo(lastPointer.x, lastPointer.y);
    }, 40);
  }

  function onPointerUp(ev: PointerEvent) {
    clearInterval(autoScrollTimer);
    autoScrollTimer = 0;
    const d = drag;
    drag = null;
    if (!d) return;
    const { x, y } = localXY(ev);
    if (d.kind === 'colResize') {
      resizeGuide = null;
      const px = Math.max(0, d.startW + (x - d.startX)) / doc.zoom;
      const cols = selectedHeaderIndices('col', d.col);
      ctl.resizeColumns(cols, px);
    } else if (d.kind === 'rowResize') {
      resizeGuide = null;
      const px = Math.max(0, d.startH + (y - d.startY)) / doc.zoom;
      const rows = selectedHeaderIndices('row', d.row);
      ctl.resizeRows(rows, px);
    } else if (d.kind === 'move') {
      ctl.dragPreview = null;
      // Option (Mac) / Ctrl (Windows) held at release copies instead of moving.
      if (d.to) dragSelection(ctl, d.to, ev.altKey || ev.ctrlKey);
    } else if (d.kind === 'fill') {
      finishFillDrag(ctl, d.source, ev.altKey || ev.ctrlKey);
    } else if (d.kind === 'select' && formatPainter.active) {
      formatPainter.apply(ctl, currentRange(doc.selection));
    } else if (d.kind === 'select' && !d.moved && !d.extend && !ev.metaKey && !ev.ctrlKey && !ctl.edit) {
      // A plain click on a linked cell follows the link, as in Excel.
      const link = hyperlinkAt(doc.ws, d.from.row, d.from.col);
      if (link) followHyperlink(ctl, link);
    }
  }

  /** Resizing a header inside a whole-column selection resizes every selected column. */
  function selectedHeaderIndices(axis: 'row' | 'col', index: number): number[] {
    const out: number[] = [];
    for (const r of doc.selection.ranges) {
      const whole = axis === 'col' ? r.r1 === 1 && r.r2 === MAX_ROW : r.c1 === 1 && r.c2 === MAX_COL;
      const lo = axis === 'col' ? r.c1 : r.r1;
      const hi = axis === 'col' ? r.c2 : r.r2;
      if (whole && index >= lo && index <= hi) for (let i = lo; i <= hi; i++) out.push(i);
    }
    return out.length > 0 ? out : [index];
  }

  function onDblClick(ev: MouseEvent) {
    const { x, y } = localXY(ev);
    const hit = geo.hit(x, y);
    if (hit.kind === 'colHeader') {
      const edge = geo.colResizeEdge(x);
      if (edge !== undefined) autofitColumns(ctl, selectedHeaderIndices('col', edge));
      return;
    }
    if (hit.kind === 'rowHeader') {
      const edge = geo.rowResizeEdge(y);
      if (edge !== undefined) autofitRows(ctl, selectedHeaderIndices('row', edge));
      return;
    }
    if (hit.kind === 'cell') {
      if (fillHandleHit(x, y)) {
        ctl.fillHandleDoubleClick();
        return;
      }
      ctl.selectCell({ row: hit.row, col: hit.col });
      ctl.startEdit();
    }
  }

  function onContextMenu(ev: MouseEvent) {
    ev.preventDefault();
    const { x, y } = localXY(ev);
    const hit = geo.hit(x, y);
    if (ctl.edit) ctl.commitEdit();
    if (hit.kind === 'cell' && !selectionContains(doc.selection, hit.row, hit.col)) ctl.selectCell({ row: hit.row, col: hit.col });
    if (hit.kind === 'colHeader' && !doc.selection.ranges.some((r) => r.r1 === 1 && r.r2 === MAX_ROW && hit.col >= r.c1 && hit.col <= r.c2)) {
      ctl.selectRange({ r1: 1, r2: MAX_ROW, c1: hit.col, c2: hit.col });
    }
    if (hit.kind === 'rowHeader' && !doc.selection.ranges.some((r) => r.c1 === 1 && r.c2 === MAX_COL && hit.row >= r.r1 && hit.row <= r.r2)) {
      ctl.selectRange({ r1: hit.row, r2: hit.row, c1: 1, c2: MAX_COL });
    }
    ctl.menu = { x: ev.clientX, y: ev.clientY, kind: hit.kind === 'colHeader' ? 'colHeader' : hit.kind === 'rowHeader' ? 'rowHeader' : 'cell' };
  }

  function onWheel(ev: WheelEvent) {
    if (ev.ctrlKey || ev.metaKey) {
      // Pinch / Ctrl+wheel zooms, in Excel's 10% steps.
      ev.preventDefault();
      const next = Math.round((doc.zoom - Math.sign(ev.deltaY) * 0.1) * 10) / 10;
      doc.setZoom(next);
    }
  }

  // Clipboard goes through the native events so pastes from Excel / Sheets
  // and copies into them carry HTML and TSV.
  function onCopy(ev: ClipboardEvent) {
    if (ctl.edit) return;
    ev.preventDefault();
    writeCopyEvent(ctl, ev, false);
  }
  function onCut(ev: ClipboardEvent) {
    if (ctl.edit) return;
    ev.preventDefault();
    writeCopyEvent(ctl, ev, true);
  }
  function onPaste(ev: ClipboardEvent) {
    if (ctl.edit) return;
    ev.preventDefault();
    pasteFromEvent(ctl, ev);
  }

  const editorFont = $derived(editStyle ? canvasFont(editStyle, doc.zoom) : '14px sans-serif');
</script>

<div
  class="grid-host"
  bind:this={host}
  style:cursor={hoverCursor === 'col-select' ? 's-resize' : hoverCursor === 'row-select' ? 'e-resize' : hoverCursor === 'cell' ? (formatPainter.active ? 'copy' : 'cell') : hoverCursor}
  role="grid"
  aria-label="Worksheet"
  tabindex="-1"
  onpointerdown={onPointerDown}
  onpointermove={onPointerMove}
  onpointerup={onPointerUp}
  onpointerleave={() => (ctl.hoverCell = null)}
  ondblclick={onDblClick}
  oncontextmenu={onContextMenu}
  onwheel={onWheel}
>
  <canvas bind:this={canvas} style:width="{geo.width}px" style:height="{geo.height}px"></canvas>
  <div class="scroller" bind:this={scroller} onscroll={onScroll}>
    <div class="spacer" style:width="{extent.w}px" style:height="{extent.h}px"></div>
  </div>

  <DrawingLayer />
  <ShapeDrawOverlay />
  <GridOverlays />
  <TotalRowButton />
  <FormulaAssist />

  {#if resizeGuide}
    {#if resizeGuide.x !== undefined}
      <div class="guide v" style:left="{resizeGuide.x}px"><span>{resizeGuide.label}</span></div>
    {:else}
      <div class="guide h" style:top="{resizeGuide.y}px"><span>{resizeGuide.label}</span></div>
    {/if}
  {/if}

  <div
    class="cell-editor"
    class:visible={editorVisible}
    style:left="{activeRect.x - 1}px"
    style:top="{activeRect.y - 1}px"
    style:min-width="{activeRect.w + 1}px"
    style:min-height="{activeRect.h + 1}px"
    style:max-width="{Math.max(activeRect.w, geo.width - activeRect.x - 4)}px"
    style:font={editorFont}
    style:color={editStyle?.color ?? '#000'}
    style:background={editStyle?.fill ?? '#fff'}
    style:text-align={editStyle?.hAlign === 'right' ? 'right' : editStyle?.hAlign === 'center' ? 'center' : 'left'}
  >
    {#if editorVisible && ctl.edit}
      <CellEditorText text={composingText ?? ctl.edit.text} refs={composingText === null ? ctl.editRefs : []} />
    {/if}
    <textarea
      bind:this={input}
      class="cell-input"
      class:editing={editorVisible}
      spellcheck="false"
      autocomplete="off"
      aria-label="Cell editor"
      oninput={onInput}
      onkeydown={onKeyDown}
      oncompositionstart={onCompositionStart}
      oncompositionend={onCompositionEnd}
      onselect={onSelect}
      onkeyup={onSelect}
      onclick={onSelect}
      oncopy={onCopy}
      oncut={onCut}
      onpaste={onPaste}
    ></textarea>
  </div>
</div>

<style>
  .grid-host {
    position: relative;
    flex: 1;
    min-height: 0;
    overflow: hidden;
    background: #fff;
    user-select: none;
    touch-action: none;
  }
  canvas {
    position: absolute;
    inset: 0;
    pointer-events: none;
  }
  .scroller {
    position: absolute;
    inset: 0;
    overflow: scroll;
    /* The canvas draws everything; the scroller only owns the scrollbars and wheel. */
    pointer-events: auto;
    opacity: 1;
  }
  .scroller .spacer {
    pointer-events: none;
  }
  .guide {
    position: absolute;
    pointer-events: none;
    z-index: 3;
  }
  .guide.v {
    top: 0;
    bottom: 0;
    border-left: 1px dashed #1e7145;
  }
  .guide.h {
    left: 0;
    right: 0;
    border-top: 1px dashed #1e7145;
  }
  .guide span {
    position: absolute;
    top: 2px;
    left: 4px;
    background: #fffde7;
    border: 1px solid #999;
    font: 11px -apple-system, sans-serif;
    padding: 1px 4px;
    white-space: nowrap;
  }
  .cell-editor {
    position: absolute;
    z-index: 4;
    box-sizing: border-box;
    opacity: 0;
    pointer-events: none;
    overflow: hidden;
    width: max-content;
  }
  .cell-editor.visible {
    opacity: 1;
    pointer-events: auto;
    border: 2px solid #1e7145;
    box-shadow: 0 1px 4px rgba(0, 0, 0, 0.18);
    padding: 0 2px;
  }
  .cell-input {
    position: absolute;
    inset: 0;
    width: 100%;
    height: 100%;
    margin: 0;
    padding: inherit;
    border: 0;
    outline: none;
    resize: none;
    overflow: hidden;
    background: transparent;
    font: inherit;
    text-align: inherit;
    color: transparent;
    caret-color: #000;
    white-space: pre-wrap;
    word-break: break-all;
    line-height: 1.25;
  }
</style>
