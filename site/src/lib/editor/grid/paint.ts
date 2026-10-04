// Canvas painter for the visible part of a worksheet.
//
// Cost is proportional to the cells on screen, never to the sheet: every pass
// walks the visible row/column window (plus a few columns either side so text
// overflowing from off-screen cells still shows) and looks cells up in the
// sparse row maps. Nothing is cached across frames except resolved styles,
// which StyleResolver keeps per style id.

import type { Cell } from '@office-kit/xlsx/cell';
import type { Workbook } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { fromStorageFormula } from '../calc/index.ts';
import type { CellPos, Range } from '../core/address.ts';
import type { OutlineLayout } from './outline-layout.ts';
import { colLetter, inRange, MAX_COL, MAX_ROW } from '../core/address.ts';
import type { MergeIndex } from '../core/merges.ts';
import type { RenderStyle, StrokeStyle, StyleResolver } from '../core/render-style.ts';
import { canvasFont } from '../core/render-style.ts';
import type { HeaderParts } from '../core/header-footer.ts';
import type { PagedAxis } from '../core/page-layout.ts';
import type { Pagination, PaperMetrics } from '../core/pages.ts';
import type { Selection } from '../core/selection.ts';
import type { TableLook } from '../core/table-style.ts';
import { currentRange } from '../core/selection.ts';
import { displayCell, fitGeneralNumber, hashes, type CellDisplay } from './display.ts';
import type { GridGeometry } from './geometry.ts';
import { paintSparkline, type SparklinePaintCell } from './sparkline-paint.ts';

export interface CellOverlay {
  /** Conditional-format fill/font overrides. */
  readonly fill?: string;
  readonly color?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly strike?: boolean;
  readonly underline?: boolean;
  /** Data bar: fraction of the cell width and colour. */
  readonly bar?: { readonly start: number; readonly end: number; readonly color: string };
  /** Icon-set glyph drawn at the left of the cell. */
  readonly icon?: { readonly glyph: string; readonly color: string };
  readonly hideValue?: boolean;
}

export interface Theme {
  readonly background: string;
  readonly gridline: string;
  readonly headerBg: string;
  readonly headerText: string;
  readonly headerLine: string;
  readonly headerSelBg: string;
  readonly headerSelText: string;
  readonly headerFullBg: string;
  readonly headerFullText: string;
  readonly accent: string;
  readonly selectionFill: string;
  readonly frozenLine: string;
}

export const LIGHT_THEME: Theme = {
  background: '#FFFFFF',
  gridline: '#E1E1E1',
  headerBg: '#F5F5F5',
  headerText: '#444444',
  headerLine: '#D5D5D5',
  headerSelBg: '#DCEFE3',
  headerSelText: '#0E5C2F',
  headerFullBg: '#1E7145',
  headerFullText: '#FFFFFF',
  accent: '#1E7145',
  selectionFill: 'rgba(30, 113, 69, 0.12)',
  frozenLine: '#9E9E9E',
};

export interface PaintInput {
  readonly ctx: CanvasRenderingContext2D;
  readonly geo: GridGeometry;
  readonly wb: Workbook;
  readonly ws: Worksheet;
  readonly styles: StyleResolver;
  readonly merges: MergeIndex;
  readonly selection: Selection;
  /** Page Layout view: sheets of paper with their header and footer, or undefined in the other views. */
  readonly paper: PaperView | undefined;
  /** View ▸ Focus Cell: tint of the active cell's row and column, or undefined when off. */
  readonly focusCell: string | undefined;
  readonly showGridlines: boolean;
  readonly showHeaders: boolean;
  readonly theme: Theme;
  /** Default style id for empty cells, from row/column formats. */
  readonly defaultStyleAt: (row: number, col: number) => number;
  readonly overlayAt: ((row: number, col: number, cell: Cell | undefined) => CellOverlay | undefined) | undefined;
  /** Table-style formatting under each cell, beneath the cell's own format. */
  readonly tableLookAt?: ((row: number, col: number) => TableLook | undefined) | undefined;
  /** Ranges outlined in colour while a formula is edited. */
  readonly refHighlights?: ReadonlyArray<{ readonly range: Range; readonly color: string }>;
  /** Marching-ants range after Copy/Cut; `antsPhase` animates it. */
  readonly copyRange?: Range | null;
  readonly antsPhase?: number;
  /** Fill-handle drag preview (dashed grey outline). */
  readonly dragPreview?: Range | null;
  /** The cell under the in-place editor; its text is not painted. */
  readonly editing?: CellPos | null;
  /** Cells carrying a note (red triangle) or a threaded comment (purple flag, grey once resolved). */
  readonly commentMarks?: ReadonlyArray<CommentMark>;
  readonly showFormulas?: boolean;
  readonly showZeros?: boolean;
  /** Trace Precedents/Dependents: an arrow from each range to the traced cell. */
  readonly traces?: ReadonlyArray<{ readonly from: CellPos; readonly range: Range; readonly kind: 'precedents' | 'dependents' }>;
  /** Circle Invalid Data marks. */
  readonly invalidCircles?: readonly Range[];
  readonly outline: OutlineLayout | undefined;
  /** Body rows of AutoFilters with active conditions; their row numbers turn blue. */
  readonly filteredRows?: ReadonlyArray<readonly [number, number]>;
  /** Page Break Preview: page grid over the printed area. */
  readonly pages: Pagination | undefined;
  readonly pageLabel?: (page: number) => string;
  /** Sparklines on the sheet; each is drawn only when its cell is in the painted pane. */
  readonly sparklines?: ReadonlyArray<SparklinePaintCell>;
}

const PAGE_LINE = '#2F5FD0';

const FILTERED_ROW_TEXT = '#0563C1';

interface Pane {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly r1: number;
  readonly r2: number;
  readonly c1: number;
  readonly c2: number;
}

function panes(geo: GridGeometry): Pane[] {
  const out: Pane[] = [];
  const [mr1, mr2] = geo.mainRows();
  const [mc1, mc2] = geo.mainCols();
  const bodyX = geo.headerW;
  const bodyY = geo.headerH;
  const fw = Math.min(geo.frozenW, geo.width - bodyX);
  const fh = Math.min(geo.frozenH, geo.height - bodyY);
  const mainX = bodyX + fw;
  const mainY = bodyY + fh;
  out.push({ x: mainX, y: mainY, w: geo.width - mainX, h: geo.height - mainY, r1: mr1, r2: mr2, c1: mc1, c2: mc2 });
  if (geo.frozenRows > 0) out.push({ x: mainX, y: bodyY, w: geo.width - mainX, h: fh, r1: 1, r2: geo.frozenRows, c1: mc1, c2: mc2 });
  if (geo.frozenCols > 0) out.push({ x: bodyX, y: mainY, w: fw, h: geo.height - mainY, r1: mr1, r2: mr2, c1: 1, c2: geo.frozenCols });
  if (geo.frozenRows > 0 && geo.frozenCols > 0) {
    out.push({ x: bodyX, y: bodyY, w: fw, h: fh, r1: 1, r2: geo.frozenRows, c1: 1, c2: geo.frozenCols });
  }
  return out;
}

function setStroke(ctx: CanvasRenderingContext2D, s: StrokeStyle, zoom: number): void {
  ctx.strokeStyle = s.color;
  ctx.lineWidth = s.double ? 1 : Math.max(1, Math.round(s.width * Math.min(zoom, 2)));
  ctx.setLineDash(s.dash.map((d) => d * Math.max(1, zoom)));
}

function strokeLine(ctx: CanvasRenderingContext2D, s: StrokeStyle, zoom: number, x1: number, y1: number, x2: number, y2: number): void {
  setStroke(ctx, s, zoom);
  const half = (ctx.lineWidth % 2) / 2;
  if (s.double) {
    const vertical = x1 === x2;
    ctx.beginPath();
    for (const off of [-1, 1]) {
      const dx = vertical ? off : 0;
      const dy = vertical ? 0 : off;
      ctx.moveTo(Math.round(x1 + dx) + 0.5, Math.round(y1 + dy) + 0.5);
      ctx.lineTo(Math.round(x2 + dx) + 0.5, Math.round(y2 + dy) + 0.5);
    }
    ctx.stroke();
    return;
  }
  ctx.beginPath();
  ctx.moveTo(Math.round(x1) + half, Math.round(y1) + half);
  ctx.lineTo(Math.round(x2) + half, Math.round(y2) + half);
  ctx.stroke();
}

const PAD = 3;

interface TextJob {
  readonly cell: Cell;
  readonly style: RenderStyle;
  readonly display: CellDisplay;
  /** Cell box (merged area for an anchor). */
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly overlay: CellOverlay | undefined;
  readonly row: number;
  readonly col: number;
}

export function paintGrid(input: PaintInput): void {
  const { ctx, geo, theme } = input;
  ctx.save();
  ctx.fillStyle = theme.background;
  ctx.fillRect(0, 0, geo.width, geo.height);
  for (const pane of panes(geo)) {
    if (pane.w <= 0 || pane.h <= 0) continue;
    ctx.save();
    ctx.beginPath();
    ctx.rect(pane.x, pane.y, pane.w, pane.h);
    ctx.clip();
    paintPane(input, pane);
    if (input.paper) paintPaper(input, input.paper);
    if (input.pages) paintPages(input, input.pages);
    paintSelection(input, pane);
    ctx.restore();
  }
  paintFrozenLines(input);
  if (input.showHeaders) paintHeaders(input);
  if (input.outline) paintOutline(input, input.outline);
  ctx.restore();
}

function styleIdFor(input: PaintInput, row: number, col: number, cell: Cell | undefined): number {
  return cell ? cell.styleId : input.defaultStyleAt(row, col);
}

// A table has a few dozen distinct looks and a sheet a few hundred styles, so
// the combined styles are kept rather than rebuilt for every painted cell.
const lookStyles = new WeakMap<TableLook, WeakMap<RenderStyle, RenderStyle>>();

/** The cell's style with the table style filling in what the cell leaves unset. */
function underTable(style: RenderStyle, look: TableLook | undefined): RenderStyle {
  if (!look) return style;
  let byStyle = lookStyles.get(look);
  if (!byStyle) {
    byStyle = new WeakMap();
    lookStyles.set(look, byStyle);
  }
  let merged = byStyle.get(style);
  if (!merged) {
    merged = {
      ...style,
      fill: style.fill ?? look.fill ?? null,
      color: style.ownColor || !look.color ? style.color : look.color,
      bold: style.bold || look.bold === true,
      top: style.top ?? look.top ?? null,
      bottom: style.bottom ?? look.bottom ?? null,
      left: style.left ?? look.left ?? null,
      right: style.right ?? look.right ?? null,
    };
    byStyle.set(style, merged);
  }
  return merged;
}

function paintPane(input: PaintInput, pane: Pane): void {
  const { ctx, geo, ws, styles, merges, theme } = input;
  const window: Range = { r1: pane.r1, c1: pane.c1, r2: pane.r2, c2: pane.c2 };
  const paneMerges = merges.intersecting(window);
  const coveredByMerge = (row: number, col: number): Range | undefined => {
    for (const m of paneMerges) if (inRange(m, row, col)) return m;
    return undefined;
  };

  // 1. Gridlines.
  if (input.showGridlines) {
    ctx.strokeStyle = theme.gridline;
    ctx.lineWidth = 1;
    ctx.setLineDash([]);
    ctx.beginPath();
    for (let c = pane.c1; c <= pane.c2 + 1; c++) {
      if (c <= MAX_COL + 1 && (c > MAX_COL || geo.colW(c) > 0 || c === pane.c2 + 1)) {
        const x = Math.round(geo.colX(c)) - 0.5;
        ctx.moveTo(x, pane.y);
        ctx.lineTo(x, pane.y + pane.h);
      }
    }
    for (let r = pane.r1; r <= pane.r2 + 1; r++) {
      if (r <= MAX_ROW + 1) {
        const y = Math.round(geo.rowY(r)) - 0.5;
        ctx.moveTo(pane.x, y);
        ctx.lineTo(pane.x + pane.w, y);
      }
    }
    ctx.stroke();
  }

  // 2. Fills (and merged areas, which hide the gridlines inside them).
  const jobs: TextJob[] = [];
  const borderCells: Array<{ style: RenderStyle; x: number; y: number; w: number; h: number }> = [];
  const doneMerges = new Set<Range>();
  const textC1 = Math.max(1, pane.c1 - 12);
  const textC2 = Math.min(MAX_COL, pane.c2 + 12);
  for (let r = pane.r1; r <= pane.r2; r++) {
    const rh = geo.rowH(r);
    if (rh <= 0) continue;
    const rowMap = ws.rows.get(r);
    const y = geo.rowY(r);
    for (let c = textC1; c <= textC2; c++) {
      const inPane = c >= pane.c1 && c <= pane.c2;
      const cell = rowMap?.get(c);
      if (!inPane && (cell === undefined || cell.value === null)) continue;
      const cw = geo.colW(c);
      if (cw <= 0) continue;
      const merge = paneMerges.length > 0 ? coveredByMerge(r, c) : undefined;
      if (merge) {
        if (doneMerges.has(merge) || !inPane) continue;
        doneMerges.add(merge);
        const anchor = ws.rows.get(merge.r1)?.get(merge.c1);
        const style = underTable(styles.get(styleIdFor(input, merge.r1, merge.c1, anchor)), input.tableLookAt?.(merge.r1, merge.c1));
        const rect = geo.rectOf(merge);
        const overlay = input.overlayAt?.(merge.r1, merge.c1, anchor);
        ctx.fillStyle = overlay?.fill ?? style.fill ?? theme.background;
        ctx.fillRect(rect.x, rect.y, rect.w - 1, rect.h - 1);
        borderCells.push({ style, ...rect });
        if (anchor && anchor.value !== null) {
          jobs.push({ cell: anchor, style, display: displayCell(input.wb, anchor, style.numFmt), ...rect, overlay, row: merge.r1, col: merge.c1 });
        }
        continue;
      }
      const styleId = styleIdFor(input, r, c, cell);
      const overlay = input.overlayAt?.(r, c, cell);
      const look = input.tableLookAt?.(r, c);
      if (styleId === 0 && !overlay && !look && (cell === undefined || cell.value === null)) continue;
      const style = underTable(styles.get(styleId), look);
      const x = geo.colX(c);
      if (inPane) {
        const fill = overlay?.fill ?? style.fill;
        if (fill) {
          ctx.fillStyle = fill;
          ctx.fillRect(x - 1, y - 1, cw + 1, rh + 1);
          if (style.pattern) paintPattern(ctx, style.pattern, x, y, cw, rh);
        }
        if (overlay?.bar) {
          const bx = x + 2 + (cw - 4) * overlay.bar.start;
          const bw = Math.max(0, (cw - 4) * (overlay.bar.end - overlay.bar.start));
          // Excel draws a classic data bar as a gradient fading to white.
          const grad = ctx.createLinearGradient(bx, 0, bx + bw, 0);
          grad.addColorStop(0, overlay.bar.color);
          grad.addColorStop(1, '#ffffff');
          ctx.fillStyle = grad;
          ctx.fillRect(bx, y + 2, bw, rh - 5);
        }
        if (style.top || style.right || style.bottom || style.left || style.diagonalDown || style.diagonalUp) {
          borderCells.push({ style, x, y, w: cw, h: rh });
        }
      }
      if (cell && cell.value !== null) {
        jobs.push({ cell, style, display: displayCell(input.wb, cell, style.numFmt), x, y, w: cw, h: rh, overlay, row: r, col: c });
      }
    }
  }

  // 3. Text, with Excel's overflow into empty neighbours.
  for (const job of jobs) {
    if (input.editing && job.row === input.editing.row && job.col === input.editing.col) continue;
    paintText(input, job, coveredByMerge);
  }

  // 4. Borders on top of fills and text.
  for (const b of borderCells) {
    const s = b.style;
    const right = b.x + b.w - 1;
    const bottom = b.y + b.h - 1;
    if (s.top) strokeLine(ctx, s.top, geo.zoom, b.x - 1, b.y - 1, right, b.y - 1);
    if (s.bottom) strokeLine(ctx, s.bottom, geo.zoom, b.x - 1, bottom, right, bottom);
    if (s.left) strokeLine(ctx, s.left, geo.zoom, b.x - 1, b.y - 1, b.x - 1, bottom);
    if (s.right) strokeLine(ctx, s.right, geo.zoom, right, b.y - 1, right, bottom);
    if (s.diagonalDown) strokeLine(ctx, s.diagonalDown, geo.zoom, b.x, b.y, right, bottom);
    if (s.diagonalUp) strokeLine(ctx, s.diagonalUp, geo.zoom, b.x, bottom, right, b.y);
  }
  ctx.setLineDash([]);

  // 4b. Sparklines, inside their cells.
  for (const sp of input.sparklines ?? []) {
    if (sp.row >= pane.r1 && sp.row <= pane.r2 && sp.col >= pane.c1 && sp.col <= pane.c2) paintSparkline(ctx, geo, sp, styles.palette);
  }

  // 5. Comment indicators. Walked from the marks, not the cells: a note or
  // comment may sit on a cell that holds no value.
  for (const m of input.commentMarks ?? []) {
    if (m.row < pane.r1 || m.row > pane.r2 || m.col < pane.c1 || m.col > pane.c2) continue;
    const x = geo.colX(m.col) + geo.colW(m.col) - 1;
    const y = geo.rowY(m.row);
    ctx.beginPath();
    if (m.kind === 'note') {
      ctx.fillStyle = NOTE_MARK_COLOR;
      ctx.moveTo(x - 6, y);
      ctx.lineTo(x, y);
      ctx.lineTo(x, y + 6);
    } else {
      // Excel draws a small speech-bubble flag in the corner.
      ctx.fillStyle = m.kind === 'thread' ? THREAD_MARK_COLOR : RESOLVED_MARK_COLOR;
      ctx.rect(x - 8, y + 1, 8, 5);
      ctx.moveTo(x - 6, y + 6);
      ctx.lineTo(x - 3, y + 6);
      ctx.lineTo(x - 6, y + 8);
    }
    ctx.fill();
  }
}

export interface CommentMark {
  readonly row: number;
  readonly col: number;
  readonly kind: 'note' | 'thread' | 'resolved';
}

const NOTE_MARK_COLOR = '#D9302C';
const THREAD_MARK_COLOR = '#7B61C4';
const RESOLVED_MARK_COLOR = '#8A8886';

function paintPattern(ctx: CanvasRenderingContext2D, pattern: NonNullable<RenderStyle['pattern']>, x: number, y: number, w: number, h: number): void {
  // Patterns are drawn as a sparse dot/line texture; exact Excel bitmaps are
  // not worth the cost for a rarely used feature.
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w - 1, h - 1);
  ctx.clip();
  ctx.fillStyle = pattern.fg;
  const step = pattern.type.startsWith('dark') ? 2 : pattern.type.startsWith('light') ? 3 : 4;
  for (let yy = y; yy < y + h; yy += step) {
    for (let xx = x + ((yy - y) % (step * 2) === 0 ? 0 : 1); xx < x + w; xx += step) ctx.fillRect(xx, yy, 1, 1);
  }
  ctx.restore();
}

function paintText(input: PaintInput, job: TextJob, mergeAt: (r: number, c: number) => Range | undefined): void {
  const { ctx, geo, ws } = input;
  const { style, display, overlay } = job;
  if (overlay?.hideValue) return;
  const zoom = geo.zoom;
  const font = canvasFont(
    { ...style, bold: overlay?.bold ?? style.bold, italic: overlay?.italic ?? style.italic },
    zoom,
  );
  ctx.font = font;
  const measure = (s: string): number => ctx.measureText(s).width;
  const pad = PAD * zoom;
  const indent = style.indent * 9 * zoom;
  const innerW = job.w - pad * 2 - indent - 1;

  let text = input.showFormulas && job.cell.value !== null && typeof job.cell.value === 'object' && 'kind' in job.cell.value && job.cell.value.kind === 'formula' ? `=${fromStorageFormula(job.cell.value.formula)}` : display.text;
  if (!input.showFormulas && input.showZeros === false && display.number === 0) return;

  // Alignment: General puts numbers right, text left, booleans/errors centred.
  let h = style.hAlign;
  if (h === 'general') h = display.kind === 'number' ? 'right' : display.kind === 'bool' || display.kind === 'error' ? 'center' : 'left';
  if (input.showFormulas) h = 'left';
  if (h === 'fill') {
    const unit = measure(text) || 1;
    text = text.repeat(Math.max(1, Math.floor(innerW / unit)));
    h = 'left';
  }

  const isNumber = display.kind === 'number' && !input.showFormulas;
  if (isNumber && measure(text) > innerW && !style.wrap && !style.shrink) {
    text = style.numFmt === 'General' && display.number !== undefined ? fitGeneralNumber(display.number, innerW, measure) : hashes(innerW, measure);
  }

  let fontPx = style.fontPx * zoom;
  if (style.shrink && !style.wrap) {
    const tw = measure(text);
    if (tw > innerW && tw > 0) {
      fontPx = Math.max(1, fontPx * (innerW / tw));
      ctx.font = font.replace(/[\d.]+px/, `${fontPx.toFixed(2)}px`);
    }
  }

  // Clip region: the cell, widened over empty neighbours for overflowing text.
  let clipL = job.x;
  let clipR = job.x + job.w;
  const canOverflow = !style.wrap && !style.shrink && display.kind === 'text' && job.w === geo.colW(job.col) && style.rotation === 0;
  const textW = measure(text);
  if (canOverflow && textW > innerW) {
    const rowMap = ws.rows.get(job.row);
    const blank = (c: number): boolean => {
      const n = rowMap?.get(c);
      return (n === undefined || n.value === null || n.value === '') && !mergeAt(job.row, c);
    };
    const need = textW - innerW;
    if (h === 'left' || h === 'center' || h === 'justify' || h === 'distributed' || h === 'centerContinuous') {
      let extra = h === 'center' ? need / 2 : need;
      for (let c = job.col + 1; c <= MAX_COL && extra > 0 && blank(c); c++) {
        clipR += geo.colW(c);
        extra -= geo.colW(c);
      }
    }
    if (h === 'right' || h === 'center') {
      let extra = h === 'center' ? need / 2 : need;
      for (let c = job.col - 1; c >= 1 && extra > 0 && blank(c); c--) {
        clipL -= geo.colW(c);
        extra -= geo.colW(c);
      }
    }
  }

  ctx.save();
  ctx.beginPath();
  ctx.rect(clipL, job.y, clipR - clipL - 1, job.h - 1);
  ctx.clip();
  ctx.fillStyle = overlay?.color ?? display.color ?? (job.cell.hyperlinkId !== undefined && style.color === '#000000' ? '#0563C1' : style.color);
  ctx.textBaseline = 'alphabetic';

  if (overlay?.icon) {
    ctx.save();
    ctx.fillStyle = overlay.icon.color;
    ctx.font = `${(12 * zoom).toFixed(1)}px sans-serif`;
    ctx.fillText(overlay.icon.glyph, job.x + pad, job.y + job.h - 4 * zoom);
    ctx.restore();
  }

  const lineH = fontPx * 1.25;
  const lines = style.wrap ? wrapLines(text, Math.max(1, innerW), measure) : text.split('\n').slice(0, 1);
  const blockH = lines.length * lineH;
  const ascent = fontPx * 0.8;
  let top: number;
  switch (style.vAlign) {
    case 'top':
      top = job.y + 2 * zoom;
      break;
    case 'center':
    case 'distributed':
    case 'justify':
      top = job.y + (job.h - blockH) / 2;
      break;
    default:
      top = job.y + job.h - blockH - 2 * zoom;
  }

  if (style.rotation !== 0) {
    paintRotated(ctx, text, style.rotation, job, ascent);
    ctx.restore();
    return;
  }

  lines.forEach((line, i) => {
    const w = measure(line);
    let x: number;
    if (h === 'right') x = job.x + job.w - pad - indent - w - 1;
    else if (h === 'center' || h === 'centerContinuous' || h === 'distributed') x = job.x + (job.w - w) / 2;
    else x = job.x + pad + indent;
    const baseline = top + i * lineH + ascent;
    ctx.fillText(line, x, baseline);
    if (overlay?.underline ?? style.underline !== 'none') {
      ctx.fillRect(x, baseline + Math.max(1, fontPx * 0.08), w, Math.max(1, zoom));
      if (style.underline === 'double') ctx.fillRect(x, baseline + Math.max(3, fontPx * 0.18), w, Math.max(1, zoom));
    }
    if (overlay?.strike ?? style.strike) ctx.fillRect(x, baseline - fontPx * 0.3, w, Math.max(1, zoom));
  });
  ctx.restore();
}

function paintRotated(ctx: CanvasRenderingContext2D, text: string, rotation: number, job: TextJob, ascent: number): void {
  const cx = job.x + job.w / 2;
  const cy = job.y + job.h / 2;
  ctx.translate(cx, cy);
  if (rotation === 255) {
    // Stacked vertical text: one character per line.
    const chars = [...text];
    const lineH = ascent * 1.3;
    const start = -((chars.length - 1) * lineH) / 2;
    chars.forEach((ch, i) => {
      const w = ctx.measureText(ch).width;
      ctx.fillText(ch, -w / 2, start + i * lineH + ascent / 2);
    });
    return;
  }
  // OOXML: 1–90 rotate counter-clockwise, 91–180 clockwise by (angle - 90).
  const deg = rotation <= 90 ? -rotation : rotation - 90;
  ctx.rotate((deg * Math.PI) / 180);
  const w = ctx.measureText(text).width;
  ctx.fillText(text, -w / 2, ascent / 2);
}

export function wrapLines(text: string, maxW: number, measure: (s: string) => number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    if (para === '') {
      out.push('');
      continue;
    }
    // Break at spaces when possible; CJK text has none, so fall back to characters.
    const tokens = para.match(/\S+\s*|\s+/g) ?? [para];
    let line = '';
    for (const token of tokens) {
      if (measure(line + token) <= maxW || line === '') {
        if (measure(token) > maxW && line === '') {
          for (const ch of token) {
            if (measure(line + ch) > maxW && line !== '') {
              out.push(line);
              line = '';
            }
            line += ch;
          }
          continue;
        }
        line += token;
      } else {
        out.push(line.trimEnd());
        line = token.trimStart();
      }
    }
    out.push(line.trimEnd());
  }
  return out;
}

function paintSelection(input: PaintInput, pane: Pane): void {
  const { ctx, geo, theme, selection } = input;
  const zoom = geo.zoom;
  const active = selection.active;
  const multi = selection.ranges.length > 1;
  // The active cell stays unshaded: cut it out of its range with even-odd fill.
  const activeMerge = input.merges.at(active.row, active.col) ?? { r1: active.row, c1: active.col, r2: active.row, c2: active.col };
  const ar = geo.rectOf(activeMerge);
  if (input.focusCell) {
    ctx.fillStyle = input.focusCell;
    ctx.globalAlpha = 0.22;
    const bands: Range[] = [
      { r1: activeMerge.r1, r2: activeMerge.r2, c1: 1, c2: MAX_COL },
      { r1: 1, r2: MAX_ROW, c1: activeMerge.c1, c2: activeMerge.c2 },
    ];
    ctx.beginPath();
    for (const band of bands) {
      const clipped = clipRange(band, pane);
      if (clipped.r1 > clipped.r2 || clipped.c1 > clipped.c2) continue;
      const rect = geo.rectOf(clipped);
      ctx.rect(rect.x, rect.y, rect.w - 1, rect.h - 1);
    }
    // Nonzero fill keeps the crossing cell from being tinted twice.
    ctx.fill();
    ctx.globalAlpha = 1;
  }
  ctx.fillStyle = theme.selectionFill;
  for (const range of selection.ranges) {
    const clipped = clipRange(range, pane);
    if (clipped.r1 > clipped.r2 || clipped.c1 > clipped.c2) continue;
    if (range.r1 === range.r2 && range.c1 === range.c2 && !multi) continue;
    const rect = geo.rectOf(clipped);
    ctx.beginPath();
    ctx.rect(rect.x, rect.y, rect.w - 1, rect.h - 1);
    if (inRange(range, active.row, active.col)) ctx.rect(ar.x, ar.y, ar.w - 1, ar.h - 1);
    ctx.fill('evenodd');
  }

  const range = currentRange(selection);
  const rr = geo.rectOf(range);
  ctx.strokeStyle = theme.accent;
  ctx.setLineDash([]);
  ctx.lineWidth = Math.max(2, Math.round(2 * Math.min(zoom, 1.5)));
  if (!multi) {
    ctx.strokeRect(rr.x - 1, rr.y - 1, rr.w, rr.h);
    // Fill handle.
    const hs = Math.max(5, Math.round(6 * Math.min(zoom, 1.5)));
    ctx.fillStyle = theme.background;
    ctx.fillRect(rr.x + rr.w - 1 - hs / 2 - 1, rr.y + rr.h - 1 - hs / 2 - 1, hs + 2, hs + 2);
    ctx.fillStyle = theme.accent;
    ctx.fillRect(rr.x + rr.w - 1 - hs / 2, rr.y + rr.h - 1 - hs / 2, hs, hs);
  } else {
    ctx.lineWidth = 1;
    ctx.strokeRect(ar.x - 0.5, ar.y - 0.5, ar.w, ar.h);
  }
  if (input.dragPreview) {
    const d = geo.rectOf(input.dragPreview);
    ctx.strokeStyle = '#7F7F7F';
    ctx.lineWidth = 1;
    ctx.setLineDash([3, 2]);
    ctx.strokeRect(d.x - 0.5, d.y - 0.5, d.w, d.h);
  }
  for (const hl of input.refHighlights ?? []) {
    const d = geo.rectOf(hl.range);
    ctx.strokeStyle = hl.color;
    ctx.lineWidth = 2;
    ctx.setLineDash([]);
    ctx.strokeRect(d.x - 1, d.y - 1, d.w, d.h);
    ctx.fillStyle = hl.color;
    for (const [px, py] of [
      [d.x - 1, d.y - 1],
      [d.x + d.w - 1, d.y - 1],
      [d.x - 1, d.y + d.h - 1],
      [d.x + d.w - 1, d.y + d.h - 1],
    ] as const) {
      ctx.fillRect(px - 2, py - 2, 4, 4);
    }
  }
  paintAudit(input);
  if (input.copyRange) {
    const d = geo.rectOf(input.copyRange);
    ctx.lineWidth = 2;
    ctx.setLineDash([4, 3]);
    ctx.lineDashOffset = -(input.antsPhase ?? 0);
    ctx.strokeStyle = theme.accent;
    ctx.strokeRect(d.x - 1, d.y - 1, d.w, d.h);
    ctx.lineDashOffset = 0;
    ctx.setLineDash([]);
  }
}

const TRACE_COLOR = '#2F5597';
const INVALID_COLOR = '#E00000';

function paintAudit(input: PaintInput): void {
  const { ctx, geo } = input;
  ctx.setLineDash([]);
  for (const t of input.traces ?? []) {
    const box = geo.rectOf(t.range);
    const cell = geo.rectOf({ r1: t.from.row, c1: t.from.col, r2: t.from.row, c2: t.from.col });
    const multi = t.range.r1 !== t.range.r2 || t.range.c1 !== t.range.c2;
    if (multi) {
      ctx.strokeStyle = TRACE_COLOR;
      ctx.lineWidth = 1.5;
      ctx.strokeRect(box.x + 0.5, box.y + 0.5, box.w - 2, box.h - 2);
    }
    // Arrows run from the precedent to the dependent, as in Excel.
    const a = t.kind === 'precedents' ? box : cell;
    const b = t.kind === 'precedents' ? cell : box;
    const x1 = a.x + Math.min(a.w, 40) / 2;
    const y1 = a.y + Math.min(a.h, 20) / 2;
    const x2 = b.x + Math.min(b.w, 40) / 2;
    const y2 = b.y + Math.min(b.h, 20) / 2;
    ctx.strokeStyle = TRACE_COLOR;
    ctx.fillStyle = TRACE_COLOR;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.arc(x1, y1, 2.5, 0, Math.PI * 2);
    ctx.fill();
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    const ang = Math.atan2(y2 - y1, x2 - x1);
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(x2 - 9 * Math.cos(ang - 0.4), y2 - 9 * Math.sin(ang - 0.4));
    ctx.lineTo(x2 - 9 * Math.cos(ang + 0.4), y2 - 9 * Math.sin(ang + 0.4));
    ctx.closePath();
    ctx.fill();
  }
  ctx.strokeStyle = INVALID_COLOR;
  ctx.lineWidth = 1.5;
  for (const r of input.invalidCircles ?? []) {
    const d = geo.rectOf(r);
    ctx.beginPath();
    ctx.ellipse(d.x + d.w / 2, d.y + d.h / 2, d.w / 2 + 4, d.h / 2 + 3, 0, 0, Math.PI * 2);
    ctx.stroke();
  }
}

function clipRange(r: Range, pane: Pane): Range {
  return {
    r1: Math.max(r.r1, pane.r1),
    c1: Math.max(r.c1, pane.c1),
    r2: Math.min(r.r2, pane.r2),
    c2: Math.min(r.c2, pane.c2),
  };
}

function paintFrozenLines(input: PaintInput): void {
  const { ctx, geo, theme } = input;
  ctx.strokeStyle = theme.frozenLine;
  ctx.lineWidth = 1;
  ctx.setLineDash([]);
  ctx.beginPath();
  if (geo.frozenRows > 0) {
    const y = Math.round(geo.headerH + geo.frozenH) - 0.5;
    ctx.moveTo(geo.headerW, y);
    ctx.lineTo(geo.width, y);
  }
  if (geo.frozenCols > 0) {
    const x = Math.round(geo.headerW + geo.frozenW) - 0.5;
    ctx.moveTo(x, geo.headerH);
    ctx.lineTo(x, geo.height);
  }
  ctx.stroke();
}

/** Page Break Preview: grey outside the printed area, blue page borders, faint page numbers. */
export interface PaperView {
  readonly cols: PagedAxis;
  readonly rows: PagedAxis;
  readonly metrics: PaperMetrics;
  /** Expanded header / footer of the page at (row page, column page), or undefined for a page outside the printed area. */
  readonly texts: (rowPage: number, colPage: number) => { header: HeaderParts; footer: HeaderParts } | undefined;
}

const PAPER_GUTTER_FILL = '#E6E6E6';

function paintPaper(input: PaintInput, paper: PaperView): void {
  const { ctx, geo } = input;
  const z = geo.zoom;
  const sx = (px: number) => geo.headerW + px * z - geo.scrollX;
  const sy = (px: number) => geo.headerH + px * z - geo.scrollY;
  const colPages = [paper.cols.pageOf(geo.colAt(geo.headerW + 1)), paper.cols.pageOf(geo.colAt(geo.width - 1))] as const;
  const rowPages = [paper.rows.pageOf(geo.rowAt(geo.headerH + 1)), paper.rows.pageOf(geo.rowAt(geo.height - 1))] as const;
  const sheets: Array<{ x: number; y: number; w: number; h: number; cp: number; rp: number }> = [];
  for (let cp = colPages[0]; cp <= colPages[1] + 1 && cp < paper.cols.starts.length; cp++) {
    for (let rp = rowPages[0]; rp <= rowPages[1] + 1 && rp < paper.rows.starts.length; rp++) {
      sheets.push({ x: sx(paper.cols.paperOffset(cp)), y: sy(paper.rows.paperOffset(rp)), w: paper.cols.paper * z, h: paper.rows.paper * z, cp, rp });
    }
  }
  ctx.save();
  // Grey gutter around the paper, then white margins over the gridlines inside it.
  ctx.fillStyle = PAPER_GUTTER_FILL;
  ctx.beginPath();
  ctx.rect(0, 0, geo.width, geo.height);
  for (const s of sheets) ctx.rect(s.x, s.y, s.w, s.h);
  ctx.fill('evenodd');
  ctx.fillStyle = input.theme.background;
  ctx.beginPath();
  for (const s of sheets) {
    ctx.rect(s.x, s.y, s.w, s.h);
    const cs = paper.cols.pageSpan(s.cp);
    const rs = paper.rows.pageSpan(s.rp);
    const x1 = sx(paper.cols.paperOffset(s.cp) + paper.cols.lead);
    const y1 = sy(paper.rows.paperOffset(s.rp) + paper.rows.lead);
    const x2 = cs.end > cs.start ? geo.colX(cs.end - 1) + geo.colW(cs.end - 1) : x1;
    const y2 = rs.end > rs.start ? geo.rowY(rs.end - 1) + geo.rowH(rs.end - 1) : y1;
    ctx.rect(x1, y1, x2 - x1, y2 - y1);
  }
  ctx.fill('evenodd');
  ctx.strokeStyle = '#C8C8C8';
  ctx.lineWidth = 1;
  for (const s of sheets) ctx.strokeRect(Math.round(s.x) - 0.5, Math.round(s.y) - 0.5, Math.round(s.w) + 1, Math.round(s.h) + 1);

  ctx.font = `${Math.max(8, Math.round(11 * z))}px Calibri, Carlito, 'Segoe UI', Arial, sans-serif`;
  ctx.fillStyle = '#000';
  const m = paper.metrics;
  for (const s of sheets) {
    const texts = paper.texts(s.rp, s.cp);
    if (!texts) continue;
    const left = s.x + m.left * z;
    const right = s.x + s.w - m.right * z;
    const center = s.x + s.w / 2;
    const headerY = s.y + m.header * z;
    const footerY = s.y + s.h - m.footer * z;
    for (const [sections, y, baseline] of [
      [texts.header, headerY, 'top'],
      [texts.footer, footerY, 'bottom'],
    ] as const) {
      ctx.textBaseline = baseline;
      ctx.textAlign = 'left';
      if (sections.left) ctx.fillText(sections.left, left, y);
      ctx.textAlign = 'center';
      if (sections.center) ctx.fillText(sections.center, center, y);
      ctx.textAlign = 'right';
      if (sections.right) ctx.fillText(sections.right, right, y);
    }
  }
  ctx.restore();
}

function paintPages(input: PaintInput, p: Pagination): void {
  const { ctx, geo } = input;
  const area = geo.rectOf(p.area);
  ctx.save();
  ctx.fillStyle = 'rgba(128, 128, 128, 0.35)';
  ctx.beginPath();
  ctx.rect(0, 0, geo.width, geo.height);
  ctx.rect(area.x, area.y, area.w, area.h);
  ctx.fill('evenodd');
  const zoom = geo.zoom;
  ctx.font = `bold ${Math.round(48 * zoom)}px -apple-system, 'Segoe UI', Arial, sans-serif`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillStyle = 'rgba(128, 128, 128, 0.25)';
  let page = 0;
  for (let ci = 0; ci < p.colStarts.length; ci++) {
    for (let ri = 0; ri < p.rowStarts.length; ri++) {
      page++;
      const r1 = p.rowStarts[ri] ?? p.area.r1;
      const c1 = p.colStarts[ci] ?? p.area.c1;
      const r2 = (p.rowStarts[ri + 1] ?? p.area.r2 + 1) - 1;
      const c2 = (p.colStarts[ci + 1] ?? p.area.c2 + 1) - 1;
      const r = geo.rectOf({ r1, c1, r2, c2 });
      if (r.x > geo.width || r.y > geo.height || r.x + r.w < 0 || r.y + r.h < 0) continue;
      ctx.fillText(input.pageLabel?.(page) ?? String(page), r.x + r.w / 2, r.y + r.h / 2);
    }
  }
  ctx.strokeStyle = PAGE_LINE;
  ctx.lineWidth = 3;
  ctx.setLineDash([]);
  ctx.strokeRect(area.x, area.y, area.w, area.h);
  ctx.lineWidth = 2;
  for (const row of p.rowStarts.slice(1)) {
    const y = geo.rowY(row);
    ctx.setLineDash(p.manualRows.has(row - 1) ? [] : [6, 4]);
    ctx.beginPath();
    ctx.moveTo(area.x, y);
    ctx.lineTo(area.x + area.w, y);
    ctx.stroke();
  }
  for (const col of p.colStarts.slice(1)) {
    const x = geo.colX(col);
    ctx.setLineDash(p.manualCols.has(col - 1) ? [] : [6, 4]);
    ctx.beginPath();
    ctx.moveTo(x, area.y);
    ctx.lineTo(x, area.y + area.h);
    ctx.stroke();
  }
  ctx.restore();
}

function paintOutline(input: PaintInput, layout: OutlineLayout): void {
  const { ctx, geo, theme } = input;
  ctx.save();
  ctx.fillStyle = theme.background;
  if (geo.outlineW > 0) ctx.fillRect(0, 0, geo.outlineW, geo.height);
  if (geo.outlineH > 0) ctx.fillRect(0, 0, geo.width, geo.outlineH);
  ctx.strokeStyle = '#6E6E6E';
  ctx.fillStyle = '#6E6E6E';
  ctx.lineWidth = 1;
  for (const b of layout.brackets) {
    ctx.beginPath();
    if (b.axis === 'row') {
      const x = Math.round(b.at) + 0.5;
      const from = Math.max(b.from, geo.headerH);
      if (b.to <= from) continue;
      ctx.moveTo(x, from + 2);
      ctx.lineTo(x, b.to - 1);
      ctx.lineTo(x + 4, b.to - 1);
    } else {
      const y = Math.round(b.at) + 0.5;
      const from = Math.max(b.from, geo.headerW);
      if (b.to <= from) continue;
      ctx.moveTo(from + 2, y);
      ctx.lineTo(b.to - 1, y);
      ctx.lineTo(b.to - 1, y + 4);
    }
    ctx.stroke();
  }
  ctx.font = `${Math.round(9 * Math.max(0.6, Math.min(geo.zoom, 2)))}px -apple-system, 'Segoe UI', Arial, sans-serif`;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  for (const b of layout.buttons) {
    const inGutter = b.action.kind === 'level' || (b.axis === 'row' ? b.y + b.size > geo.headerH : b.x + b.size > geo.headerW);
    if (!inGutter) continue;
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(Math.round(b.x), Math.round(b.y), b.size, b.size);
    ctx.strokeStyle = '#8A8A8A';
    ctx.strokeRect(Math.round(b.x) + 0.5, Math.round(b.y) + 0.5, b.size - 1, b.size - 1);
    ctx.fillStyle = '#333333';
    const cx = Math.round(b.x) + b.size / 2;
    const cy = Math.round(b.y) + b.size / 2;
    if (b.action.kind === 'level') ctx.fillText(String(b.action.level), cx, cy + 0.5);
    else {
      ctx.fillRect(Math.round(cx - b.size / 2 + 3), Math.round(cy) - 0.5, b.size - 6, 1);
      if (b.action.collapsed) ctx.fillRect(Math.round(cx) - 0.5, Math.round(cy - b.size / 2 + 3), 1, b.size - 6);
    }
  }
  ctx.textAlign = 'start';
  ctx.restore();
}

function paintHeaders(input: PaintInput): void {
  const { ctx, geo, theme, selection } = input;
  const zoom = Math.max(0.6, Math.min(geo.zoom, 2));
  const selCols = (c: number) => selection.ranges.some((r) => c >= r.c1 && c <= r.c2);
  const selRows = (r: number) => selection.ranges.some((s) => r >= s.r1 && r <= s.r2);
  const fullCol = (c: number) => selection.ranges.some((r) => r.r1 === 1 && r.r2 === MAX_ROW && c >= r.c1 && c <= r.c2);
  const fullRow = (r: number) => selection.ranges.some((s) => s.c1 === 1 && s.c2 === MAX_COL && r >= s.r1 && r <= s.r2);
  ctx.font = `${(11 * zoom).toFixed(1)}px -apple-system, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif`;
  ctx.textBaseline = 'middle';

  // Column header strip (below any column outline strip).
  const top = geo.outlineH;
  const left = geo.outlineW;
  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(left, top, geo.width - left, geo.headerH - top);
  const colSpans: Array<[number, number, number, number]> = [];
  const [mc1, mc2] = geo.mainCols();
  for (let c = 1; c <= geo.frozenCols; c++) colSpans.push([c, geo.colX(c), geo.colW(c), 0]);
  for (let c = mc1; c <= mc2; c++) colSpans.push([c, geo.colX(c), geo.colW(c), 1]);
  ctx.save();
  ctx.beginPath();
  ctx.rect(geo.headerW, top, geo.width - geo.headerW, geo.headerH - top);
  ctx.clip();
  for (const [c, x, w, scrolling] of colSpans) {
    if (w <= 0) continue;
    if (scrolling && x + w <= geo.headerW + geo.frozenW) continue;
    const full = fullCol(c);
    const sel = selCols(c);
    if (full || sel) {
      ctx.fillStyle = full ? theme.headerFullBg : theme.headerSelBg;
      ctx.fillRect(x, top, w, geo.headerH - top);
    }
    ctx.fillStyle = full ? theme.headerFullText : sel ? theme.headerSelText : theme.headerText;
    const label = colLetter(c);
    const tw = ctx.measureText(label).width;
    ctx.fillText(label, x + (w - tw) / 2, (top + geo.headerH) / 2 + 1);
    ctx.fillStyle = theme.headerLine;
    ctx.fillRect(Math.round(x + w) - 1, top, 1, geo.headerH - top);
    if (sel && !full) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(x, geo.headerH - 2, w, 2);
    }
  }
  ctx.restore();

  // Row header gutter (right of any row outline strip).
  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(left, geo.headerH, geo.headerW - left, geo.height - geo.headerH);
  const [mr1, mr2] = geo.mainRows();
  const rowsToPaint: Array<[number, number]> = [];
  for (let r = 1; r <= geo.frozenRows; r++) rowsToPaint.push([r, 0]);
  for (let r = mr1; r <= mr2; r++) rowsToPaint.push([r, 1]);
  ctx.save();
  ctx.beginPath();
  ctx.rect(left, geo.headerH, geo.headerW - left, geo.height - geo.headerH);
  ctx.clip();
  for (const [r, scrolling] of rowsToPaint) {
    const y = geo.rowY(r);
    const h = geo.rowH(r);
    if (h <= 0) continue;
    if (scrolling && y + h <= geo.headerH + geo.frozenH) continue;
    const full = fullRow(r);
    const sel = selRows(r);
    if (full || sel) {
      ctx.fillStyle = full ? theme.headerFullBg : theme.headerSelBg;
      ctx.fillRect(left, y, geo.headerW - left, h);
    }
    const filtered = input.filteredRows?.some(([a, b]) => r >= a && r <= b);
    ctx.fillStyle = full ? theme.headerFullText : sel ? theme.headerSelText : filtered ? FILTERED_ROW_TEXT : theme.headerText;
    const label = String(r);
    const tw = ctx.measureText(label).width;
    if (h > 6) ctx.fillText(label, left + (geo.headerW - left - tw) / 2, y + h / 2 + 1);
    ctx.fillStyle = theme.headerLine;
    ctx.fillRect(left, Math.round(y + h) - 1, geo.headerW - left, 1);
    if (sel && !full) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(geo.headerW - 2, y, 2, h);
    }
  }
  ctx.restore();

  // Header borders and the select-all corner.
  ctx.fillStyle = theme.headerLine;
  ctx.fillRect(left, geo.headerH - 1, geo.width - left, 1);
  ctx.fillRect(geo.headerW - 1, top, 1, geo.height - top);
  ctx.fillStyle = theme.headerBg;
  ctx.fillRect(left, top, geo.headerW - 1 - left, geo.headerH - 1 - top);
  ctx.fillStyle = '#B7B7B7';
  ctx.beginPath();
  ctx.moveTo(geo.headerW - 4, geo.headerH - 4);
  ctx.lineTo(geo.headerW - 4, geo.headerH - 4 - 10 * zoom);
  ctx.lineTo(geo.headerW - 4 - 10 * zoom, geo.headerH - 4);
  ctx.fill();
}
