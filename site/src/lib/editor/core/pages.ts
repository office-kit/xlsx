// Pagination: where printed pages break, from the sheet's page setup. Used by
// Page Break Preview and the page count; printing itself goes through the
// browser, which paginates on its own.

import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { parseRangeAddress, type Range } from './address.ts';
import type { AxisIndex } from './axis.ts';
import { usedRange } from './cells.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';

/** ECMA-376 §18.3.1.63 paper size codes → inches (portrait). */
const PAPER_INCHES: Readonly<Record<number, readonly [number, number]>> = {
  1: [8.5, 11],
  5: [8.5, 14],
  8: [11.69, 16.54],
  9: [8.27, 11.69],
  11: [5.83, 8.27],
  13: [7.17, 10.12],
};
const DEFAULT_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
const PX_PER_INCH = 96;

export interface Pagination {
  /** The printed area (print area or used range). */
  readonly area: Range;
  /** First row / column of each page, in order; the first is the area start. */
  readonly rowStarts: readonly number[];
  readonly colStarts: readonly number[];
  /** Breaks the user inserted (rows/cols after which a page starts). */
  readonly manualRows: ReadonlySet<number>;
  readonly manualCols: ReadonlySet<number>;
}

export function printArea(doc: SpreadsheetEditor): Range | undefined {
  const dn = doc.wb.definedNames.find((d) => d.name === '_xlnm.Print_Area' && d.scope === doc.activeSheetIndex);
  const first = dn?.value.split(',')[0];
  const parsed = first ? parseRangeAddress(first.trim()) : undefined;
  return parsed?.range ?? usedRange(doc.ws);
}

/** Paper and margins in sheet pixels at 100% zoom (divided by the print scale). */
export interface PaperMetrics {
  readonly paperW: number;
  readonly paperH: number;
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
  readonly header: number;
  readonly footer: number;
}

export function paperMetrics(ws: Worksheet, scale: number): PaperMetrics {
  const ps = ws.pageSetup ?? {};
  const [pw, ph] = PAPER_INCHES[ps.paperSize ?? 1] ?? PAPER_INCHES[1] ?? [8.5, 11];
  const landscape = ps.orientation === 'landscape';
  const m = ws.pageMargins ?? DEFAULT_MARGINS;
  const px = (inches: number) => (inches * PX_PER_INCH) / scale;
  return {
    paperW: px(landscape ? ph : pw),
    paperH: px(landscape ? pw : ph),
    left: px(m.left),
    right: px(m.right),
    top: px(m.top),
    bottom: px(m.bottom),
    header: px(m.header),
    footer: px(m.footer),
  };
}

/** Printable page size in sheet pixels at 100% zoom, after margins and scaling. */
export function pageBox(ws: Worksheet, scale: number): { w: number; h: number } {
  const m = paperMetrics(ws, scale);
  return { w: m.paperW - m.left - m.right, h: m.paperH - m.top - m.bottom };
}

function breaksAlong(axis: AxisIndex, from: number, to: number, size: number, manual: ReadonlySet<number>): number[] {
  const starts = [from];
  let used = 0;
  for (let i = from; i <= to; i++) {
    const s = axis.sizeOf(i);
    if (i > from && (manual.has(i - 1) || (used + s > size && used > 0))) {
      starts.push(i);
      used = 0;
    }
    used += s;
  }
  return starts;
}

/** Manual page breaks: the row / column after which a new page starts. */
export function manualBreaks(ws: Worksheet): { rows: Set<number>; cols: Set<number> } {
  return {
    rows: new Set(ws.rowBreaks.flatMap((b) => (b.id !== undefined ? [b.id] : []))),
    cols: new Set(ws.colBreaks.flatMap((b) => (b.id !== undefined ? [b.id] : []))),
  };
}

/** The print scale: the page setup's percentage, or the shrink that fits the area on W×H pages. */
export function printScale(doc: SpreadsheetEditor, area: Range): number {
  const ws = doc.ws;
  const ps = ws.pageSetup ?? {};
  const fit = ws.sheetProperties?.pageSetUpPr?.fitToPage === true;
  if (!fit) return (ps.scale ?? 100) / 100;
  const box = pageBox(ws, 1);
  const totalW = doc.cols.offsetOf(area.c2 + 1) - doc.cols.offsetOf(area.c1);
  const totalH = doc.rows.offsetOf(area.r2 + 1) - doc.rows.offsetOf(area.r1);
  const fw = ps.fitToWidth ?? 1;
  const fh = ps.fitToHeight ?? 1;
  // 0 pages means unconstrained along that axis.
  const sw = fw > 0 ? Math.min(1, (box.w * fw) / totalW) : 1;
  const sh = fh > 0 ? Math.min(1, (box.h * fh) / totalH) : 1;
  return Math.max(0.1, Math.min(sw, sh));
}

export function paginate(doc: SpreadsheetEditor): Pagination | undefined {
  const ws = doc.ws;
  const area = printArea(doc);
  if (!area) return undefined;
  const { rows: manualRows, cols: manualCols } = manualBreaks(ws);
  const scale = printScale(doc, area);
  const box = pageBox(ws, scale);
  return {
    area,
    rowStarts: breaksAlong(doc.rows, area.r1, area.r2, box.h, manualRows),
    colStarts: breaksAlong(doc.cols, area.c1, area.c2, box.w, manualCols),
    manualRows,
    manualCols,
  };
}

export function pageCount(p: Pagination | undefined): number {
  return p ? p.rowStarts.length * p.colStarts.length : 0;
}
