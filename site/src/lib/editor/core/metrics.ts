// Unit conversions between the OOXML model (column widths in "characters",
// row heights in points) and screen pixels at 100% zoom, plus the per-sheet
// axis layouts built from them.

import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { MAX_COL, MAX_ROW } from './address.ts';
import { AxisIndex } from './axis.ts';

/**
 * Maximum digit width of the default font in pixels. ECMA-376 §18.3.1.13
 * defines column widths in multiples of this; 7 px is Calibri/Aptos 11 at
 * 96 dpi, the figure Excel uses for the default Normal style.
 */
export const MAX_DIGIT_WIDTH = 7;
export const PX_PER_PT = 96 / 72;

// Stored widths (`<col width>`, `defaultColWidth`) include 5 px of cell
// padding; the "8.43" Excel shows in the Column Width dialog does not.
export const DEFAULT_COL_PX = 64;
export const DEFAULT_ROW_HEIGHT_PT = 15;

/** Stored column width → pixels (ECMA-376 §18.3.1.13). */
export function colWidthToPx(width: number): number {
  if (width <= 0) return 0;
  return Math.trunc(((256 * width + Math.trunc(128 / MAX_DIGIT_WIDTH)) / 256) * MAX_DIGIT_WIDTH);
}

/** Pixels → stored column width, in the 1/256 character steps Excel writes. */
export function pxToColWidth(px: number): number {
  return Math.trunc((px / MAX_DIGIT_WIDTH) * 256) / 256;
}

/** The width Excel's Column Width dialog shows for a stored width (e.g. 9.140625 → 8.43). */
export function displayColWidth(width: number): number {
  return Math.max(0, Math.trunc(((colWidthToPx(width) - 5) / MAX_DIGIT_WIDTH) * 100 + 0.5) / 100);
}

/** Stored width for a width typed into the Column Width dialog. */
export function storedColWidth(display: number): number {
  if (display <= 0) return 0;
  return Math.trunc(((display * MAX_DIGIT_WIDTH + 5) / MAX_DIGIT_WIDTH) * 256) / 256;
}

export function ptToPx(pt: number): number {
  return Math.round(pt * PX_PER_PT);
}

export function pxToPt(px: number): number {
  return Math.round((px / PX_PER_PT) * 4) / 4;
}

export function defaultColPx(ws: Worksheet): number {
  if (ws.defaultColumnWidth !== undefined) return colWidthToPx(ws.defaultColumnWidth);
  // Excel rounds the base-width default up to a multiple of 8 px.
  if (ws.baseColWidth !== undefined) return Math.ceil((ws.baseColWidth * MAX_DIGIT_WIDTH + 5) / 8) * 8;
  return DEFAULT_COL_PX;
}

export function defaultRowPx(ws: Worksheet): number {
  return ptToPx(ws.defaultRowHeight ?? DEFAULT_ROW_HEIGHT_PT);
}

export function buildColumnAxis(ws: Worksheet): AxisIndex {
  const defaultPx = defaultColPx(ws);
  const overrides = new Map<number, number>();
  for (const dim of ws.columnDimensions.values()) {
    const px = dim.hidden ? 0 : dim.width !== undefined ? colWidthToPx(dim.width) : defaultPx;
    if (px === defaultPx) continue;
    const last = Math.min(dim.max, MAX_COL);
    for (let c = dim.min; c <= last; c++) if (!overrides.has(c)) overrides.set(c, px);
  }
  return new AxisIndex(MAX_COL, defaultPx, overrides);
}

export function buildRowAxis(ws: Worksheet): AxisIndex {
  const defaultPx = defaultRowPx(ws);
  const overrides = new Map<number, number>();
  for (const [row, dim] of ws.rowDimensions) {
    const px = dim.hidden ? 0 : dim.height !== undefined ? ptToPx(dim.height) : defaultPx;
    if (px !== defaultPx) overrides.set(row, px);
  }
  return new AxisIndex(MAX_ROW, defaultPx, overrides);
}
