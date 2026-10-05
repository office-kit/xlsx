// Sparklines drawn into their cells on the grid canvas. Only the sparklines
// whose cell is inside the painted pane are drawn, and their values are read
// once per document version.

import type { SparklineGroup } from '@office-kit/xlsx/worksheet';
import { resolveColor, type ThemePalette } from '../core/theme.ts';
import type { GridGeometry } from './geometry.ts';

export interface SparklinePaintCell {
  readonly row: number;
  readonly col: number;
  readonly group: SparklineGroup;
  /** Plotted values; null for blank / text cells. */
  readonly values: () => ReadonlyArray<number | null>;
}

const PAD = 3;

/** Shared vertical scale of a sparkline, honouring custom axis bounds. */
function bounds(group: SparklineGroup, values: readonly number[]): { min: number; max: number } {
  let min = Math.min(...values);
  let max = Math.max(...values);
  if (group.minAxisType === 'custom' && group.manualMin !== undefined) min = group.manualMin;
  if (group.maxAxisType === 'custom' && group.manualMax !== undefined) max = group.manualMax;
  // Columns grow from zero, as Excel draws them.
  if (group.type === 'column') {
    min = Math.min(0, min);
    max = Math.max(0, max);
  }
  return { min, max: max === min ? min + 1 : max };
}

/** Indices of the first / last / high / low points, for the highlight colours. */
function marks(values: ReadonlyArray<number | null>): { first: number; last: number; high: number; low: number } {
  let first = -1;
  let last = -1;
  let high = -1;
  let low = -1;
  values.forEach((v, i) => {
    if (v === null) return;
    if (first < 0) first = i;
    last = i;
    if (high < 0 || v > (values[high] ?? -Infinity)) high = i;
    if (low < 0 || v < (values[low] ?? Infinity)) low = i;
  });
  return { first, last, high, low };
}

function pointColor(g: SparklineGroup, i: number, v: number, m: ReturnType<typeof marks>, palette: ThemePalette, base: string): string | undefined {
  if (g.high && i === m.high) return resolveColor(g.colorHigh, palette, base);
  if (g.low && i === m.low) return resolveColor(g.colorLow, palette, base);
  if (g.first && i === m.first) return resolveColor(g.colorFirst, palette, base);
  if (g.last && i === m.last) return resolveColor(g.colorLast, palette, base);
  if (g.negative && v < 0) return resolveColor(g.colorNegative, palette, base);
  return undefined;
}

export function paintSparkline(ctx: CanvasRenderingContext2D, geo: GridGeometry, cell: SparklinePaintCell, palette: ThemePalette): void {
  const w = geo.colW(cell.col);
  const h = geo.rowH(cell.row);
  if (w <= PAD * 2 || h <= PAD * 2) return;
  const g = cell.group;
  const raw = cell.values();
  const values = g.displayEmptyCellsAs === 'zero' ? raw.map((v) => v ?? 0) : raw;
  const nums = values.filter((v): v is number => v !== null);
  if (nums.length === 0) return;
  const x0 = geo.colX(cell.col) + PAD;
  const y0 = geo.rowY(cell.row) + PAD;
  const iw = w - PAD * 2;
  const ih = h - PAD * 2;
  const series = resolveColor(g.colorSeries, palette, '#376092');
  const m = marks(values);
  const n = values.length;
  ctx.save();
  ctx.beginPath();
  ctx.rect(geo.colX(cell.col), geo.rowY(cell.row), w, h);
  ctx.clip();

  if (g.type === 'stacked') {
    // Win/Loss: equal-height bars above or below the middle by sign.
    const bw = iw / n;
    values.forEach((v, i) => {
      if (v === null || v === 0) return;
      ctx.fillStyle = pointColor(g, i, v, m, palette, series) ?? series;
      const top = v > 0 ? y0 : y0 + ih / 2;
      ctx.fillRect(x0 + i * bw + bw * 0.1, top, Math.max(1, bw * 0.8), ih / 2);
    });
  } else {
    const { min, max } = bounds(g, nums);
    const yOf = (v: number) => y0 + ih - ((v - min) / (max - min)) * ih;
    if (g.displayXAxis && min < 0 && max > 0) {
      ctx.strokeStyle = resolveColor(g.colorAxis, palette, '#000000');
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(x0, Math.round(yOf(0)) + 0.5);
      ctx.lineTo(x0 + iw, Math.round(yOf(0)) + 0.5);
      ctx.stroke();
    }
    if (g.type === 'column') {
      const bw = iw / n;
      const zero = yOf(Math.max(min, Math.min(max, 0)));
      values.forEach((v, i) => {
        if (v === null) return;
        ctx.fillStyle = pointColor(g, i, v, m, palette, series) ?? series;
        const y = yOf(v);
        ctx.fillRect(x0 + i * bw + bw * 0.1, Math.min(y, zero), Math.max(1, bw * 0.8), Math.max(1, Math.abs(zero - y)));
      });
    } else {
      const xOf = (i: number) => (n === 1 ? x0 + iw / 2 : x0 + (i / (n - 1)) * iw);
      ctx.strokeStyle = series;
      ctx.lineWidth = Math.max(1, ((g.lineWeight ?? 0.75) * 96) / 72) * geo.zoom;
      ctx.lineJoin = 'round';
      ctx.beginPath();
      let pen = false;
      values.forEach((v, i) => {
        if (v === null) {
          if (g.displayEmptyCellsAs !== 'span') pen = false;
          return;
        }
        if (pen) ctx.lineTo(xOf(i), yOf(v));
        else ctx.moveTo(xOf(i), yOf(v));
        pen = true;
      });
      ctx.stroke();
      const r = Math.max(1.5, 2 * geo.zoom);
      values.forEach((v, i) => {
        if (v === null) return;
        const color = pointColor(g, i, v, m, palette, series) ?? (g.markers ? resolveColor(g.colorMarkers, palette, series) : undefined);
        if (!color) return;
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.arc(xOf(i), yOf(v), r, 0, Math.PI * 2);
        ctx.fill();
      });
    }
  }
  ctx.restore();
}
