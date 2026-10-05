// SVG rendering of a resolved chart (core/chart-data.ts `ChartData`), styled
// after Excel's default chart style: grey axis text, light gridlines, legend
// at the bottom. Output is a markup string so the drawing layer can drop it in
// with {@html}; every piece of user text goes through `esc`.

import type { ChartData, LabelSpec, RenderSeries } from '../core/chart-data.ts';
import { seriesColor } from '../core/chart-data.ts';
import type { ThemePalette } from '../core/theme.ts';

const FONT = "-apple-system, 'Segoe UI', 'Hiragino Sans', Calibri, Arial, sans-serif";

function esc(s: string): string {
  return s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;');
}

const f = (n: number): string => (Math.round(n * 10) / 10).toString();

/** Round axis bounds and step, the way chart axes pick 1/2/5 × 10ⁿ intervals. */
export function niceScale(lo: number, hi: number, maxTicks: number): { min: number; max: number; step: number } {
  if (lo === hi) {
    if (lo === 0) return { min: 0, max: 1, step: 0.2 };
    lo = Math.min(0, lo);
    hi = Math.max(0, hi);
  }
  const raw = (hi - lo) / Math.max(1, maxTicks);
  const mag = 10 ** Math.floor(Math.log10(raw));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * mag;
  return { min: Math.floor(lo / step) * step, max: Math.ceil(hi / step) * step, step };
}

function tickLabel(v: number, step: number): string {
  const decimals = Math.max(0, -Math.floor(Math.log10(step) + 1e-9));
  const fixed = v.toFixed(Math.min(10, decimals));
  const [int = '', frac] = fixed.split('.');
  const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return frac ? `${grouped}.${frac}` : grouped;
}

/** A data value as a label: up to 4 decimals, trailing zeros dropped, grouped thousands. */
function valueLabel(v: number): string {
  const [int = '', frac = ''] = (Math.round(v * 10_000) / 10_000).toFixed(4).split('.');
  const trimmed = frac.replace(/0+$/, '');
  const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return trimmed ? `${grouped}.${trimmed}` : grouped;
}

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Colours that depend on the chart background, so a dark chart style keeps readable text. */
interface Ink {
  readonly text: string;
  readonly grid: string;
  readonly minor: string;
  readonly fs: number;
}

function luminance(hex: string): number {
  const h = hex.replace('#', '');
  const [r, g, b] = [0, 2, 4].map((i) => Number.parseInt(h.slice(i, i + 2), 16) / 255);
  return 0.2126 * (r ?? 1) + 0.7152 * (g ?? 1) + 0.0722 * (b ?? 1);
}

function text(ink: Ink, x: number, y: number, s: string, size: number, anchor: 'start' | 'middle' | 'end' = 'middle', extra = ''): string {
  return `<text x="${f(x)}" y="${f(y)}" font-size="${f(size)}" fill="${ink.text}" text-anchor="${anchor}" ${extra}>${esc(s)}</text>`;
}

function legendEntries(data: ChartData, palette: ThemePalette): Array<{ name: string; color: string; line: boolean }> {
  const byPoint = data.kind === 'pie' || data.kind === 'doughnut' || data.kind === 'treemap' || data.kind === 'sunburst' || (data.varyColors && data.series.length === 1);
  if (byPoint) {
    const s = data.series[0];
    return data.categories.map((name, i) => ({ name, color: s?.pointColors?.[i] ?? seriesColor(i, palette), line: false }));
  }
  if (data.kind === 'waterfall' || data.kind === 'funnel' || data.kind === 'histogram' || data.kind === 'pareto' || data.kind === 'boxWhisker') return [];
  return data.series.map((s) => ({ name: s.name, color: s.color, line: (data.kind === 'line' || data.kind === 'scatter' || data.kind === 'radar') && !data.radarFilled && s.line }));
}

/** Draw the legend into the free side of `box` and return the remaining plot box. */
function drawLegend(out: string[], data: ChartData, palette: ThemePalette, box: Box, ink: Ink): Box {
  const pos = data.legend;
  if (!pos) return box;
  const entries = legendEntries(data, palette);
  if (entries.length === 0) return box;
  const fs = ink.fs;
  const swatch = fs * 0.75;
  const itemW = (e: { name: string }) => swatch + 6 + e.name.length * fs * 0.55 + 12;
  const key = (x: number, y: number, e: { name: string; color: string; line: boolean }) =>
    (e.line
      ? `<line x1="${f(x)}" y1="${f(y - swatch / 2)}" x2="${f(x + swatch * 1.6)}" y2="${f(y - swatch / 2)}" stroke="${e.color}" stroke-width="2"/>`
      : `<rect x="${f(x)}" y="${f(y - swatch)}" width="${f(swatch)}" height="${f(swatch)}" fill="${e.color}"/>`) +
    text(ink, x + (e.line ? swatch * 1.6 : swatch) + 4, y, e.name, fs, 'start');
  if (pos === 'b' || pos === 't') {
    const total = entries.reduce((n, e) => n + itemW(e), 0);
    let x = box.x + Math.max(0, (box.w - total) / 2);
    const y = pos === 'b' ? box.y + box.h - fs * 0.4 : box.y + fs * 1.1;
    for (const e of entries) {
      out.push(key(x, y, e));
      x += itemW(e);
    }
    return pos === 'b' ? { ...box, h: box.h - fs * 1.8 } : { ...box, y: box.y + fs * 1.8, h: box.h - fs * 1.8 };
  }
  const width = Math.min(box.w * 0.4, Math.max(...entries.map(itemW)));
  const x = pos === 'l' ? box.x + 4 : box.x + box.w - width;
  let y = (pos === 'tr' ? box.y + fs : box.y + (box.h - entries.length * fs * 1.5) / 2) + fs;
  for (const e of entries) {
    out.push(key(x, y, e));
    y += fs * 1.5;
  }
  return pos === 'l' ? { ...box, x: box.x + width + 8, w: box.w - width - 8 } : { ...box, w: box.w - width - 8 };
}

function stackedRange(series: readonly RenderSeries[]): [number, number] {
  let lo = 0;
  let hi = 0;
  const n = Math.max(0, ...series.map((s) => s.values.length));
  for (let i = 0; i < n; i++) {
    let pos = 0;
    let neg = 0;
    for (const s of series) {
      const v = s.values[i] ?? 0;
      if (v >= 0) pos += v;
      else neg += v;
    }
    hi = Math.max(hi, pos);
    lo = Math.min(lo, neg);
  }
  return [lo, hi];
}

function flatRange(series: readonly RenderSeries[]): [number, number] {
  let lo = 0;
  let hi = 0;
  for (const s of series)
    for (const v of s.values) {
      if (v === null) continue;
      hi = Math.max(hi, v);
      lo = Math.min(lo, v);
    }
  return [lo, hi];
}

function polyline(points: Array<[number, number] | null>, smooth: boolean): string {
  let d = '';
  let prev: [number, number] | null = null;
  for (const p of points) {
    if (!p) {
      prev = null;
      continue;
    }
    if (!prev) d += `M${f(p[0])},${f(p[1])}`;
    else if (smooth) {
      const mx = (prev[0] + p[0]) / 2;
      d += `C${f(mx)},${f(prev[1])} ${f(mx)},${f(p[1])} ${f(p[0])},${f(p[1])}`;
    } else d += `L${f(p[0])},${f(p[1])}`;
    prev = p;
  }
  return d;
}

function marker(x: number, y: number, color: string, r: number): string {
  return `<circle cx="${f(x)}" cy="${f(y)}" r="${f(r)}" fill="${color}" stroke="${color}"/>`;
}

function labelText(spec: LabelSpec, value: number, category: string, series: string, percent?: number): string {
  const parts: string[] = [];
  if (spec.series) parts.push(series);
  if (spec.category) parts.push(category);
  if (spec.value) parts.push(valueLabel(value));
  if (spec.percent && percent !== undefined) parts.push(`${Math.round(percent * 100)}%`);
  return parts.join(', ');
}

// ---- value/category axes ------------------------------------------------------------------

/** The plot frame of a category chart: where values and categories land on screen. */
interface CategoryFrame {
  readonly plot: Box;
  readonly n: number;
  readonly band: number;
  readonly horizontal: boolean;
  valPos(v: number): number;
  catCenter(i: number): number;
  readonly zero: number;
}

function drawCategoryAxes(out: string[], data: ChartData, box: Box, ink: Ink, n: number, range: [number, number], percent = false): CategoryFrame {
  const fs = ink.fs;
  const horizontal = data.horizontal;
  // Axis titles take a strip on their side first.
  let area = { ...box };
  if (data.valTitle) {
    if (horizontal) {
      out.push(text(ink, area.x + area.w / 2, area.y + area.h - fs * 0.3, data.valTitle, fs * 1.05, 'middle', 'font-weight="600"'));
      area = { ...area, h: area.h - fs * 1.6 };
    } else {
      out.push(text(ink, area.x + fs, area.y + area.h / 2, data.valTitle, fs * 1.05, 'middle', `font-weight="600" transform="rotate(-90 ${f(area.x + fs)} ${f(area.y + area.h / 2)})"`));
      area = { ...area, x: area.x + fs * 1.6, w: area.w - fs * 1.6 };
    }
  }
  if (data.catTitle) {
    if (horizontal) {
      out.push(text(ink, area.x + fs, area.y + area.h / 2, data.catTitle, fs * 1.05, 'middle', `font-weight="600" transform="rotate(-90 ${f(area.x + fs)} ${f(area.y + area.h / 2)})"`));
      area = { ...area, x: area.x + fs * 1.6, w: area.w - fs * 1.6 };
    } else {
      out.push(text(ink, area.x + area.w / 2, area.y + area.h - fs * 0.3, data.catTitle, fs * 1.05, 'middle', 'font-weight="600"'));
      area = { ...area, h: area.h - fs * 1.6 };
    }
  }
  const [lo, hi] = [data.valMin ?? range[0], data.valMax ?? range[1]];
  const valLen = horizontal ? area.w : area.h;
  const scale = niceScale(lo, hi, Math.max(2, Math.floor(valLen / (fs * 3))));
  const label = (v: number) => (percent ? `${Math.round(v * 100)}%` : tickLabel(v, scale.step));
  const catLabelW = Math.min(area.w * 0.3, Math.max(1, ...data.categories.map((c) => c.length)) * fs * 0.55 + 6);
  const valLabelW = Math.max(label(scale.max).length, label(scale.min).length) * fs * 0.6 + 8;
  const leftW = horizontal ? (data.catAxis ? catLabelW : 4) : data.valAxis ? valLabelW : 4;
  const bottomH = horizontal ? (data.valAxis ? fs * 1.6 : 4) : data.catAxis ? fs * 1.8 : 4;
  const plot: Box = { x: area.x + leftW, y: area.y + 6, w: Math.max(10, area.w - leftW - 8), h: Math.max(10, area.h - bottomH - 6) };
  if (data.plotFill) out.push(`<rect x="${f(plot.x)}" y="${f(plot.y)}" width="${f(plot.w)}" height="${f(plot.h)}" fill="${data.plotFill}"/>`);
  const span = scale.max - scale.min || 1;
  const valPos = (v: number): number => (horizontal ? plot.x + ((v - scale.min) / span) * plot.w : plot.y + plot.h - ((v - scale.min) / span) * plot.h);
  const gridLine = (p: number, color: string): string =>
    horizontal
      ? `<line x1="${f(p)}" y1="${f(plot.y)}" x2="${f(p)}" y2="${f(plot.y + plot.h)}" stroke="${color}"/>`
      : `<line x1="${f(plot.x)}" y1="${f(p)}" x2="${f(plot.x + plot.w)}" y2="${f(p)}" stroke="${color}"/>`;
  if (data.minorGridlines) {
    for (let v = scale.min; v <= scale.max + 1e-9; v += scale.step / 5) out.push(gridLine(valPos(v), ink.minor));
  }
  for (let v = scale.min; v <= scale.max + scale.step / 2; v += scale.step) {
    const p = valPos(v);
    if (data.majorGridlines) out.push(gridLine(p, ink.grid));
    if (!data.valAxis) continue;
    if (horizontal) out.push(text(ink, p, plot.y + plot.h + fs * 1.3, label(v), fs));
    else out.push(text(ink, plot.x - 6, p + fs * 0.35, label(v), fs, 'end'));
  }
  const band = (horizontal ? plot.h : plot.w) / Math.max(1, n);
  // Excel draws horizontal-bar categories bottom-up.
  const catCenter = (i: number): number => (horizontal ? plot.y + plot.h - (i + 0.5) * band : plot.x + (i + 0.5) * band);
  if (data.catGridlines) {
    for (let i = 0; i <= n; i++) {
      const p = horizontal ? plot.y + plot.h - i * band : plot.x + i * band;
      out.push(horizontal ? `<line x1="${f(plot.x)}" y1="${f(p)}" x2="${f(plot.x + plot.w)}" y2="${f(p)}" stroke="${ink.grid}"/>` : `<line x1="${f(p)}" y1="${f(plot.y)}" x2="${f(p)}" y2="${f(plot.y + plot.h)}" stroke="${ink.grid}"/>`);
    }
  }
  if (data.catAxis) {
    const labelEvery = Math.max(1, Math.ceil((n * fs * 3) / (horizontal ? plot.h * 2 : plot.w)));
    data.categories.forEach((c, i) => {
      if (i % labelEvery !== 0) return;
      if (horizontal) out.push(text(ink, plot.x - 6, catCenter(i) + fs * 0.35, c, fs, 'end'));
      else out.push(text(ink, catCenter(i), plot.y + plot.h + fs * 1.3, c, fs));
    });
  }
  const zero = valPos(Math.min(Math.max(0, scale.min), scale.max));
  out.push(
    horizontal
      ? `<line x1="${f(zero)}" y1="${f(plot.y)}" x2="${f(zero)}" y2="${f(plot.y + plot.h)}" stroke="${ink.grid}"/>`
      : `<line x1="${f(plot.x)}" y1="${f(zero)}" x2="${f(plot.x + plot.w)}" y2="${f(zero)}" stroke="${ink.grid}"/>`,
  );
  return { plot, n, band, horizontal, valPos, catCenter, zero };
}

function categoryCount(data: ChartData): number {
  return Math.max(data.categories.length, ...data.series.map((s) => s.values.length));
}

function drawCategoryChart(out: string[], data: ChartData, palette: ThemePalette, box: Box, ink: Ink): void {
  const n = categoryCount(data);
  if (n === 0) return;
  const stacked = data.grouping === 'stacked' || data.grouping === 'percentStacked';
  const percent = data.grouping === 'percentStacked';
  const fr = drawCategoryAxes(out, data, box, ink, n, percent ? [0, 1] : stacked ? stackedRange(data.series) : flatRange(data.series), percent);
  const totals: number[] = [];
  if (percent) for (let i = 0; i < n; i++) totals.push(data.series.reduce((sum, s) => sum + Math.abs(s.values[i] ?? 0), 0) || 1);
  const norm = (v: number, i: number): number => (percent ? v / (totals[i] ?? 1) : v);
  const labels: string[] = [];

  if (data.kind === 'bar') {
    const m = data.series.length;
    const gap = data.gapWidth / 100;
    const overlap = stacked ? 1 : data.overlap / 100;
    const w = fr.band / (m - (m - 1) * overlap + gap);
    const posBase = new Array<number>(n).fill(0);
    const negBase = new Array<number>(n).fill(0);
    data.series.forEach((s, si) => {
      for (let i = 0; i < n; i++) {
        const raw = s.values[i];
        if (raw === null || raw === undefined) continue;
        const v = norm(raw, i);
        const start = stacked ? (v >= 0 ? (posBase[i] ?? 0) : (negBase[i] ?? 0)) : 0;
        const end = start + v;
        if (stacked) {
          if (v >= 0) posBase[i] = end;
          else negBase[i] = end;
        }
        const offset = stacked ? 0 : si * w * (1 - overlap);
        const color = s.pointColors?.[i] ?? (data.varyColors && m === 1 ? seriesColor(i, palette) : s.color);
        const a = fr.valPos(start);
        const b = fr.valPos(end);
        let cx: number;
        let cy: number;
        if (fr.horizontal) {
          const y = fr.catCenter(i) + fr.band / 2 - (gap * w) / 2 - offset - w;
          out.push(`<rect x="${f(Math.min(a, b))}" y="${f(y)}" width="${f(Math.abs(b - a))}" height="${f(w)}" fill="${color}"/>`);
          cx = s.labels?.pos === 'ctr' ? (a + b) / 2 : s.labels?.pos === 'inBase' ? a + 14 : s.labels?.pos === 'inEnd' ? b - 14 : b + 14;
          cy = y + w / 2 + ink.fs * 0.35;
        } else {
          const x = fr.catCenter(i) - fr.band / 2 + (gap * w) / 2 + offset;
          out.push(`<rect x="${f(x)}" y="${f(Math.min(a, b))}" width="${f(w)}" height="${f(Math.abs(b - a))}" fill="${color}"/>`);
          cx = x + w / 2;
          cy = s.labels?.pos === 'ctr' ? (a + b) / 2 + ink.fs * 0.35 : s.labels?.pos === 'inBase' ? a - 4 : s.labels?.pos === 'inEnd' ? b + ink.fs * 1.1 : b - 4;
        }
        if (s.labels) labels.push(text(ink, cx, cy, labelText(s.labels, raw, data.categories[i] ?? '', s.name), ink.fs * 0.9));
      }
    });
    out.push(...labels);
    return;
  }

  const base = new Array<number>(n).fill(0);
  const layers: Array<{ s: RenderSeries; top: Array<number | null>; bottom: number[] }> = [];
  for (const s of data.series) {
    const top: Array<number | null> = [];
    const bottom: number[] = [];
    for (let i = 0; i < n; i++) {
      const raw = s.values[i];
      const b = stacked ? (base[i] ?? 0) : 0;
      bottom.push(b);
      if (raw === null || raw === undefined) {
        top.push(stacked ? b : null);
        continue;
      }
      const v = b + norm(raw, i);
      top.push(v);
      if (stacked) base[i] = v;
    }
    layers.push({ s, top, bottom });
  }
  const pt = (i: number, v: number): [number, number] => [fr.catCenter(i), fr.valPos(v)];
  if (data.kind === 'area') {
    // Painted back to front so the first series stays on top, like Excel.
    for (const { s, top, bottom } of [...layers].reverse()) {
      const upper = top.map((v, i) => pt(i, v ?? 0));
      const lower = bottom.map((v, i) => pt(i, v)).reverse();
      const d = `M${[...upper, ...lower].map(([x, y]) => `${f(x)},${f(y)}`).join('L')}Z`;
      out.push(`<path d="${d}" fill="${s.color}" fill-opacity="0.85"/>`);
    }
  } else if (data.kind === 'stock') {
    drawStock(out, data, fr);
    return;
  } else {
    for (const { s, top } of layers) {
      const pts = top.map((v, i) => (v === null ? null : pt(i, v)));
      if (s.line) out.push(`<path d="${polyline(pts, s.smooth)}" fill="none" stroke="${s.color}" stroke-width="2.25" stroke-linejoin="round" stroke-linecap="round"/>`);
      if (s.markers) for (const p of pts) if (p) out.push(marker(p[0], p[1], s.color, 3));
    }
  }
  for (const { s, top } of layers) {
    const spec = s.labels;
    if (!spec) continue;
    top.forEach((v, i) => {
      const raw = s.values[i];
      if (v === null || raw === null || raw === undefined) return;
      const [x, y] = pt(i, v);
      out.push(text(ink, x, spec.pos === 'b' ? y + ink.fs * 1.2 : spec.pos === 'ctr' ? y + ink.fs * 0.35 : y - 6, labelText(spec, raw, data.categories[i] ?? '', s.name), ink.fs * 0.9));
    });
  }
}

/** High-low lines across the series per category; with four series (open, high, low, close), up/down bars. */
function drawStock(out: string[], data: ChartData, fr: CategoryFrame): void {
  const values = (i: number): number[] => data.series.map((s) => s.values[i]).filter((v): v is number => typeof v === 'number');
  for (let i = 0; i < fr.n; i++) {
    const vs = values(i);
    if (vs.length < 2) continue;
    const x = fr.catCenter(i);
    out.push(`<line x1="${f(x)}" y1="${f(fr.valPos(Math.max(...vs)))}" x2="${f(x)}" y2="${f(fr.valPos(Math.min(...vs)))}" stroke="#595959" stroke-width="1.25"/>`);
    const open = data.series[0]?.values[i];
    const close = data.series[data.series.length - 1]?.values[i];
    if (data.upDownBars && typeof open === 'number' && typeof close === 'number') {
      const w = fr.band / 2.5;
      const a = fr.valPos(open);
      const b = fr.valPos(close);
      out.push(`<rect x="${f(x - w / 2)}" y="${f(Math.min(a, b))}" width="${f(w)}" height="${f(Math.max(1, Math.abs(b - a)))}" fill="${close >= open ? '#FFFFFF' : '#404040'}" stroke="#404040"/>`);
    } else if (typeof close === 'number') {
      const y = fr.valPos(close);
      out.push(`<line x1="${f(x)}" y1="${f(y)}" x2="${f(x + fr.band / 5)}" y2="${f(y)}" stroke="#595959" stroke-width="1.25"/>`);
    }
  }
}

// ---- XY charts ---------------------------------------------------------------------------------

function drawScatter(out: string[], data: ChartData, box: Box, ink: Ink): void {
  const fs = ink.fs;
  const xs: number[] = [];
  const ys: number[] = [];
  let maxSize = 0;
  for (const s of data.series) {
    for (let i = 0; i < s.values.length; i++) {
      const y = s.values[i];
      const x = s.xs?.[i];
      if (y === null || y === undefined || x === null || x === undefined) continue;
      xs.push(x);
      ys.push(y);
      maxSize = Math.max(maxSize, Math.abs(s.sizes?.[i] ?? 0));
    }
  }
  if (xs.length === 0) return;
  const ysc = niceScale(data.valMin ?? Math.min(0, ...ys), data.valMax ?? Math.max(...ys), Math.max(2, Math.floor(box.h / (fs * 3))));
  const labelW = data.valAxis ? tickLabel(ysc.max, ysc.step).length * fs * 0.6 + 8 : 4;
  const plot: Box = { x: box.x + labelW, y: box.y + 6, w: box.w - labelW - 10, h: box.h - (data.catAxis ? fs * 1.8 : 4) - 6 };
  if (data.plotFill) out.push(`<rect x="${f(plot.x)}" y="${f(plot.y)}" width="${f(plot.w)}" height="${f(plot.h)}" fill="${data.plotFill}"/>`);
  const xsc = niceScale(Math.min(0, ...xs), Math.max(...xs), Math.max(2, Math.floor(plot.w / (fs * 5))));
  const X = (v: number) => plot.x + ((v - xsc.min) / (xsc.max - xsc.min || 1)) * plot.w;
  const Y = (v: number) => plot.y + plot.h - ((v - ysc.min) / (ysc.max - ysc.min || 1)) * plot.h;
  for (let v = ysc.min; v <= ysc.max + ysc.step / 2; v += ysc.step) {
    if (data.majorGridlines) out.push(`<line x1="${f(plot.x)}" y1="${f(Y(v))}" x2="${f(plot.x + plot.w)}" y2="${f(Y(v))}" stroke="${ink.grid}"/>`);
    if (data.valAxis) out.push(text(ink, plot.x - 6, Y(v) + fs * 0.35, tickLabel(v, ysc.step), fs, 'end'));
  }
  for (let v = xsc.min; v <= xsc.max + xsc.step / 2; v += xsc.step) {
    if (data.catGridlines) out.push(`<line x1="${f(X(v))}" y1="${f(plot.y)}" x2="${f(X(v))}" y2="${f(plot.y + plot.h)}" stroke="${ink.grid}"/>`);
    if (data.catAxis) out.push(text(ink, X(v), plot.y + plot.h + fs * 1.3, tickLabel(v, xsc.step), fs));
  }
  out.push(`<line x1="${f(plot.x)}" y1="${f(plot.y + plot.h)}" x2="${f(plot.x + plot.w)}" y2="${f(plot.y + plot.h)}" stroke="${ink.grid}"/>`);
  // Bubble area is proportional to size; the largest bubble is about a quarter of the plot's short side.
  const maxR = Math.min(plot.w, plot.h) / 8;
  data.series.forEach((s) => {
    const pts = s.values.map((y, i): [number, number] | null => {
      const x = s.xs?.[i];
      return y === null || x === null || x === undefined ? null : [X(x), Y(y)];
    });
    if (s.line) out.push(`<path d="${polyline(pts, s.smooth)}" fill="none" stroke="${s.color}" stroke-width="2"/>`);
    pts.forEach((p, i) => {
      if (!p) return;
      const color = s.pointColors?.[i] ?? s.color;
      if (data.kind === 'bubble') {
        const r = maxSize > 0 ? Math.sqrt(Math.abs(s.sizes?.[i] ?? 0) / maxSize) * maxR : maxR / 2;
        out.push(`<circle cx="${f(p[0])}" cy="${f(p[1])}" r="${f(r)}" fill="${color}" fill-opacity="0.75" stroke="${color}"/>`);
      } else if (s.markers) out.push(marker(p[0], p[1], color, 3));
      const y = s.values[i];
      if (s.labels && typeof y === 'number') out.push(text(ink, p[0], p[1] - 6, labelText(s.labels, y, '', s.name), fs * 0.9));
    });
  });
}

// ---- pies -----------------------------------------------------------------------------------------

function drawPie(out: string[], data: ChartData, palette: ThemePalette, box: Box, ink: Ink): void {
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  const outer = Math.max(4, Math.min(box.w, box.h) / 2 - 6);
  const rings = data.kind === 'doughnut' || data.kind === 'sunburst' ? data.series : data.series.slice(0, 1);
  const hole = data.kind === 'doughnut' ? (outer * data.holeSize) / 100 : data.kind === 'sunburst' ? outer * 0.3 : 0;
  const ringW = (outer - hole) / Math.max(1, rings.length);
  rings.forEach((s, ri) => {
    const r1 = outer - ri * ringW;
    const r0 = r1 - ringW;
    const vals = s.values.map((v) => Math.max(0, v ?? 0));
    const total = vals.reduce((a, b) => a + b, 0);
    if (total <= 0) return;
    // Excel starts the first slice at 12 o'clock, turned by firstSliceAng, and goes clockwise.
    let angle = -Math.PI / 2 + (data.firstSliceAng * Math.PI) / 180;
    vals.forEach((v, i) => {
      if (v === 0) return;
      const sweep = (v / total) * Math.PI * 2;
      const color = s.pointColors?.[i] ?? seriesColor(i, palette);
      const a2 = angle + sweep;
      const p = (r: number, a: number) => `${f(cx + r * Math.cos(a))},${f(cy + r * Math.sin(a))}`;
      if (sweep >= Math.PI * 2 - 1e-9) {
        out.push(`<circle cx="${f(cx)}" cy="${f(cy)}" r="${f((r1 + r0) / 2)}" fill="none" stroke="${color}" stroke-width="${f(r0 > 0 ? r1 - r0 : r1 * 2)}"/>`);
      } else {
        const large = sweep > Math.PI ? 1 : 0;
        const d =
          r0 > 0
            ? `M${p(r1, angle)}A${f(r1)},${f(r1)} 0 ${large} 1 ${p(r1, a2)}L${p(r0, a2)}A${f(r0)},${f(r0)} 0 ${large} 0 ${p(r0, angle)}Z`
            : `M${f(cx)},${f(cy)}L${p(r1, angle)}A${f(r1)},${f(r1)} 0 ${large} 1 ${p(r1, a2)}Z`;
        out.push(`<path d="${d}" fill="${color}" stroke="#FFFFFF" stroke-width="1"/>`);
      }
      const spec: LabelSpec | undefined = s.labels ?? (data.kind === 'sunburst' ? { pos: 'ctr', value: false, percent: false, category: true, series: false } : undefined);
      if (spec) {
        const mid = angle + sweep / 2;
        const lr = spec.pos === 'outEnd' || spec.pos === 'bestFit' ? r1 + ink.fs * 0.9 : spec.pos === 'inEnd' ? r1 - ink.fs * 1.4 : (r1 + Math.max(r0, 0)) / 2;
        const label = labelText(spec, v, data.categories[i] ?? '', s.name, v / total);
        out.push(text({ ...ink, text: lr > r1 ? ink.text : '#FFFFFF' }, cx + lr * Math.cos(mid), cy + lr * Math.sin(mid) + ink.fs * 0.35, label, ink.fs * 0.9));
      }
      angle = a2;
    });
  });
}

function drawRadar(out: string[], data: ChartData, box: Box, ink: Ink): void {
  const n = categoryCount(data);
  if (n < 3) return;
  const fs = ink.fs;
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  const r = Math.max(4, Math.min(box.w, box.h) / 2 - fs * 1.5);
  const [lo, hi] = flatRange(data.series);
  const sc = niceScale(Math.min(0, lo), hi, 4);
  const at = (i: number, v: number): [number, number] => {
    const a = -Math.PI / 2 + (i / n) * Math.PI * 2;
    const k = ((v - sc.min) / (sc.max - sc.min || 1)) * r;
    return [cx + k * Math.cos(a), cy + k * Math.sin(a)];
  };
  for (let v = sc.min + sc.step; v <= sc.max + sc.step / 2; v += sc.step) {
    const ring = Array.from({ length: n }, (_, i) => at(i, v));
    out.push(`<path d="${polyline([...ring, ring[0] ?? null], false)}" fill="none" stroke="${ink.grid}"/>`);
  }
  data.categories.forEach((c, i) => {
    const [x, y] = at(i, sc.max + sc.step * 0.35);
    out.push(text(ink, x, y + fs * 0.35, c, fs));
  });
  for (const s of data.series) {
    const pts = Array.from({ length: n }, (_, i) => at(i, s.values[i] ?? 0));
    const d = polyline([...pts, pts[0] ?? null], false);
    if (data.radarFilled) out.push(`<path d="${d}Z" fill="${s.color}" fill-opacity="0.6" stroke="${s.color}"/>`);
    else out.push(`<path d="${d}" fill="none" stroke="${s.color}" stroke-width="2"/>`);
    if (s.markers) for (const p of pts) out.push(marker(p[0], p[1], s.color, 3));
  }
}

/** Surface / contour charts as a colour grid: rows are series, columns are categories. */
function drawSurface(out: string[], data: ChartData, palette: ThemePalette, box: Box, ink: Ink): void {
  const n = categoryCount(data);
  const m = data.series.length;
  if (n === 0 || m === 0) return;
  const [lo, hi] = flatRange(data.series);
  const sc = niceScale(lo, hi, 5);
  const bands = Math.max(1, Math.round((sc.max - sc.min) / sc.step));
  const labelW = Math.max(...data.series.map((s) => s.name.length), 1) * ink.fs * 0.55 + 6;
  const plot: Box = { x: box.x + labelW, y: box.y + 4, w: box.w - labelW - 4, h: box.h - ink.fs * 1.8 - 4 };
  const cw = plot.w / n;
  const ch = plot.h / m;
  data.series.forEach((s, si) => {
    const y = plot.y + plot.h - (si + 1) * ch;
    out.push(text(ink, plot.x - 4, y + ch / 2 + ink.fs * 0.35, s.name, ink.fs, 'end'));
    for (let i = 0; i < n; i++) {
      const v = s.values[i] ?? sc.min;
      const band = Math.min(bands - 1, Math.floor((v - sc.min) / sc.step));
      out.push(`<rect x="${f(plot.x + i * cw)}" y="${f(y)}" width="${f(cw + 0.5)}" height="${f(ch + 0.5)}" fill="${seriesColor(band, palette)}"/>`);
    }
  });
  data.categories.forEach((c, i) => out.push(text(ink, plot.x + (i + 0.5) * cw, plot.y + plot.h + ink.fs * 1.3, c, ink.fs)));
}

// ---- Excel 2016 charts --------------------------------------------------------------------------

function firstValues(data: ChartData): number[] {
  return (data.series[0]?.values ?? []).map((v) => v ?? 0);
}

function drawWaterfall(out: string[], data: ChartData, palette: ThemePalette, box: Box, ink: Ink): void {
  const vals = firstValues(data);
  const totals = new Set(data.subtotals);
  let run = 0;
  const steps = vals.map((v, i) => {
    const total = totals.has(i);
    const start = total ? 0 : run;
    const end = total ? v : run + v;
    run = end;
    return { start, end, total, v };
  });
  const lo = Math.min(0, ...steps.map((s) => Math.min(s.start, s.end)));
  const hi = Math.max(0, ...steps.map((s) => Math.max(s.start, s.end)));
  const fr = drawCategoryAxes(out, data, box, ink, vals.length, [lo, hi]);
  const w = fr.band / 1.5;
  // Excel's waterfall colours: increase, decrease, total.
  const [up, down, total] = [seriesColor(0, palette), seriesColor(1, palette), seriesColor(2, palette)];
  steps.forEach((s, i) => {
    const a = fr.valPos(s.start);
    const b = fr.valPos(s.end);
    const x = fr.catCenter(i) - w / 2;
    out.push(`<rect x="${f(x)}" y="${f(Math.min(a, b))}" width="${f(w)}" height="${f(Math.max(1, Math.abs(b - a)))}" fill="${s.total ? total : s.v >= 0 ? up : down}"/>`);
    out.push(text(ink, x + w / 2, Math.min(a, b) - 4, valueLabel(s.v), ink.fs * 0.9));
    const next = steps[i + 1];
    if (next && !next.total) out.push(`<line x1="${f(x + w)}" y1="${f(b)}" x2="${f(fr.catCenter(i + 1) - w / 2)}" y2="${f(b)}" stroke="${ink.grid}"/>`);
  });
}

function drawFunnel(out: string[], data: ChartData, box: Box, ink: Ink): void {
  const vals = firstValues(data);
  if (vals.length === 0) return;
  const max = Math.max(...vals.map(Math.abs), 1);
  const labelW = Math.max(...data.categories.map((c) => c.length), 1) * ink.fs * 0.55 + 8;
  const plot: Box = { x: box.x + labelW, y: box.y + 4, w: box.w - labelW - 4, h: box.h - 8 };
  const band = plot.h / vals.length;
  const color = data.series[0]?.color ?? '#4472C4';
  vals.forEach((v, i) => {
    const w = (Math.abs(v) / max) * plot.w;
    const y = plot.y + i * band + band * 0.1;
    out.push(`<rect x="${f(plot.x + (plot.w - w) / 2)}" y="${f(y)}" width="${f(w)}" height="${f(band * 0.8)}" fill="${color}"/>`);
    out.push(text({ ...ink, text: '#FFFFFF' }, plot.x + plot.w / 2, y + band * 0.4 + ink.fs * 0.35, valueLabel(v), ink.fs * 0.9));
    out.push(text(ink, plot.x - 6, y + band * 0.4 + ink.fs * 0.35, data.categories[i] ?? '', ink.fs, 'end'));
  });
}

/** Squarified treemap layout (Bruls et al.): rows of tiles whose aspect ratios stay close to 1. */
export function squarify(values: readonly number[], box: Box): Box[] {
  const total = values.reduce((a, b) => a + Math.max(0, b), 0);
  const out: Box[] = new Array<Box>(values.length);
  if (total <= 0) return values.map(() => ({ x: box.x, y: box.y, w: 0, h: 0 }));
  const scale = (box.w * box.h) / total;
  const items = values.map((v, i) => ({ i, a: Math.max(0, v) * scale })).sort((p, q) => q.a - p.a);
  let rect = { ...box };
  let row: typeof items = [];
  const worst = (r: typeof items, side: number): number => {
    const s = r.reduce((n, it) => n + it.a, 0);
    const max = Math.max(...r.map((it) => it.a));
    const min = Math.min(...r.map((it) => it.a));
    return Math.max((side * side * max) / (s * s), (s * s) / (side * side * min));
  };
  const place = (r: typeof items): void => {
    const s = r.reduce((n, it) => n + it.a, 0);
    const horizontal = rect.w >= rect.h;
    const thick = s / (horizontal ? rect.h : rect.w);
    let pos = horizontal ? rect.y : rect.x;
    for (const it of r) {
      const len = it.a / thick;
      out[it.i] = horizontal ? { x: rect.x, y: pos, w: thick, h: len } : { x: pos, y: rect.y, w: len, h: thick };
      pos += len;
    }
    rect = horizontal ? { x: rect.x + thick, y: rect.y, w: rect.w - thick, h: rect.h } : { x: rect.x, y: rect.y + thick, w: rect.w, h: rect.h - thick };
  };
  for (const it of items) {
    if (it.a <= 0) {
      out[it.i] = { x: rect.x, y: rect.y, w: 0, h: 0 };
      continue;
    }
    const side = Math.min(rect.w, rect.h);
    if (row.length === 0 || worst([...row, it], side) <= worst(row, side)) row.push(it);
    else {
      place(row);
      row = [it];
    }
  }
  if (row.length > 0) place(row);
  return out;
}

function drawTreemap(out: string[], data: ChartData, palette: ThemePalette, box: Box, ink: Ink): void {
  const vals = firstValues(data);
  squarify(vals, box).forEach((r, i) => {
    if (r.w <= 0 || r.h <= 0) return;
    out.push(`<rect x="${f(r.x)}" y="${f(r.y)}" width="${f(r.w)}" height="${f(r.h)}" fill="${seriesColor(i, palette)}" stroke="#FFFFFF" stroke-width="1.5"/>`);
    if (r.w > ink.fs * 3 && r.h > ink.fs * 1.5) out.push(text({ ...ink, text: '#FFFFFF' }, r.x + 4, r.y + ink.fs * 1.2, data.categories[i] ?? '', ink.fs, 'start'));
  });
}

/** Histogram bins by Scott's rule, the width Excel picks for an automatic bin count. */
export function histogramBins(values: readonly number[]): Array<{ lo: number; hi: number; count: number }> {
  if (values.length === 0) return [];
  const min = Math.min(...values);
  const max = Math.max(...values);
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  const sd = Math.sqrt(values.reduce((a, b) => a + (b - mean) ** 2, 0) / Math.max(1, values.length - 1));
  const width = sd > 0 ? (3.5 * sd) / Math.cbrt(values.length) : 1;
  const count = Math.max(1, Math.min(50, Math.ceil((max - min) / width) || 1));
  const bins = Array.from({ length: count }, (_, i) => ({ lo: min + i * width, hi: min + (i + 1) * width, count: 0 }));
  for (const v of values) {
    const i = Math.min(count - 1, Math.max(0, Math.ceil((v - min) / width) - 1));
    const bin = bins[i];
    if (bin) bin.count++;
  }
  return bins;
}

function drawHistogram(out: string[], data: ChartData, palette: ThemePalette, box: Box, ink: Ink, pareto: boolean): void {
  const raw = firstValues(data);
  let labels: string[];
  let counts: number[];
  if (pareto && data.categories.length > 0) {
    // Pareto over categories: bars sorted descending with a cumulative-percent line.
    const pairs = data.categories.map((c, i) => ({ c, v: raw[i] ?? 0 })).sort((a, b) => b.v - a.v);
    labels = pairs.map((p) => p.c);
    counts = pairs.map((p) => p.v);
  } else {
    const bins = histogramBins(raw);
    labels = bins.map((b, i) => `${i === 0 ? '[' : '('}${valueLabel(b.lo)}, ${valueLabel(b.hi)}]`);
    counts = bins.map((b) => b.count);
    if (pareto) {
      const order = counts.map((c, i) => ({ c, l: labels[i] ?? '' })).sort((a, b) => b.c - a.c);
      counts = order.map((o) => o.c);
      labels = order.map((o) => o.l);
    }
  }
  const view: ChartData = { ...data, categories: labels };
  const fr = drawCategoryAxes(out, view, box, ink, counts.length, [0, Math.max(0, ...counts)]);
  const w = fr.band * (pareto ? 0.8 : 1);
  counts.forEach((c, i) => {
    const y = fr.valPos(c);
    out.push(`<rect x="${f(fr.catCenter(i) - w / 2)}" y="${f(y)}" width="${f(w)}" height="${f(fr.zero - y)}" fill="${data.series[0]?.color ?? seriesColor(0, palette)}" stroke="#FFFFFF" stroke-width="0.5"/>`);
  });
  if (!pareto) return;
  const total = counts.reduce((a, b) => a + b, 0) || 1;
  let acc = 0;
  const pts = counts.map((c, i): [number, number] => {
    acc += c;
    return [fr.catCenter(i), fr.plot.y + fr.plot.h - (acc / total) * fr.plot.h];
  });
  const lineColor = seriesColor(1, palette);
  out.push(`<path d="${polyline(pts, false)}" fill="none" stroke="${lineColor}" stroke-width="2"/>`);
  for (const p of pts) out.push(marker(p[0], p[1], lineColor, 2.5));
  for (const pct of [0, 0.5, 1]) out.push(text(ink, fr.plot.x + fr.plot.w + 4, fr.plot.y + fr.plot.h - pct * fr.plot.h + ink.fs * 0.35, `${pct * 100}%`, ink.fs, 'start'));
}

function quartile(sorted: readonly number[], q: number): number {
  // Exclusive method, Excel's default for box & whisker.
  const pos = q * (sorted.length + 1) - 1;
  const lo = Math.min(sorted.length - 1, Math.max(0, Math.floor(pos)));
  const hi = Math.min(sorted.length - 1, lo + 1);
  const a = sorted[lo] ?? 0;
  const b = sorted[hi] ?? a;
  return a + (b - a) * Math.min(1, Math.max(0, pos - lo));
}

function drawBoxWhisker(out: string[], data: ChartData, palette: ThemePalette, box: Box, ink: Ink): void {
  const raw = data.series[0]?.values ?? [];
  // One box per distinct category, like Excel; no categories means one box for all values.
  const groups = new Map<string, number[]>();
  raw.forEach((v, i) => {
    if (v === null) return;
    const key = data.categories[i] ?? '';
    const list = groups.get(key) ?? [];
    list.push(v);
    groups.set(key, list);
  });
  const names = [...groups.keys()];
  const all = raw.filter((v): v is number => v !== null);
  if (all.length === 0) return;
  const fr = drawCategoryAxes(out, { ...data, categories: names }, box, ink, names.length, [Math.min(0, ...all), Math.max(...all)]);
  const color = data.series[0]?.color ?? seriesColor(0, palette);
  names.forEach((name, i) => {
    const sorted = [...(groups.get(name) ?? [])].sort((a, b) => a - b);
    const q1 = quartile(sorted, 0.25);
    const q2 = quartile(sorted, 0.5);
    const q3 = quartile(sorted, 0.75);
    const iqr = q3 - q1;
    const inside = sorted.filter((v) => v >= q1 - 1.5 * iqr && v <= q3 + 1.5 * iqr);
    const lo = inside[0] ?? q1;
    const hi = inside[inside.length - 1] ?? q3;
    const x = fr.catCenter(i);
    const w = fr.band / 3;
    out.push(`<line x1="${f(x)}" y1="${f(fr.valPos(hi))}" x2="${f(x)}" y2="${f(fr.valPos(lo))}" stroke="#595959"/>`);
    out.push(`<rect x="${f(x - w / 2)}" y="${f(fr.valPos(q3))}" width="${f(w)}" height="${f(Math.max(1, fr.valPos(q1) - fr.valPos(q3)))}" fill="${color}" stroke="#595959"/>`);
    out.push(`<line x1="${f(x - w / 2)}" y1="${f(fr.valPos(q2))}" x2="${f(x + w / 2)}" y2="${f(fr.valPos(q2))}" stroke="#FFFFFF" stroke-width="1.5"/>`);
    const mean = sorted.reduce((a, b) => a + b, 0) / sorted.length;
    out.push(`<text x="${f(x)}" y="${f(fr.valPos(mean) + 3)}" font-size="${f(ink.fs)}" fill="#FFFFFF" text-anchor="middle">×</text>`);
    for (const v of sorted) if (v < lo || v > hi) out.push(`<circle cx="${f(x)}" cy="${f(fr.valPos(v))}" r="2.5" fill="none" stroke="${color}"/>`);
  });
}

/** Height of the title strip, so the drawing layer can place an inline title editor over it. */
export function titleBandHeight(width: number, height: number): number {
  return chartFontSize(width, height) * 2.4 + 6;
}

function chartFontSize(width: number, height: number): number {
  return Math.max(8, Math.min(13, Math.min(width, height) / 22));
}

/** The whole chart as an `<svg>` string of the given pixel size. */
export function renderChartSvg(data: ChartData, width: number, height: number, palette: ThemePalette): string {
  const fs = chartFontSize(width, height);
  const dark = data.chartFill !== null && luminance(data.chartFill) < 0.45;
  const ink: Ink = dark ? { text: '#F2F2F2', grid: '#7F7F7F', minor: '#595959', fs } : { text: '#595959', grid: '#D9D9D9', minor: '#EFEFEF', fs };
  const out: string[] = [];
  const lw = data.chartLine?.width ?? 0;
  out.push(
    `<rect x="${f(lw / 2)}" y="${f(lw / 2)}" width="${f(width - lw)}" height="${f(height - lw)}" fill="${data.chartFill ?? 'none'}"${data.chartLine ? ` stroke="${data.chartLine.color}" stroke-width="${f(lw)}"` : ''}/>`,
  );
  let box: Box = { x: 8, y: 6, w: width - 16, h: height - 12 };
  const title = data.title ?? (data.series.length === 1 && (data.kind === 'pie' || data.kind === 'doughnut') ? data.series[0]?.name : undefined);
  if (title) {
    out.push(text(ink, width / 2, box.y + fs * 1.5, title, fs * 1.4, 'middle'));
    box = { ...box, y: box.y + fs * 2.4, h: box.h - fs * 2.4 };
  }
  box = drawLegend(out, data, palette, box, ink);
  switch (data.kind) {
    case 'pie':
    case 'doughnut':
    case 'sunburst':
      drawPie(out, data, palette, box, ink);
      break;
    case 'scatter':
    case 'bubble':
      drawScatter(out, data, box, ink);
      break;
    case 'radar':
      drawRadar(out, data, box, ink);
      break;
    case 'surface':
      drawSurface(out, data, palette, box, ink);
      break;
    case 'waterfall':
      drawWaterfall(out, data, palette, box, ink);
      break;
    case 'funnel':
      drawFunnel(out, data, box, ink);
      break;
    case 'treemap':
      drawTreemap(out, data, palette, box, ink);
      break;
    case 'histogram':
    case 'pareto':
      drawHistogram(out, data, palette, box, ink, data.kind === 'pareto');
      break;
    case 'boxWhisker':
      drawBoxWhisker(out, data, palette, box, ink);
      break;
    case 'bar':
    case 'line':
    case 'area':
    case 'stock':
      drawCategoryChart(out, data, palette, box, ink);
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" font-family="${esc(FONT)}">${out.join('')}</svg>`;
}
