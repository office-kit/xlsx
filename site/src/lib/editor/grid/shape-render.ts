// Drawing shapes → SVG. Preset geometries are drawn from simplified versions
// of their ECMA-376 presetShapeDefinitions at the default adjust values; a
// preset outside the table falls back to its bounding rectangle. Colours
// resolve against the workbook theme, with the shape's `<xdr:style>` theme
// references filling in whatever `spPr` leaves unset, as Excel does.

import type { DmlColorWithMods, Fill, LineProperties, ShapeReference, TextBody } from '@office-kit/xlsx/drawing';
import { hslToRgb, rgbToHsl, type ThemePalette } from '../core/theme.ts';

/** EMU per CSS pixel at 96 dpi. */
const EMU_PER_PX = 9525;

// SpreadsheetML theme indices: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink.
const SCHEME_SLOT: Readonly<Record<string, number>> = {
  bg1: 0,
  lt1: 0,
  tx1: 1,
  dk1: 1,
  bg2: 2,
  lt2: 2,
  tx2: 3,
  dk2: 3,
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9,
  hlink: 10,
  folHlink: 11,
};

const PRESET_COLORS: Readonly<Record<string, string>> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
};

/** Theme line widths (Office theme `a:lnStyleLst`), by `lnRef idx`. */
const THEME_LINE_EMU = [0, 6350, 12700, 19050];

export interface CssColor {
  readonly color: string;
  readonly opacity: number;
}

const clamp01 = (v: number): number => Math.min(1, Math.max(0, v));

/** Apply `f(channel, k)` to each RGB channel of `hex`. */
function mapChannels(hex: string, k: number, f: (v: number, k: number) => number): string {
  return [0, 2, 4]
    .map((i) =>
      Math.round(Math.min(255, Math.max(0, f(Number.parseInt(hex.slice(i, i + 2), 16), k))))
        .toString(16)
        .padStart(2, '0'),
    )
    .join('');
}

/** A DrawingML colour with its modifiers applied, as CSS. */
export function resolveDmlColor(c: DmlColorWithMods, palette: ThemePalette): CssColor {
  const b = c.base;
  let hex: string;
  switch (b.kind) {
    case 'srgb':
      hex = b.value;
      break;
    case 'schemeClr':
      hex = palette[SCHEME_SLOT[b.value] ?? 4] ?? '4472C4';
      break;
    case 'sysClr':
      hex = b.lastClr ?? (b.value === 'window' ? 'FFFFFF' : '000000');
      break;
    case 'prstClr':
      hex = PRESET_COLORS[b.value] ?? '000000';
      break;
    case 'hslClr':
      hex = hslToRgb(b.hue / 21_600_000, b.sat / 100_000, b.lum / 100_000);
      break;
    case 'scrgbClr': {
      const ch = (v: number) =>
        Math.round(clamp01(v / 100_000) * 255)
          .toString(16)
          .padStart(2, '0');
      hex = `${ch(b.r)}${ch(b.g)}${ch(b.b)}`;
      break;
    }
  }
  let [h, s, l] = rgbToHsl(hex);
  let opacity = 1;
  for (const m of c.mods) {
    switch (m.kind) {
      case 'lumMod':
        l = clamp01(l * (m.val / 100_000));
        hex = hslToRgb(h, s, l);
        break;
      case 'lumOff':
        l = clamp01(l + m.val / 100_000);
        hex = hslToRgb(h, s, l);
        break;
      case 'satMod':
        s = clamp01(s * (m.val / 100_000));
        hex = hslToRgb(h, s, l);
        break;
      case 'shade':
        hex = mapChannels(hex, m.val / 100_000, (v, k) => v * k);
        [h, s, l] = rgbToHsl(hex);
        break;
      case 'tint':
        hex = mapChannels(hex, m.val / 100_000, (v, k) => v + (255 - v) * (1 - k));
        [h, s, l] = rgbToHsl(hex);
        break;
      case 'alpha':
        opacity = clamp01(m.val / 100_000);
        break;
      default:
        break;
    }
  }
  return { color: `#${hex.toUpperCase()}`, opacity };
}

const fillColor = (fill: Fill, palette: ThemePalette): CssColor | null => {
  switch (fill.kind) {
    case 'noFill':
      return null;
    case 'solidFill':
      return resolveDmlColor(fill.color, palette);
    case 'gradFill': {
      const first = fill.stops[0];
      return first ? resolveDmlColor(first.color, palette) : null;
    }
    case 'pattFill':
      return fill.fgClr ? resolveDmlColor(fill.fgClr, palette) : null;
    case 'blipFill':
    case 'grpFill':
      return null;
  }
};

export interface ShapePaint {
  readonly fill: CssColor | null;
  readonly stroke: CssColor | null;
  readonly strokeWidth: number;
  readonly dash: string | undefined;
  readonly headEnd: boolean;
  readonly tailEnd: boolean;
  readonly fontColor: string;
}

const DASHES: Readonly<Record<string, number[]>> = {
  dot: [1, 1],
  sysDot: [1, 1],
  dash: [4, 3],
  sysDash: [3, 1],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  sysDashDot: [3, 1, 1, 1],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

const hasEnd = (e: LineProperties['headEnd']): boolean => e !== undefined && e.type !== undefined && e.type !== 'none';

/** Resolved fill, outline and default text colour of a shape. */
export function shapePaint(shape: ShapeReference, palette: ThemePalette): ShapePaint {
  const { spPr, style } = shape;
  const fill = spPr.fill
    ? fillColor(spPr.fill, palette)
    : style && style.fillRef.idx > 0 && style.fillRef.color && !shape.connector
      ? resolveDmlColor(style.fillRef.color, palette)
      : null;
  const ln = spPr.ln;
  let stroke: CssColor | null = null;
  if (ln?.fill) stroke = fillColor(ln.fill, palette);
  else if (style && style.lnRef.idx > 0 && style.lnRef.color) stroke = resolveDmlColor(style.lnRef.color, palette);
  const widthEmu = ln?.w ?? THEME_LINE_EMU[Math.min(3, style?.lnRef.idx ?? 0)] ?? 9525;
  const dashPattern = ln?.dash?.kind === 'preset' ? DASHES[ln.dash.val] : ln?.dash?.kind === 'custDash' ? ln.dash.pattern.map((p) => p / 100_000) : undefined;
  const strokeWidth = Math.max(0.75, widthEmu / EMU_PER_PX);
  const fontRefColor = style?.fontRef.color ? resolveDmlColor(style.fontRef.color, palette).color : undefined;
  return {
    fill,
    stroke,
    strokeWidth,
    dash: dashPattern?.map((d) => d * strokeWidth).join(' '),
    headEnd: hasEnd(ln?.headEnd),
    tailEnd: hasEnd(ln?.tailEnd),
    fontColor: fontRefColor ?? `#${palette[1] ?? '000000'}`,
  };
}

// ---- geometry -----------------------------------------------------------------

type Pt = readonly [number, number];

const poly = (pts: readonly Pt[]): string => `M${pts.map(([x, y]) => `${fmt(x)},${fmt(y)}`).join(' L')} Z`;
const fmt = (n: number): string => (Math.round(n * 100) / 100).toString();

function roundRectPath(w: number, h: number, r: number): string {
  const k = Math.min(r, w / 2, h / 2);
  return `M${fmt(k)},0 H${fmt(w - k)} A${fmt(k)},${fmt(k)} 0 0 1 ${fmt(w)},${fmt(k)} V${fmt(h - k)} A${fmt(k)},${fmt(k)} 0 0 1 ${fmt(w - k)},${fmt(h)} H${fmt(k)} A${fmt(k)},${fmt(k)} 0 0 1 0,${fmt(h - k)} V${fmt(k)} A${fmt(k)},${fmt(k)} 0 0 1 ${fmt(k)},0 Z`;
}

function ellipsePath(w: number, h: number): string {
  return `M0,${fmt(h / 2)} A${fmt(w / 2)},${fmt(h / 2)} 0 1 1 ${fmt(w)},${fmt(h / 2)} A${fmt(w / 2)},${fmt(h / 2)} 0 1 1 0,${fmt(h / 2)} Z`;
}

/** Star with `n` points and the inner radius as a fraction of the outer one. */
function starPath(w: number, h: number, n: number, inner: number): string {
  const pts: Pt[] = [];
  for (let i = 0; i < n * 2; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / n;
    const k = i % 2 === 0 ? 1 : inner;
    pts.push([w / 2 + (w / 2) * k * Math.cos(a), h / 2 + (h / 2) * k * Math.sin(a)]);
  }
  return poly(pts);
}

/** A right-pointing block arrow in a `len` × `thick` box, shaft half the thickness. */
function arrowPoints(len: number, thick: number): Pt[] {
  const head = Math.min(len, thick * 0.5);
  const t1 = thick / 4;
  const t2 = (thick * 3) / 4;
  return [
    [0, t1],
    [len - head, t1],
    [len - head, 0],
    [len, thick / 2],
    [len - head, thick],
    [len - head, t2],
    [0, t2],
  ];
}

/** Ellipse with a wedge to `tip`, leaving the ellipse between two angles near the tip. */
function wedgeEllipsePath(w: number, h: number, tip: Pt): string {
  const rx = w / 2;
  const ry = h / 2;
  const at = (deg: number): Pt => [rx + rx * Math.cos((deg * Math.PI) / 180), ry + ry * Math.sin((deg * Math.PI) / 180)];
  const a = Math.atan2((tip[1] - ry) / ry, (tip[0] - rx) / rx) * (180 / Math.PI);
  const [x1, y1] = at(a - 12);
  const [x2, y2] = at(a + 12);
  return `M${fmt(x2)},${fmt(y2)} A${fmt(rx)},${fmt(ry)} 0 1 1 ${fmt(x1)},${fmt(y1)} L${fmt(tip[0])},${fmt(tip[1])} Z`;
}

/** Rectangle callout with its tail at Excel's default spot below the lower-left quarter. */
function wedgeRectPath(w: number, h: number, rounded: boolean): string {
  const tip: Pt = [w * 0.2917, h * 1.125];
  const r = rounded ? Math.min(w, h) * 0.16667 : 0;
  const x1 = w / 6;
  const x2 = w / 3;
  if (r === 0) return poly([[0, 0], [w, 0], [w, h], [x2, h], tip, [x1, h], [0, h]]);
  return (
    `M${fmt(r)},0 H${fmt(w - r)} A${fmt(r)},${fmt(r)} 0 0 1 ${fmt(w)},${fmt(r)} V${fmt(h - r)} A${fmt(r)},${fmt(r)} 0 0 1 ${fmt(w - r)},${fmt(h)}` +
    ` H${fmt(x2)} L${fmt(tip[0])},${fmt(tip[1])} L${fmt(x1)},${fmt(h)} H${fmt(r)} A${fmt(r)},${fmt(r)} 0 0 1 0,${fmt(h - r)} V${fmt(r)} A${fmt(r)},${fmt(r)} 0 0 1 ${fmt(r)},0 Z`
  );
}

/** Lines and connectors: drawn as an open stroke from the top-left to the bottom-right corner. */
export const LINE_PRESETS: ReadonlySet<string> = new Set(['line', 'straightConnector1', 'lineInv']);

/** SVG path data for preset `prst` in a `w` × `h` box. */
export function presetPath(prst: string, w: number, h: number): string {
  const ss = Math.min(w, h);
  switch (prst) {
    case 'line':
    case 'straightConnector1':
      return `M0,0 L${fmt(w)},${fmt(h)}`;
    case 'lineInv':
      return `M0,${fmt(h)} L${fmt(w)},0`;
    case 'roundRect':
      return roundRectPath(w, h, ss * 0.16667);
    case 'flowChartTerminator':
      return roundRectPath(w, h, h / 2);
    case 'ellipse':
    case 'flowChartConnector':
      return ellipsePath(w, h);
    case 'triangle':
      return poly([[w / 2, 0], [w, h], [0, h]]);
    case 'rtTriangle':
      return poly([[0, 0], [w, h], [0, h]]);
    case 'diamond':
    case 'flowChartDecision':
      return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': {
      const o = ss * 0.25;
      return poly([[o, 0], [w, 0], [w - o, h], [0, h]]);
    }
    case 'trapezoid': {
      const o = ss * 0.25;
      return poly([[0, h], [o, 0], [w - o, 0], [w, h]]);
    }
    case 'pentagon':
      return poly([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]);
    case 'hexagon': {
      const o = ss * 0.25;
      return poly([[o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h], [0, h / 2]]);
    }
    case 'octagon': {
      const o = ss * 0.2929;
      return poly([[o, 0], [w - o, 0], [w, o], [w, h - o], [w - o, h], [o, h], [0, h - o], [0, o]]);
    }
    case 'plus':
    case 'mathPlus': {
      const o = ss * 0.25;
      return poly([[o, 0], [w - o, 0], [w - o, o], [w, o], [w, h - o], [w - o, h - o], [w - o, h], [o, h], [o, h - o], [0, h - o], [0, o], [o, o]]);
    }
    case 'rightArrow':
      return poly(arrowPoints(w, h));
    case 'leftArrow':
      return poly(arrowPoints(w, h).map(([x, y]) => [w - x, y] as const));
    case 'downArrow':
      return poly(arrowPoints(h, w).map(([x, y]) => [y, x] as const));
    case 'upArrow':
      return poly(arrowPoints(h, w).map(([x, y]) => [y, h - x] as const));
    case 'leftRightArrow': {
      const head = Math.min(w / 2, h * 0.5);
      return poly([[0, h / 2], [head, 0], [head, h / 4], [w - head, h / 4], [w - head, 0], [w, h / 2], [w - head, h], [w - head, (h * 3) / 4], [head, (h * 3) / 4], [head, h]]);
    }
    case 'chevron': {
      const o = Math.min(w, ss * 0.5);
      return poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h], [o, h / 2]]);
    }
    case 'homePlate': {
      const o = Math.min(w, ss * 0.5);
      return poly([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h]]);
    }
    case 'star4':
      return starPath(w, h, 4, 0.25);
    case 'star5':
      return starPath(w, h, 5, 0.382);
    case 'star6':
      return starPath(w, h, 6, 0.577);
    case 'star8':
      return starPath(w, h, 8, 0.765);
    case 'wedgeRectCallout':
      return wedgeRectPath(w, h, false);
    case 'wedgeRoundRectCallout':
      return wedgeRectPath(w, h, true);
    case 'wedgeEllipseCallout':
      return wedgeEllipsePath(w, h, [w * 0.2917, h * 1.125]);
    default:
      return poly([[0, 0], [w, 0], [w, h], [0, h]]);
  }
}

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

const paintAttr = (name: 'fill' | 'stroke', c: CssColor | null): string =>
  c ? ` ${name}="${c.color}"${c.opacity < 1 ? ` ${name}-opacity="${c.opacity}"` : ''}` : ` ${name}="none"`;

/**
 * The shape's geometry as an SVG document of `w` × `h` pixels. Callouts draw
 * their tails outside that box, so the caller lets the SVG overflow.
 */
export function renderShapeSvg(shape: ShapeReference, w: number, h: number, palette: ThemePalette, idPrefix: string): string {
  const paint = shapePaint(shape, palette);
  const geom = shape.spPr.geometry;
  const prst = geom?.kind === 'preset' ? geom.prst : 'rect';
  const isLine = shape.connector === true || LINE_PRESETS.has(prst);
  const xfrm = shape.spPr.xfrm;
  const transforms: string[] = [];
  if (xfrm?.rot) transforms.push(`rotate(${fmt(xfrm.rot / 60_000)} ${fmt(w / 2)} ${fmt(h / 2)})`);
  if (xfrm?.flipH) transforms.push(`translate(${fmt(w)} 0) scale(-1 1)`);
  if (xfrm?.flipV) transforms.push(`translate(0 ${fmt(h)}) scale(1 -1)`);
  const stroke = paint.stroke;
  const defs: string[] = [];
  const markerAttrs: string[] = [];
  if (isLine && stroke && (paint.headEnd || paint.tailEnd)) {
    const id = `${idPrefix}-arrow`;
    defs.push(
      `<marker id="${esc(id)}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 Z" fill="${stroke.color}"/></marker>`,
    );
    if (paint.headEnd) markerAttrs.push(`marker-start="url(#${esc(id)})"`);
    if (paint.tailEnd) markerAttrs.push(`marker-end="url(#${esc(id)})"`);
  }
  return [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(w)}" height="${fmt(h)}" overflow="visible">`,
    defs.length > 0 ? `<defs>${defs.join('')}</defs>` : '',
    transforms.length > 0 ? `<g transform="${transforms.join(' ')}">` : '<g>',
    `<path d="${presetPath(prst, w, h)}"`,
    paintAttr('fill', isLine ? null : paint.fill),
    paintAttr('stroke', stroke),
    ` stroke-width="${fmt(paint.strokeWidth)}"`,
    paint.dash ? ` stroke-dasharray="${paint.dash}"` : '',
    ' stroke-linejoin="round"',
    markerAttrs.length > 0 ? ` ${markerAttrs.join(' ')}` : '',
    '/></g></svg>',
  ].join('');
}

// ---- text ---------------------------------------------------------------------

export interface TextLayout {
  /** CSS `justify-content` for the body's vertical anchor. */
  readonly justify: 'flex-start' | 'center' | 'flex-end';
  /** Insets in CSS px at 100% zoom: top, right, bottom, left. */
  readonly insets: readonly [number, number, number, number];
  readonly wrap: boolean;
  readonly paragraphs: ReadonlyArray<{
    readonly align: 'left' | 'center' | 'right' | 'justify';
    readonly runs: ReadonlyArray<{ readonly text: string; readonly style: string }>;
    /** Font size (pt) of an empty paragraph, so blank lines keep their height. */
    readonly emptySizePt: number;
  }>;
}

const ALIGN: Readonly<Record<string, 'left' | 'center' | 'right' | 'justify'>> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' };

/** Laid-out text of a text body: paragraphs of styled runs for an HTML overlay. */
export function layoutText(body: TextBody, paint: ShapePaint, palette: ThemePalette, zoom: number): TextLayout {
  const bp = body.bodyPr;
  const justify = bp.anchor === 'ctr' ? 'center' : bp.anchor === 'b' ? 'flex-end' : 'flex-start';
  const ins = (v: number | undefined, d: number) => (v ?? d) / EMU_PER_PX;
  return {
    justify,
    insets: [ins(bp.tIns, 45720), ins(bp.rIns, 91440), ins(bp.bIns, 45720), ins(bp.lIns, 91440)],
    wrap: bp.wrap !== 'none',
    paragraphs: body.paragraphs.map((p) => {
      const runs: Array<{ text: string; style: string }> = [];
      for (const r of p.runs) {
        if (r.kind === 'br') {
          runs.push({ text: '\n', style: '' });
          continue;
        }
        const text = r.kind === 'r' ? r.t : (r.t ?? '');
        const rp = { ...p.pPr?.defRPr, ...r.rPr };
        const color = rp.fill?.kind === 'solidFill' ? resolveDmlColor(rp.fill.color, palette).color : paint.fontColor;
        const sizePt = (rp.sz ?? 1100) / 100;
        const deco = [rp.u && rp.u !== 'none' ? 'underline' : '', rp.strike && rp.strike !== 'noStrike' ? 'line-through' : ''].filter(Boolean).join(' ');
        runs.push({
          text,
          style:
            `color:${color};font-size:${fmt(((sizePt * 96) / 72) * zoom)}px` +
            (rp.b ? ';font-weight:700' : '') +
            (rp.i ? ';font-style:italic' : '') +
            (deco ? `;text-decoration:${deco}` : '') +
            (rp.latin?.typeface && !rp.latin.typeface.startsWith('+') ? `;font-family:${JSON.stringify(rp.latin.typeface).replace(/"/g, "'")}` : ''),
        });
      }
      return { align: ALIGN[p.pPr?.algn ?? 'l'] ?? 'left', runs, emptySizePt: (p.endParaRPr?.sz ?? p.pPr?.defRPr?.sz ?? 1100) / 100 };
    }),
  };
}
