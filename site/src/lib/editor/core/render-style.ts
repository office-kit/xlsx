// Resolves a cell's style id into everything the canvas painter needs.
//
// The stylesheet pools are append-only with de-duplication (an edit registers a
// new xf rather than mutating one), so a resolved style for a given id never
// goes stale while the workbook object lives. The cache is therefore keyed by
// id and only dropped when the workbook or theme changes.

import type { Workbook } from '@office-kit/xlsx/workbook';
import type { Border, Fill, HorizontalAlignment, Side, SideStyle, VerticalAlignment } from '@office-kit/xlsx/styles';
import { builtinFormatCode, DEFAULT_FONT } from '@office-kit/xlsx/styles';
import { paletteFromThemeXml, resolveColor, themeFontsFromXml, type ThemeFonts, type ThemePalette } from './theme.ts';

export interface StrokeStyle {
  readonly width: number;
  readonly color: string;
  readonly dash: readonly number[];
  readonly double: boolean;
}

export interface RenderStyle {
  /** CSS font shorthand at 100% zoom; the painter scales `fontPx`. */
  readonly fontFamily: string;
  readonly fontPx: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: 'none' | 'single' | 'double';
  readonly strike: boolean;
  readonly color: string;
  /** The font colour differs from the workbook default, so a table style must not recolour it. */
  readonly ownColor: boolean;
  readonly fill: string | null;
  readonly pattern: { readonly type: string; readonly fg: string; readonly bg: string } | null;
  readonly top: StrokeStyle | null;
  readonly right: StrokeStyle | null;
  readonly bottom: StrokeStyle | null;
  readonly left: StrokeStyle | null;
  readonly diagonalUp: StrokeStyle | null;
  readonly diagonalDown: StrokeStyle | null;
  readonly hAlign: HorizontalAlignment | 'general';
  readonly vAlign: VerticalAlignment;
  readonly wrap: boolean;
  readonly shrink: boolean;
  readonly indent: number;
  readonly rotation: number;
  readonly numFmt: string;
}

const FONT_FALLBACK = "'Segoe UI', 'Helvetica Neue', Arial, sans-serif";
const YU_GOTHIC = "'Yu Gothic', YuGothic, 'Hiragino Sans', 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif";
const YU_MINCHO = "'Yu Mincho', YuMincho, 'Hiragino Mincho ProN', 'MS PMincho', serif";
const P_GOTHIC = "'MS PGothic', 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif";
const P_MINCHO = "'MS PMincho', 'Hiragino Mincho ProN', YuMincho, serif";
const MS_GOTHIC = "'MS Gothic', Osaka-Mono, 'Hiragino Kaku Gothic ProN', monospace";
const MS_MINCHO = "'MS Mincho', 'Hiragino Mincho ProN', YuMincho, serif";
const MEIRYO = "Meiryo, 'Hiragino Sans', 'Hiragino Kaku Gothic ProN', sans-serif";
// Fonts named in files are often Windows/Office-only, and a browser can only
// draw fonts installed where it runs. List metric-compatible substitutes so
// text keeps its look (serif, monospace, Japanese mincho) and widths stay close
// to Excel's on other systems. Japanese fonts are listed under the English and
// the Japanese name, since files carry either.
const FONT_SUBSTITUTES: Readonly<Record<string, string>> = {
  calibri: "Calibri, Carlito, 'Segoe UI', Arial, sans-serif",
  'calibri light': "'Calibri Light', Calibri, Carlito, 'Segoe UI', Arial, sans-serif",
  'aptos narrow': "'Aptos Narrow', 'Arial Narrow', Calibri, Arial, sans-serif",
  aptos: "Aptos, Calibri, 'Segoe UI', Arial, sans-serif",
  'aptos display': "'Aptos Display', Aptos, Calibri, 'Segoe UI', Arial, sans-serif",
  arial: 'Arial, Helvetica, sans-serif',
  cambria: "Cambria, Caladea, Georgia, serif",
  consolas: "Consolas, Menlo, Monaco, 'Courier New', monospace",
  'segoe ui': "'Segoe UI', -apple-system, 'Helvetica Neue', Arial, sans-serif",
  tahoma: 'Tahoma, Verdana, sans-serif',
  'times new roman': "'Times New Roman', Times, serif",
  'yu gothic': YU_GOTHIC,
  '游ゴシック': YU_GOTHIC,
  'yu mincho': YU_MINCHO,
  '游明朝': YU_MINCHO,
  'ms pgothic': P_GOTHIC,
  'ｍｓ ｐゴシック': P_GOTHIC,
  'ms pmincho': P_MINCHO,
  'ｍｓ ｐ明朝': P_MINCHO,
  'ms gothic': MS_GOTHIC,
  'ｍｓ ゴシック': MS_GOTHIC,
  'ms mincho': MS_MINCHO,
  'ｍｓ 明朝': MS_MINCHO,
  meiryo: MEIRYO,
  'メイリオ': MEIRYO,
  'ヒラギノ角ゴシック': "'Hiragino Sans', 'Hiragino Kaku Gothic ProN', sans-serif",
};

export function cssFontFamily(name: string | undefined): string {
  if (!name) return FONT_SUBSTITUTES.calibri ?? FONT_FALLBACK;
  return FONT_SUBSTITUTES[name.toLowerCase()] ?? `'${name.replaceAll("'", '')}', ${FONT_FALLBACK}`;
}

function stroke(side: Side | undefined, palette: ThemePalette): StrokeStyle | null {
  const style: SideStyle | undefined = side?.style;
  if (!side || !style || style === 'none') return null;
  const color = resolveColor(side.color, palette, '#000000');
  switch (style) {
    case 'hair':
      return { width: 1, color, dash: [1, 1], double: false };
    case 'thin':
      return { width: 1, color, dash: [], double: false };
    case 'dotted':
      return { width: 1, color, dash: [1, 2], double: false };
    case 'dashed':
      return { width: 1, color, dash: [3, 1], double: false };
    case 'dashDot':
      return { width: 1, color, dash: [6, 2, 2, 2], double: false };
    case 'dashDotDot':
      return { width: 1, color, dash: [6, 2, 2, 2, 2, 2], double: false };
    case 'medium':
      return { width: 2, color, dash: [], double: false };
    case 'mediumDashed':
      return { width: 2, color, dash: [6, 2], double: false };
    case 'mediumDashDot':
    case 'slantDashDot':
      return { width: 2, color, dash: [8, 2, 3, 2], double: false };
    case 'mediumDashDotDot':
      return { width: 2, color, dash: [8, 2, 3, 2, 3, 2], double: false };
    case 'thick':
      return { width: 3, color, dash: [], double: false };
    case 'double':
      return { width: 3, color, dash: [], double: true };
  }
}

function solidFill(fill: Fill | undefined, palette: ThemePalette): Pick<RenderStyle, 'fill' | 'pattern'> {
  if (!fill) return { fill: null, pattern: null };
  if (fill.kind === 'gradient') {
    // A cell gradient is approximated by its first stop; exact gradients are
    // rare in grids and expensive to paint per cell.
    const first = fill.stops[0];
    return { fill: first ? resolveColor(first.color, palette, '#FFFFFF') : null, pattern: null };
  }
  const type = fill.patternType;
  if (!type || type === 'none') return { fill: null, pattern: null };
  if (type === 'solid') return { fill: resolveColor(fill.fgColor, palette, '#000000'), pattern: null };
  return {
    fill: resolveColor(fill.bgColor, palette, '#FFFFFF'),
    pattern: { type, fg: resolveColor(fill.fgColor, palette, '#000000'), bg: resolveColor(fill.bgColor, palette, '#FFFFFF') },
  };
}

const PX_PER_PT = 96 / 72;

export class StyleResolver {
  readonly #wb: Workbook;
  readonly palette: ThemePalette;
  readonly themeFonts: ThemeFonts;
  readonly #cache = new Map<number, RenderStyle>();

  constructor(wb: Workbook) {
    this.#wb = wb;
    this.palette = paletteFromThemeXml(wb.themeXml);
    this.themeFonts = themeFontsFromXml(wb.themeXml);
  }

  get(styleId: number): RenderStyle {
    let style = this.#cache.get(styleId);
    if (!style) {
      style = this.#resolve(styleId);
      this.#cache.set(styleId, style);
    }
    return style;
  }

  #resolve(styleId: number): RenderStyle {
    const ss = this.#wb.styles;
    const xf = ss.cellXfs[styleId] ?? ss.cellXfs[0];
    const font = (xf ? ss.fonts[xf.fontId] : undefined) ?? DEFAULT_FONT;
    const border: Border | undefined = xf ? ss.borders[xf.borderId] : undefined;
    const align = xf?.alignment ?? {};
    const numFmtId = xf?.numFmtId ?? 0;
    const palette = this.palette;
    const underline = font.underline === 'double' || font.underline === 'doubleAccounting' ? 'double' : font.underline && font.underline !== 'none' ? 'single' : 'none';
    return {
      fontFamily: cssFontFamily(font.name),
      fontPx: (font.size ?? 11) * PX_PER_PT,
      bold: font.bold === true,
      italic: font.italic === true,
      underline,
      strike: font.strike === true,
      color: resolveColor(font.color, palette, '#000000'),
      ownColor: resolveColor(font.color, palette, '') !== resolveColor(ss.fonts[0]?.color, palette, ''),
      ...solidFill(xf ? ss.fills[xf.fillId] : undefined, palette),
      top: stroke(border?.top, palette),
      right: stroke(border?.right, palette),
      bottom: stroke(border?.bottom, palette),
      left: stroke(border?.left, palette),
      diagonalUp: border?.diagonalUp ? stroke(border.diagonal, palette) : null,
      diagonalDown: border?.diagonalDown ? stroke(border.diagonal, palette) : null,
      hAlign: align.horizontal ?? 'general',
      vAlign: align.vertical ?? 'bottom',
      wrap: align.wrapText === true,
      shrink: align.shrinkToFit === true,
      indent: align.indent ?? 0,
      rotation: align.textRotation ?? 0,
      numFmt: builtinFormatCode(numFmtId) ?? ss.numFmts.get(numFmtId) ?? 'General',
    };
  }
}

export function canvasFont(style: RenderStyle, zoom: number): string {
  return `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${(style.fontPx * zoom).toFixed(2)}px ${style.fontFamily}`;
}
