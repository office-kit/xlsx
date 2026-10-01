// Cell font value object. Mirrors openpyxl/openpyxl/styles/fonts.py.
//
// Font is the most varied of the styles slots — most fields are nested
// elements with a single `val` attribute (`<sz val="11"/>`), the boolean
// toggles (`<b/>`, `<i/>`, ...) are presence-only marker tags, and
// `<color>` is a fully nested object element. The schema layer carries
// each pattern as its own ElementDef kind.

import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { type Color, colorToHex, makeColor } from './colors.js';

/** Underline styles per ECMA-376 ST_UnderlineValues, including explicit no underline. */
export type UnderlineStyle = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | 'none';
export const UNDERLINE_STYLES: ReadonlyArray<UnderlineStyle> = Object.freeze([
  'single',
  'double',
  'singleAccounting',
  'doubleAccounting',
  'none',
]);

export type VertAlign = 'baseline' | 'superscript' | 'subscript';
export const VERT_ALIGNS: ReadonlyArray<VertAlign> = Object.freeze(['baseline', 'superscript', 'subscript']);

export type FontScheme = 'major' | 'minor';
export const FONT_SCHEMES: ReadonlyArray<FontScheme> = Object.freeze(['major', 'minor']);

const UNDERLINE_SET: ReadonlySet<string> = new Set(UNDERLINE_STYLES);
const VERT_SET: ReadonlySet<string> = new Set(VERT_ALIGNS);
const SCHEME_SET: ReadonlySet<string> = new Set(FONT_SCHEMES);

export interface Font {
  readonly name?: string;
  readonly charset?: number;
  /** Font family code (0..14). */
  readonly family?: number;
  /** Point size; openpyxl name = `sz`. */
  readonly size?: number;
  readonly color?: Color;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly strike?: boolean;
  readonly outline?: boolean;
  readonly shadow?: boolean;
  readonly condense?: boolean;
  readonly extend?: boolean;
  /** `none` explicitly disables underline; omission leaves it unspecified (e.g. in a DXF). */
  readonly underline?: UnderlineStyle;
  readonly vertAlign?: VertAlign;
  readonly scheme?: FontScheme;
}

/**
 * Input shape for building a Font. Wider than `Partial<Font>` under
 * `exactOptionalPropertyTypes`: a field may be present and `undefined`, which
 * {@link makeFont} drops the same way it drops an absent one. That is what lets
 * a caller merging two fonts express "this field goes away".
 */
export type FontPatch = { [K in keyof Font]?: Font[K] | undefined };

export function makeFont(opts: FontPatch = {}): Font {
  const out: { -readonly [K in keyof Font]: Font[K] } = {};
  if (opts.name !== undefined) {
    if (typeof opts.name !== 'string') {
      throw new OpenXmlSchemaError(`Font name must be a string; got ${typeof opts.name}`);
    }
    out.name = opts.name;
  }
  if (opts.charset !== undefined) {
    if (!Number.isInteger(opts.charset)) {
      throw new OpenXmlSchemaError(`Font charset must be an integer; got ${opts.charset}`);
    }
    out.charset = opts.charset;
  }
  if (opts.family !== undefined) {
    if (!Number.isInteger(opts.family) || opts.family < 0 || opts.family > 14) {
      throw new OpenXmlSchemaError(`Font family must be 0..14; got ${opts.family}`);
    }
    out.family = opts.family;
  }
  if (opts.size !== undefined) {
    if (!Number.isFinite(opts.size) || opts.size <= 0) {
      throw new OpenXmlSchemaError(`Font size must be positive; got ${opts.size}`);
    }
    out.size = opts.size;
  }
  if (opts.color !== undefined) {
    out.color = Object.isFrozen(opts.color) ? opts.color : makeColor(opts.color);
  }
  if (opts.bold !== undefined) out.bold = opts.bold;
  if (opts.italic !== undefined) out.italic = opts.italic;
  if (opts.strike !== undefined) out.strike = opts.strike;
  if (opts.outline !== undefined) out.outline = opts.outline;
  if (opts.shadow !== undefined) out.shadow = opts.shadow;
  if (opts.condense !== undefined) out.condense = opts.condense;
  if (opts.extend !== undefined) out.extend = opts.extend;
  if (opts.underline !== undefined) {
    if (!UNDERLINE_SET.has(opts.underline)) {
      throw new OpenXmlSchemaError(
        `Font underline must be one of [${UNDERLINE_STYLES.join(', ')}]; got "${opts.underline}"`,
      );
    }
    out.underline = opts.underline;
  }
  if (opts.vertAlign !== undefined) {
    if (!VERT_SET.has(opts.vertAlign)) {
      throw new OpenXmlSchemaError(
        `Font vertAlign must be one of [${VERT_ALIGNS.join(', ')}]; got "${opts.vertAlign}"`,
      );
    }
    out.vertAlign = opts.vertAlign;
  }
  if (opts.scheme !== undefined) {
    if (!SCHEME_SET.has(opts.scheme)) {
      throw new OpenXmlSchemaError(`Font scheme must be one of [${FONT_SCHEMES.join(', ')}]; got "${opts.scheme}"`);
    }
    out.scheme = opts.scheme;
  }
  return Object.freeze(out);
}

/** Excel's default cell font: Calibri 11, minor scheme, theme=1 colour. */
export const DEFAULT_FONT: Font = makeFont({
  name: 'Calibri',
  size: 11,
  family: 2,
  scheme: 'minor',
  color: makeColor({ theme: 1 }),
});

/**
 * Translate a {@link Font} to a CSS-property record suitable for HTML
 * rendering / preview. Boolean toggles map to weight/style/decoration,
 * `size` becomes `font-size: <n>pt`, and an explicit rgb/indexed
 * `color` is rendered as `#RRGGBB` (alpha dropped). theme/auto colours
 * are skipped — callers without a theme can't resolve them.
 *
 * `vertAlign='superscript'|'subscript'` lowers `font-size` to 0.83em
 * (W3C convention) and sets `vertical-align`.
 *
 * Returns `{}` for an empty Font so callers can spread it without
 * overwriting upstream defaults.
 */
export function fontToCss(font: Font | undefined): Record<string, string> {
  const css: Record<string, string> = {};
  if (!font) return css;
  if (font.name !== undefined) {
    // CSS expects quoted family names that contain spaces; quote unconditionally
    // to keep the generator simple.
    css['font-family'] = `'${font.name.replace(/'/g, "\\'")}'`;
  }
  if (font.size !== undefined) css['font-size'] = `${font.size}pt`;
  if (font.bold) css['font-weight'] = 'bold';
  if (font.italic) css['font-style'] = 'italic';
  const decorations: string[] = [];
  if (font.underline !== undefined && font.underline !== 'none') decorations.push('underline');
  if (font.strike) decorations.push('line-through');
  if (decorations.length > 0) css['text-decoration'] = decorations.join(' ');
  else if (font.underline === 'none') css['text-decoration'] = 'none';
  if (font.color !== undefined) {
    const argb = colorToHex(font.color);
    if (argb !== undefined) {
      // Drop alpha (CSS hex with alpha is poorly-supported in Excel-style preview).
      css['color'] = `#${argb.slice(2)}`;
    }
  }
  if (font.vertAlign === 'superscript' || font.vertAlign === 'subscript') {
    css['vertical-align'] = font.vertAlign;
    // Match W3C-recommended scaling for sup/sub when font-size is not explicit.
    if (css['font-size'] === undefined) css['font-size'] = '0.83em';
  }
  return css;
}
