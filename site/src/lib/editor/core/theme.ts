// Theme colour resolution. Cells reference theme slots (`<color theme="4"
// tint="0.4"/>`) that only mean something against the workbook's
// `xl/theme/theme1.xml`, which the library carries as raw bytes.

import type { Color } from '@office-kit/xlsx/styles';
import { colorToHex } from '@office-kit/xlsx/styles';

/** Twelve RGB hex strings (no '#') in SpreadsheetML theme-index order. */
export type ThemePalette = readonly string[];

// SpreadsheetML theme indices swap the first two pairs relative to the
// DrawingML scheme order: 0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2.
const SLOT_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** Office 2013–2022 theme: what Excel assumes when a workbook ships no theme part. */
export const OFFICE_PALETTE: ThemePalette = [
  'FFFFFF',
  '000000',
  'E7E6E6',
  '44546A',
  '4472C4',
  'ED7D31',
  'A5A5A5',
  'FFC000',
  '5B9BD5',
  '70AD47',
  '0563C1',
  '954F72',
];

export function paletteFromThemeXml(xml: Uint8Array | undefined): ThemePalette {
  if (!xml) return OFFICE_PALETTE;
  const text = new TextDecoder().decode(xml);
  const scheme = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(text)?.[0];
  if (!scheme) return OFFICE_PALETTE;
  return SLOT_ORDER.map((slot, i) => {
    const body = new RegExp(`<a:${slot}>([\\s\\S]*?)</a:${slot}>`).exec(scheme)?.[1] ?? '';
    const rgb = /srgbClr val="([0-9A-Fa-f]{6})"/.exec(body)?.[1] ?? /lastClr="([0-9A-Fa-f]{6})"/.exec(body)?.[1];
    return (rgb ?? OFFICE_PALETTE[i] ?? '000000').toUpperCase();
  });
}

export function rgbToHsl(hex: string): [number, number, number] {
  const r = Number.parseInt(hex.slice(0, 2), 16) / 255;
  const g = Number.parseInt(hex.slice(2, 4), 16) / 255;
  const b = Number.parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

export function hslToRgb(h: number, s: number, l: number): string {
  const hue = (p: number, q: number, t0: number): number => {
    let t = t0;
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r = l;
  let g = l;
  let b = l;
  if (s !== 0) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue(p, q, h + 1 / 3);
    g = hue(p, q, h);
    b = hue(p, q, h - 1 / 3);
  }
  const hex = (v: number) =>
    Math.round(v * 255)
      .toString(16)
      .padStart(2, '0');
  return `${hex(r)}${hex(g)}${hex(b)}`.toUpperCase();
}

/** Excel's tint: darken toward black for negative values, lighten toward white for positive. */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, s, l] = rgbToHsl(hex);
  const lum = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return hslToRgb(h, s, lum);
}

/**
 * CSS colour for a style colour, or `fallback` when the colour is absent or
 * "automatic". Theme and tint are resolved against `palette`.
 */
export function resolveColor(color: Color | undefined, palette: ThemePalette, fallback: string): string {
  if (!color || color.auto) return fallback;
  let rgb: string | undefined;
  if (color.theme !== undefined) rgb = palette[color.theme];
  else {
    const argb = colorToHex(color);
    // Indexed 64/65 are the system foreground/background: "automatic".
    if (argb === undefined) return fallback;
    rgb = argb.slice(-6);
  }
  if (rgb === undefined) return fallback;
  return `#${color.tint ? applyTint(rgb, color.tint) : rgb}`;
}
