// Public surface for the style value-objects + cell ↔ stylesheet bridge. —
// Color / Font / Fill / Border / Alignment / Protection / NumberFormat are
// plain objects with `make*` factories, and Stylesheet pools dedup equal values
// via stable keys.

export type { Alignment, HorizontalAlignment, VerticalAlignment } from './alignment.js';
export { alignmentToCss, makeAlignment } from './alignment.js';
export type { Border, Side, SideStyle } from './borders.js';
export { borderToCss, makeBorder, makeSide } from './borders.js';
export type { CellStyleSpec } from './cell-style.js';
export {
  alignCellHorizontal,
  alignCellVertical,
  applyBuiltinStyle,
  applyNamedStyle,
  cellStyleToCss,
  centerCell,
  clearCellBackground,
  clearCellStyle,
  clearRangeStyle,
  cloneCellStyle,
  copyCellStyle,
  formatAsHeader,
  indentCell,
  patchCellFont,
  registerCellStyle,
  rotateCellText,
  setBold,
  setCellBackgroundColor,
  setFontColor,
  setFontName,
  setFontSize,
  setItalic,
  setStrikethrough,
  setUnderline,
  wrapCellText,
  getCellAlignment,
  getCellBorder,
  getCellFill,
  getCellFont,
  getCellNumberFormat,
  getCellProtection,
  setCellAlignment,
  setCellAsCurrency,
  setCellAsDate,
  setCellAsNumber,
  setCellAsPercent,
  setCellBorder,
  setCellBorderAll,
  setCellFill,
  setCellFont,
  setCellNumberFormat,
  setCellProtection,
  setCellStyle,
  setRangeBackgroundColor,
  setRangeBorderBox,
  setRangeFont,
  setRangeNumberFormat,
  setRangeAlignment,
  setRangeProtection,
  setRangeStyle,
  setRangeWrapText,
} from './cell-style.js';
export type { DifferentialStyle } from './differential.js';
export { addDxf, makeDifferentialStyle } from './differential.js';
export type { NamedStyle, StylesheetNamedStyle } from './named-styles.js';
export { addNamedStyle, BUILTIN_NAMED_STYLES, ensureBuiltinStyle } from './named-styles.js';
export type { Color } from './colors.js';
export {
  adjustLightness,
  adjustSaturation,
  colorToHex,
  contrastRatio,
  darken,
  hexToHsl,
  hslToHex,
  lighten,
  luminance,
  makeColor,
  mixColors,
  normaliseRgb,
  pickReadableTextColor,
  resolveIndexedColor,
  rgbColor,
  rotateHue,
} from './colors.js';
export type { Fill, GradientFill, GradientFillType, GradientStop, PatternFill, PatternType } from './fills.js';
export { fillToCss, makeFill, makeGradientFill, makeGradientStop, makePatternFill } from './fills.js';
export type { Font, FontScheme, UnderlineStyle, VertAlign } from './fonts.js';
export { DEFAULT_FONT, fontToCss, makeFont } from './fonts.js';
export type { Protection } from './protection.js';
export { makeProtection } from './protection.js';
export type { CellXf, Stylesheet } from './stylesheet.js';
export {
  addBorder,
  addCellStyleXf,
  addCellXf,
  addFill,
  addFont,
  addNumFmt,
  defaultCellXf,
  listBorders,
  listCellStyleXfs,
  listCellXfs,
  listFills,
  listFonts,
  makeStylesheet,
} from './stylesheet.js';
export type { NumberFormat } from './numbers.js';
export {
  BUILTIN_FORMATS,
  BUILTIN_FORMATS_MAX_SIZE,
  builtinFormatCode,
  builtinFormatId,
  classifyDateFormat,
  FORMAT_DATE_DATETIME,
  FORMAT_DATE_TIMEDELTA,
  FORMAT_DATE_YYYYMMDD2,
  FORMAT_GENERAL,
  FORMAT_NUMBER,
  FORMAT_NUMBER_00,
  FORMAT_PERCENTAGE,
  FORMAT_PERCENTAGE_00,
  FORMAT_TEXT,
  isBuiltinFormat,
  isDateFormat,
  isTimedeltaFormat,
  makeNumberFormat,
} from './numbers.js';
