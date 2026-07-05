// Public surface for drawings (charts/images embedded in a worksheet),
// anchors, image bytes, and DML shape properties.
//
// `@office-kit/xlsx/drawing` doubles as the home for the DML primitives that
// `ShapeProperties` / `TextBody` are built from — colours, fills, lines,
// effects, and the rich-text body — so callers building chart styling can
// import everything from a single subpath.

export type {
  ChartReference,
  Drawing,
  DrawingItem,
  PictureReference,
} from './drawing';
export {
  addChartAt,
  addImageAt,
  listChartsOnSheet,
  listImagesOnSheet,
  makeChartDrawingItem,
  makeDrawing,
  makePictureDrawingItem,
  removeAllCharts,
  removeAllDrawingItems,
  removeAllImages,
} from './drawing';
export type { XlsxImage, XlsxImageFormat } from './image';
export { loadImage } from './image';
export type {
  AnchorMarker,
  DrawingAnchor,
  Point2D,
  PositiveSize2D,
} from './anchor';
export { makeOneCellAnchor } from './anchor';
export type { BlackWhiteMode, ShapeProperties, Transform2D } from './dml/shape-properties';
export { makeShapeProperties } from './dml/shape-properties';

// ---- DML colours -----------------------------------------------------------
export type { ColorMod, DmlColor, DmlColorWithMods, SchemeColorName } from './dml/colors';
export {
  makeColor,
  makeSchemeColor,
  makeSrgbColor,
  SCHEME_COLOR_NAMES,
  VALUED_COLOR_MOD_KINDS,
  VALUELESS_COLOR_MOD_KINDS,
} from './dml/colors';

// ---- DML fills -------------------------------------------------------------
export type {
  Blip,
  BlipEffect,
  Fill,
  GradientLineDir,
  GradientStop,
  RelativeRect,
  TileFill,
  TileFlip,
} from './dml/fill';
export {
  makeGradientFill,
  makeNoFill,
  makePatternFill,
  makeSolidFill,
  PRESET_PATTERN_NAMES,
} from './dml/fill';

// ---- DML text body (used by chart axis / title / legend `txPr`) ------------
export type {
  AutoFit,
  BulletProperties,
  FontAlign,
  HyperlinkInfo,
  ParagraphAlign,
  ParagraphProperties,
  RunProperties,
  TabStop,
  TextAnchor,
  TextBody,
  TextBodyProperties,
  TextCap,
  TextFont,
  TextHorzOverflow,
  TextListStyle,
  TextOverflow,
  TextParagraph,
  TextRun,
  TextSpacing,
  TextStrike,
  TextUnderline,
  TextVertical,
  TextWrap,
} from './dml/text';
export {
  makeBreak,
  makeParagraph,
  makeRun,
  makeRunProperties,
  makeSimpleTextBody,
  makeTextBody,
} from './dml/text';
