// Utility surfaces — coordinate / datetime / units / inference / escape /
// css / exception types.

export type {
  AbsoluteRefOptions,
  CellCoordinate,
  CellCoordinateNumeric,
  CellRangeBoundaries,
  RangeRef,
} from './coordinate.js';
export {
  boundariesToRangeString,
  columnIndexFromLetter,
  columnLetterFromIndex,
  coordinateFromString,
  coordinateToTuple,
  formatSheetQualifiedRef,
  isValidCellRef,
  isValidColumnLetter,
  isValidColumnNumber,
  isValidRangeRef,
  isValidRowNumber,
  MAX_COL,
  MAX_ROW,
  parseSheetRange,
  rangeBoundaries,
  tupleToCoordinate,
} from './coordinate.js';
export type { ExcelEpoch } from './datetime.js';
export {
  dateToExcel,
  durationToExcel,
  excelToDate,
  excelToDuration,
  fromIso8601,
  MAC_EPOCH_MS,
  toIso8601,
  WINDOWS_EPOCH_MS,
} from './datetime.js';
export { cssRecordToInlineStyle } from './css.js';
export { escapeCellString, unescapeCellString } from './escape.js';
export type { OpenXmlErrorOptions } from './exceptions.js';
export {
  OpenXmlContentLimitError,
  OpenXmlDecompressionBombError,
  OpenXmlError,
  OpenXmlInvalidWorkbookError,
  OpenXmlIoError,
  OpenXmlNotImplementedError,
  OpenXmlSchemaError,
} from './exceptions.js';
export type { CellDataType } from './inference.js';
export { ERROR_CODES, inferCellType } from './inference.js';
export {
  cmFromEmu,
  EMU_PER_CM,
  EMU_PER_INCH,
  EMU_PER_PIXEL,
  EMU_PER_POINT,
  emuFromCm,
  emuFromInch,
  emuFromPoint,
  emuFromPx,
  inchFromEmu,
  pixelToPoint,
  pointFromEmu,
  pointToPixel,
  pxFromEmu,
} from './units.js';
