// Public surface for the Cell value-model (CellValue + formula kinds +
// inline rich-text).

export type {
  Cell,
  CellValue,
  CellValueAsStringOptions,
  DataTableFormulaOpts,
  ExcelErrorCode,
  FormulaKind,
  FormulaValue,
  MergedCell,
} from './cell.js';
export {
  bindValue,
  cellValueAsBoolean,
  cellValueAsDate,
  cellValueAsNumber,
  cellValueAsPrimitive,
  cellValueAsString,
  getCachedFormulaValue,
  getCoordinate,
  getFormulaText,
  isDurationValue,
  isEmptyCell,
  isErrorCell,
  isErrorValue,
  isFormulaCell,
  isFormulaValue,
  isMergedCell,
  isRichTextCell,
  isRichTextValue,
  makeCell,
  makeDurationValue,
  makeErrorValue,
  makeFormula,
  setArrayFormula,
  setCellValue,
  setDataTableFormula,
  setFormula,
  setSharedFormula,
} from './cell.js';
export type {
  InlineFont,
  InlineUnderline,
  InlineVertAlign,
  RichText,
  TextRun,
} from './rich-text.js';
export { makeRichText, makeTextRun, richTextToString } from './rich-text.js';
