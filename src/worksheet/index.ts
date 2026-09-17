// Public surface for the Worksheet model + cell-range / dimensions / views /
// comments / hyperlinks / data-validations / conditional-formatting /
// auto-filter / tables / page-setup / protection / errors / smart-tags /
// ole-objects / sort-state / scenarios / data-consolidate / web-publish /
// phonetic / protected-ranges / properties / custom-sheet-views.

export type { ContentLimits } from './content-budget.js';
export type {
  AppendRowOptions,
  CellsByKindCounts,
  IterRowsOptions,
  Worksheet,
} from './worksheet.js';
export {
  addCellWatch,
  addConditionalFormatting,
  addDataValidation,
  addIgnoredError,
  addTable,
  appendRow,
  appendRows,
  applyToRange,
  autofitColumns,
  clearAllCells,
  clearRange,
  collapseColumnGroup,
  collapseRowGroup,
  copyRange,
  countCells,
  countCellsByKind,
  deleteCell,
  ensureCell,
  ensureCellByCoord,
  expandColumnGroup,
  expandRowGroup,
  findCells,
  getAutoFilter,
  getCell,
  getCellByCoord,
  getCellsInColumn,
  getCellsInRange,
  getCellsInRow,
  getColumnDimension,
  getDataExtent,
  getFreezePanes,
  getMaxCol,
  getMaxRow,
  getMergedCells,
  getMergedRangeAt,
  getNonEmptyCellCount,
  getPopulatedColumnIndices,
  getPopulatedRowIndices,
  getRangeValues,
  getRowDimension,
  getTable,
  groupColumns,
  groupRows,
  hideColumn,
  hideColumns,
  hideRow,
  hideRows,
  isMergedCell,
  isWorksheetEmpty,
  iterCells,
  iterRows,
  iterValues,
  listComments,
  listDataValidations,
  listHyperlinks,
  listTables,
  makeWorksheet,
  mergeCells,
  moveRange,
  removeAllComments,
  removeAllConditionalFormatting,
  removeAllDataValidations,
  removeAllHyperlinks,
  removeAllMergedRanges,
  removeAllTables,
  removeCellWatches,
  removeDataValidations,
  removeHyperlink,
  removeIgnoredErrors,
  removeTable,
  setAutoFilter,
  setCell,
  setCellByCoord,
  setColumnDimension,
  setColumnWidth,
  setColumnWidths,
  setComment,
  setComments,
  setDefaultColumnWidth,
  setDefaultRowHeight,
  setFreezePanes,
  setHyperlink,
  setHyperlinks,
  setRangeValues,
  setRowDimension,
  setRowHeight,
  setRowHeights,
  setSheetTabColor,
  setSheetViewMode,
  setSheetZoom,
  ungroupColumns,
  ungroupRows,
  unhideColumn,
  unhideColumns,
  unhideRow,
  unhideRows,
  unmergeCells,
  unmergeCellsAt,
  writeRange,
} from './worksheet.js';
export type { MultiCellRange } from './cell-range.js';
export {
  expandRangeStr,
  intersectionRange,
  isCellInRange,
  isRangeInRange,
  rangeArea,
  rangeContainsCell,
  rangeContainsRange,
  rangesOverlap,
  shiftRange,
  unionRange,
} from './cell-range.js';
export type { ColumnDimension, RowDimension } from './dimensions.js';
export { makeColumnDimension, makeRowDimension } from './dimensions.js';
export type {
  FreezeCounts,
  Pane,
  PaneState,
  PaneType,
  Selection,
  SheetView,
  SheetViewMode,
} from './views.js';
export { freezePaneRef, makeFreezePane, makeSheetView } from './views.js';
export type { LegacyComment } from './comments.js';
export { makeLegacyComment } from './comments.js';
export type {
  DataValidation,
  DataValidationErrorStyle,
  DataValidationOperator,
  DataValidationType,
  ValidationCommon,
} from './data-validations.js';
export { makeDataValidation } from './data-validations.js';
export type {
  CellIsOperator,
  Cfvo,
  CfvoType,
  ConditionalFormatting,
  ConditionalFormattingRule,
  ConditionalFormattingRuleType,
  IconSetStyle,
  TextOperator,
  TimePeriod,
} from './conditional-formatting.js';
export { makeCfRule, makeConditionalFormatting } from './conditional-formatting.js';
export type { Hyperlink } from './hyperlinks.js';
export { makeHyperlink } from './hyperlinks.js';
export type { AutoFilter, FilterColumn } from './auto-filter.js';
export { makeAutoFilter, makeFilterColumn } from './auto-filter.js';
export type { TableColumn, TableDefinition, TableStyleInfo } from './table.js';
export { addExcelTable, makeTableColumn, makeTableDefinition } from './table.js';
export type { CellWatch, IgnoredError } from './errors.js';
export { makeCellWatch, makeIgnoredError } from './errors.js';
export type { OutlineProperties, PageSetupProperties, SheetProperties } from './properties.js';
export { makeSheetProperties } from './properties.js';
export type { SheetProtection } from './protection.js';
export { makeSheetProtection } from './protection.js';
export type { ProtectedRange } from './protected-ranges.js';
export { makeProtectedRange } from './protected-ranges.js';
export type {
  SortBy,
  SortCondition,
  SortIconSet,
  SortMethod,
  SortState,
} from './sort-state.js';
export { makeSortCondition, makeSortState } from './sort-state.js';
export type {
  CellSmartTag,
  CellSmartTagProperty,
  CellSmartTags,
} from './smart-tags.js';
export {
  makeCellSmartTag,
  makeCellSmartTagProperty,
  makeCellSmartTags,
} from './smart-tags.js';
export type {
  FormControl,
  OleDvAspect,
  OleObject,
  OleUpdateMode,
} from './ole-objects.js';
export { makeFormControl, makeOleObject } from './ole-objects.js';
export type { CustomSheetView, CustomSheetViewState } from './custom-sheet-views.js';
export { makeCustomSheetView } from './custom-sheet-views.js';
export type {
  CellCommentMode,
  HeaderFooter,
  HeaderFooterSection,
  PageBreak,
  PageMargins,
  PageOrder,
  PageOrientation,
  PageSetup,
  PrintErrorMode,
  PrintOptions,
} from './page-setup.js';
export {
  buildHeaderFooterText,
  HEADER_FOOTER_CODES,
  makeHeaderFooter,
  makePageBreak,
  makePageMargins,
  makePageSetup,
  makePrintOptions,
} from './page-setup.js';
export type { WebPublishItem, WorksheetCustomProperty } from './web-publish.js';
export { makeWebPublishItem, makeWorksheetCustomProperty } from './web-publish.js';
export type {
  PhoneticAlignment,
  PhoneticType,
  WorksheetPhoneticProperties,
} from './phonetic.js';
export { makeWorksheetPhoneticProperties } from './phonetic.js';
export type {
  DataConsolidate,
  DataConsolidateFunction,
  DataReference,
} from './data-consolidate.js';
export { makeDataConsolidate } from './data-consolidate.js';
export type { Scenario, ScenarioInputCell, ScenarioList } from './scenarios.js';
export { makeScenario, makeScenarioInputCell, makeScenarioList } from './scenarios.js';
