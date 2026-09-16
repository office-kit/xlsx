// Public surface for the Workbook root model + sub-features (defined names,
// shared strings, protection, views, calc properties, workbook properties,
// file metadata, smart tags, function groups).

export type {
  CellSummary,
  SheetRef,
  SheetState,
  Workbook,
  WorkbookOverview,
  WorkbookSheetOverview,
  WorkbookStats,
} from './workbook.js';
export {
  addChartsheet,
  addWorksheet,
  createWorkbook,
  describeWorkbook,
  getActiveSheet,
  getCellAtAddress,
  getCellSummary,
  getChartsheet,
  getSheet,
  getSheetState,
  getWorkbookCellsByKind,
  getWorkbookStats,
  iterWorksheets,
  listCustomXmlParts,
  moveSheet,
  removeSheet,
  renameSheet,
  setActiveSheet,
  setCellAtAddress,
  setSheetState,
  sheetNames,
} from './workbook.js';
export type { DefinedName, DefinedNameTarget } from './defined-names.js';
export {
  addDefinedName,
  getDefinedName,
  getDefinedNameTarget,
  listDefinedNames,
  makeDefinedName,
  removeDefinedName,
} from './defined-names.js';
export type { WorkbookProtection } from './protection.js';
export { makeWorkbookProtection } from './protection.js';
export type { SharedStringEntry, SharedStringsTable } from './shared-strings.js';
export {
  addSharedString,
  getSharedStringAt,
  getSharedStringIndex,
  makeSharedStrings,
  sharedStringCount,
} from './shared-strings.js';
export type {
  CustomViewShowComments,
  CustomViewShowObjects,
  CustomWorkbookView,
  WorkbookView,
  WorkbookViewVisibility,
} from './views.js';
export { makeCustomWorkbookView, makeWorkbookView } from './views.js';
export type { CalcMode, CalcProperties, RefMode } from './calc-properties.js';
export {
  makeCalcProperties,
  setCalcMode,
  setCalcOnSave,
  setFullCalcOnLoad,
  setFullPrecision,
  setIterativeCalc,
} from './calc-properties.js';
export type {
  ShowObjectsMode,
  UpdateLinksMode,
  WorkbookProperties,
} from './workbook-properties.js';
export { makeWorkbookProperties } from './workbook-properties.js';
export type { FileVersion } from './file-version.js';
export { makeFileVersion } from './file-version.js';
export type { FileSharing } from './file-sharing.js';
export { makeFileSharing } from './file-sharing.js';
export type { FileRecoveryProperties } from './file-recovery.js';
export { makeFileRecoveryProperties } from './file-recovery.js';
export type { SmartTagProperties, SmartTagShowMode, SmartTagType } from './smart-tags.js';
export { makeSmartTagProperties, makeSmartTagType } from './smart-tags.js';
export type { FunctionGroup, FunctionGroups } from './function-groups.js';
export { makeFunctionGroup, makeFunctionGroups } from './function-groups.js';
