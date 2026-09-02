// Workbook-level <bookViews> typed model. Per ECMA-376 §18.2.30 and
// openpyxl/openpyxl/workbook/views.py.
//
// `<bookViews>` carries one or more `<workbookView>` entries. The
// first entry drives Excel's default tab strip (firstSheet / activeTab)
// and window position. Most workbooks have exactly one entry.

export type WorkbookViewVisibility = 'visible' | 'hidden' | 'veryHidden';

export interface WorkbookView {
  visibility?: WorkbookViewVisibility;
  minimized?: boolean;
  showHorizontalScroll?: boolean;
  showVerticalScroll?: boolean;
  showSheetTabs?: boolean;
  /** Window x position in screen pixels — Excel restores it when re-opening. */
  xWindow?: number;
  /** Window y position. */
  yWindow?: number;
  windowWidth?: number;
  windowHeight?: number;
  /** Width of the sheet tab strip relative to the horizontal scroll bar (0..1000, default 600). */
  tabRatio?: number;
  /** Index of the leftmost visible sheet tab (0-based). */
  firstSheet?: number;
  /** Index of the currently active sheet tab (0-based). */
  activeTab?: number;
  autoFilterDateGrouping?: boolean;
  /**
   * Namespaced attributes outside SpreadsheetML, keyed in Clark notation
   * (`{ns}local`). Excel writes `xr2:uid="{…}"` here to tie the view to its
   * revision-tracking record; kept verbatim because the model has no slot for
   * it and dropping it changes what Excel reads back.
   */
  extAttrs?: Readonly<Record<string, string>>;
}

export const makeWorkbookView = (opts: WorkbookView = {}): WorkbookView => ({ ...opts });

export type CustomViewShowComments = 'commNone' | 'commIndicator' | 'commIndAndComment';
export type CustomViewShowObjects = 'all' | 'placeholders' | 'none';

export interface CustomWorkbookView {
  name: string;
  guid: string;
  /** Window width in screen pixels — required when the element is present. */
  windowWidth: number;
  windowHeight: number;
  /** Index (0-based) of the sheet active in this saved view. */
  activeSheetId: number;
  autoUpdate?: boolean;
  /** Auto-merge interval in minutes (for shared workbooks). */
  mergeInterval?: number;
  changesSavedWin?: boolean;
  onlySync?: boolean;
  personalView?: boolean;
  includePrintSettings?: boolean;
  includeHiddenRowCol?: boolean;
  maximized?: boolean;
  minimized?: boolean;
  showHorizontalScroll?: boolean;
  showVerticalScroll?: boolean;
  showSheetTabs?: boolean;
  xWindow?: number;
  yWindow?: number;
  tabRatio?: number;
  showFormulaBar?: boolean;
  showStatusbar?: boolean;
  showComments?: CustomViewShowComments;
  showObjects?: CustomViewShowObjects;
}

export const makeCustomWorkbookView = (
  opts: Pick<CustomWorkbookView, 'name' | 'guid' | 'windowWidth' | 'windowHeight' | 'activeSheetId'> &
    Partial<CustomWorkbookView>,
): CustomWorkbookView => ({ ...opts });

import type { Workbook } from './workbook';

/**
 * Get-or-create the primary `<workbookView>` entry. Most workbooks have
 * exactly one `<workbookView>`; this helper is the right place to hang
 * tab-strip / window state edits without forcing the caller to allocate
 * the array themselves.
 */
const ensurePrimaryView = (wb: Workbook): WorkbookView => {
  const existing = wb.bookViews?.[0];
  if (existing) return existing;
  const fresh = makeWorkbookView();
  wb.bookViews = [fresh];
  return fresh;
};

/** Get the index of the active sheet tab (0-based) from the primary workbookView, or 0 if unset. */
export const getActiveTab = (wb: Workbook): number => wb.bookViews?.[0]?.activeTab ?? 0;

/** Set the active sheet tab (0-based) on the primary workbookView. */
export const setActiveTab = (wb: Workbook, index: number): void => {
  ensurePrimaryView(wb).activeTab = index;
};

/** Get the index of the leftmost visible sheet tab from the primary workbookView, or 0 if unset. */
export const getFirstSheet = (wb: Workbook): number => wb.bookViews?.[0]?.firstSheet ?? 0;

/** Set the leftmost visible sheet tab on the primary workbookView. */
export const setFirstSheet = (wb: Workbook, index: number): void => {
  ensurePrimaryView(wb).firstSheet = index;
};

/** Set the tab strip width ratio (0..1000, Excel default 600). */
export const setTabRatio = (wb: Workbook, ratio: number): void => {
  ensurePrimaryView(wb).tabRatio = ratio;
};

/** Toggle the sheet tab strip visibility. */
export const setShowSheetTabs = (wb: Workbook, show: boolean): void => {
  ensurePrimaryView(wb).showSheetTabs = show;
};

/** Toggle the horizontal scroll bar in Excel's window chrome. */
export const setShowHorizontalScroll = (wb: Workbook, show: boolean): void => {
  ensurePrimaryView(wb).showHorizontalScroll = show;
};

/** Toggle the vertical scroll bar in Excel's window chrome. */
export const setShowVerticalScroll = (wb: Workbook, show: boolean): void => {
  ensurePrimaryView(wb).showVerticalScroll = show;
};

/**
 * Toggle the workbook window minimised state. Excel honours this on
 * reopen so the file restores into the minimised state it was saved
 * with.
 */
export const setWorkbookMinimized = (wb: Workbook, minimized: boolean): void => {
  ensurePrimaryView(wb).minimized = minimized;
};

/** Set the visibility of the workbook window itself ('visible' / 'hidden' / 'veryHidden'). */
export const setWorkbookVisibility = (wb: Workbook, visibility: WorkbookViewVisibility): void => {
  ensurePrimaryView(wb).visibility = visibility;
};

/**
 * Set window position + size on the primary workbookView in one call.
 * Pass `undefined` for any axis to leave it untouched.
 */
export const setWorkbookWindow = (
  wb: Workbook,
  opts: { xWindow?: number; yWindow?: number; windowWidth?: number; windowHeight?: number },
): void => {
  const v = ensurePrimaryView(wb);
  if (opts.xWindow !== undefined) v.xWindow = opts.xWindow;
  if (opts.yWindow !== undefined) v.yWindow = opts.yWindow;
  if (opts.windowWidth !== undefined) v.windowWidth = opts.windowWidth;
  if (opts.windowHeight !== undefined) v.windowHeight = opts.windowHeight;
};
