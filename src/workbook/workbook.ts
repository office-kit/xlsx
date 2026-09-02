// Workbook root model. / §4.2 / the Workbook is a plain mutable object the user
// composes via free functions. The Stylesheet pool is held inline so styling
// operations don't need a side channel.

import type { Chartsheet } from '../chartsheet/chartsheet';
import { makeChartsheet } from '../chartsheet/chartsheet';
import { makeAbsoluteAnchor } from '../drawing/anchor';
import { type ChartReference, makeChartDrawingItem, makeDrawing } from '../drawing/drawing';
import type { CoreProperties } from '../packaging/core';
import type { CustomProperties } from '../packaging/custom';
import type { ExtendedProperties } from '../packaging/extended';
import {
  getCellAlignment,
  getCellBorder,
  getCellFill,
  getCellFont,
  getCellNumberFormat,
  getCellProtection,
} from '../styles/cell-style';
import type { Stylesheet } from '../styles/stylesheet';
import { makeStylesheet } from '../styles/stylesheet';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { type CellValue, isFormulaValue } from '../cell/cell';
import type { Alignment } from '../styles/alignment';
import type { Border } from '../styles/borders';
import type { Fill } from '../styles/fills';
import type { Font } from '../styles/fonts';
import type { Protection } from '../styles/protection';
import { coordinateToTuple, parseSheetRange } from '../utils/coordinate';
import { multiCellRangeContainsCell, parseRange, rangeContainsCell, rangeToString } from '../worksheet/cell-range';
import type { LegacyComment } from '../worksheet/comments';
import type { Hyperlink } from '../worksheet/hyperlinks';
import type { CellsByKindCounts, Worksheet } from '../worksheet/worksheet';
import {
  classifyCellValue,
  countCellsByKind,
  getCell,
  getCellComment,
  getCellHyperlink,
  getMergedRangeAt,
  isWorksheetEmpty,
  makeWorksheet,
  setCellByCoord,
} from '../worksheet/worksheet';

export type SheetState = 'visible' | 'hidden' | 'veryHidden';

/**
 * Discriminated union over the two kinds of sheet a workbook can host. Both
 * variants share `title` (via `sheet.title`) plus the OOXML `sheetId` and
 * `state` attributes; consumers narrow on `kind` to reach the worksheet- vs
 * chartsheet-specific data.
 */
export type SheetRef =
  | { kind: 'worksheet'; sheet: Worksheet; sheetId: number; state: SheetState; rId?: string }
  | { kind: 'chartsheet'; sheet: Chartsheet; sheetId: number; state: SheetState; rId?: string };

export interface Workbook {
  sheets: SheetRef[];
  /** Index into `sheets` of the sheet shown when Excel opens the file. */
  activeSheetIndex: number;
  /** Style pool; cells reference its cellXfs by index. */
  styles: Stylesheet;
  /** Date1904 mode toggles between Excel's two epoch systems. */
  date1904: boolean;
  /** Document properties (docProps/core.xml), typically auto-filled on save. */
  properties?: CoreProperties;
  appProperties?: ExtendedProperties;
  customProperties?: CustomProperties;
  /** Author display names, shared between threaded comments. */
  authors: string[];
  /** Workbook + sheet-scope defined names (named ranges, print areas etc). */
  definedNames: import('./defined-names').DefinedName[];
  /**
   * Raw `xl/theme/theme1.xml` payload kept verbatim across read → write. The
   * theme XML is large and seldom edited by writers; we just shuttle it.
   */
  themeXml?: Uint8Array;
  /**
   * `xl/vbaProject.bin` payload (macro-enabled workbooks). Round-tripped
   * byte-identical when present; the writer also promotes the workbook Override
   * to `vnd.ms-excel.sheet.macroEnabled.main+xml`.
   */
  vbaProject?: Uint8Array;
  /** `xl/vbaProjectSignature.bin` payload, when the macros are signed. */
  vbaSignature?: Uint8Array;
  /**
   * Pass-through bytes for parts we don't model (pivot tables, ActiveX
   * controls, OLE embeddings, customUI ribbons, customXml items …). Keys are
   * archive-relative paths; values are the raw bytes the loader pulled out of
   * the zip and the writer pushes back in unchanged.
   */
  passthrough?: Map<string, Uint8Array>;
  /**
   * Override content type per pass-through path. Excel uses these in
   * `[Content_Types].xml` so manifest validation stays intact across
   * round-trips. Paths without an explicit override fall back to the archive
   * Default extension.
   */
  passthroughContentTypes?: Map<string, string>;
  /**
   * Default content type per extension for pass-through parts that had no
   * Override in the source manifest (`vml`, `jpeg`, `bin` …). The writer only
   * knows the extensions of parts it produces itself, so these are re-emitted
   * as `<Default>` entries to keep carried-over parts typed.
   */
  passthroughDefaults?: Map<string, string>;
  /**
   * Top-level `<workbook>` children that aren't `<sheets>` or `<definedNames>`
   * (e.g. `<fileVersion>`, `<workbookPr>`, `<bookViews>`, `<calcPr>`,
   * `<pivotCaches>`, `<extLst>`). Captured verbatim so re-saving keeps
   * Excel-rendering fidelity for things we don't model. Split into the two
   * halves the writer needs: anything before `<sheets>` is emitted ahead of the
   * `<sheets>` element, the rest after `<definedNames>`.
   */
  workbookXmlExtras?: {
    beforeSheets: import('../xml/tree').XmlNode[];
    afterSheets: import('../xml/tree').XmlNode[];
  };
  /**
   * The `<workbook>` root's markup-compatibility header, as Excel wrote it:
   * its namespace declarations and the `mc:Ignorable` list. The two only make
   * sense together — `mc:Ignorable="x15 xr xr6 …"` names prefixes that are
   * declared nowhere else — and the captured children have to keep the same
   * prefixes, so the whole header round-trips verbatim rather than being
   * rebuilt from the namespaces the serializer happens to walk.
   */
  workbookXmlRoot?: {
    namespaces: ReadonlyArray<{ prefix: string; ns: string }>;
    ignorable?: string;
  };
  /**
   * `<workbookProtection>` — locks structure / window / revision tracking with
   * the modern hash quad or the legacy 16-bit hash. Round-tripped verbatim;
   * password hashing helpers come later.
   */
  workbookProtection?: import('./protection').WorkbookProtection;
  /**
   * `<bookViews>` — the workbook's window/tab-strip presets. Most workbooks
   * have a single entry whose `firstSheet` / `activeTab` drive the tab the user
   * sees first. Stored as an array because Excel allows multiple views (rare).
   */
  bookViews?: import('./views').WorkbookView[];
  /**
   * `<customWorkbookViews>` — saved per-user view presets used by the
   * deprecated "Shared Workbook" feature. Each entry carries its own window
   * position, active sheet, and visibility toggles.
   */
  customWorkbookViews?: import('./views').CustomWorkbookView[];
  /** `<calcPr>` — calculation engine settings (calcMode / iterate / fullPrecision etc.). */
  calcProperties?: import('./calc-properties').CalcProperties;
  /** `<fileVersion>` — Office app/version metadata Excel records on save. */
  fileVersion?: import('./file-version').FileVersion;
  /** `<fileSharing>` — read-only-recommended toggle + write-protection password. */
  fileSharing?: import('./file-sharing').FileSharing;
  /**
   * `<oleSize ref="…">` — bounding range Excel uses when the workbook is
   * embedded as an OLE object inside another Office document.
   */
  oleSize?: string;
  /** `<fileRecoveryPr>` — autoRecover-style flags Excel writes after a recovery save. */
  fileRecoveryPr?: import('./file-recovery').FileRecoveryProperties;
  /**
   * `<pivotCaches>` — links from workbook root to xl/pivotCache parts. The
   * underlying parts survive via the passthrough archive; this typed array
   * preserves the cacheId ↔ rId mapping for consumers that want to introspect
   * the pivot links.
   */
  pivotCaches?: ReadonlyArray<{ cacheId: number; rId: string }>;
  /**
   * `<externalReferences>` — links from workbook root to xl/externalLinks
   * parts. The numeric token in cross-workbook formulas like `[1]Sheet!A1` is
   * the 1-based index into this array. Underlying parts continue via
   * passthrough archive.
   */
  externalReferences?: ReadonlyArray<{ rId: string }>;
  /** `<smartTagPr>` — Excel 2003 smart-tag persistence flags. */
  smartTagPr?: import('./smart-tags').SmartTagProperties;
  /** `<smartTagTypes>` — Excel 2003 smart-tag type registrations. */
  smartTagTypes?: ReadonlyArray<import('./smart-tags').SmartTagType>;
  /** `<functionGroups>` — built-in + user-defined XLL function groups. */
  functionGroups?: import('./function-groups').FunctionGroups;
  /**
   * `<workbookPr>` — VBA codeName, defaultThemeVersion, link-update prompt
   * mode, etc. `date1904` is mirrored here for completeness but the canonical
   * source remains `wb.date1904`.
   */
  workbookProperties?: import('./workbook-properties').WorkbookProperties;
  /**
   * Workbook-level rels that don't match a modeled type. Re-emitted with their
   * original Id so captured `<pivotCaches r:id="…"/>` etc. still resolve after
   * a round-trip.
   */
  workbookRelsExtras?: ReadonlyArray<{ id: string; type: string; target: string }>;
  /**
   * Original rIds for the modeled non-sheet workbook rels so a captured extras
   * XML referencing one of them still resolves after re-save.
   */
  workbookRelOriginalIds?: {
    sharedStrings?: string;
    styles?: string;
    theme?: string;
    vbaProject?: string;
  };
}

/** Build an empty Workbook ready to host worksheets. */
export function createWorkbook(opts?: { date1904?: boolean }): Workbook {
  return {
    sheets: [],
    activeSheetIndex: 0,
    styles: makeStylesheet(),
    date1904: opts?.date1904 ?? false,
    authors: [],
    definedNames: [],
  };
}

/**
 * Validate a sheet title against Excel's character + length rules. Returns the
 * reason string when the title is rejected; `undefined` when valid. The same
 * rules apply to worksheets and chartsheets.
 *
 * Rules:
 *  - Type must be `string`; non-empty; length ≤ 31.
 *  - May not contain any of `:`, `\`, `/`, `?`, `*`, `[`, `]`.
 *  - May not start or end with an apostrophe `'`.
 *  - May not be the literal `"History"` (case-insensitive — Excel
 * reserves that name for the change-tracking sheet).
 *
 * Uniqueness is **not** checked here; pass through `addWorksheet` /
 * `renameSheet` for the workbook-aware duplicate check.
 */
export function validateSheetTitle(title: unknown): string | undefined {
  if (typeof title !== 'string') return 'must be a string';
  if (title.length === 0) return 'must be 1..31 chars';
  if (title.length > 31) return 'must be 1..31 chars';
  if (/[:\\/?*[\]]/.test(title)) return 'must not contain : \\ / ? * [ ]';
  if (title.startsWith("'") || title.endsWith("'")) return 'must not start or end with an apostrophe';
  if (title.toLowerCase() === 'history') return '"History" is reserved by Excel';
  return undefined;
}

/** Boolean form of {@link validateSheetTitle}. */
export const isValidSheetTitle = (title: unknown): title is string => validateSheetTitle(title) === undefined;

/**
 * Pick a unique sheet title based on `base`. If `base` itself is available,
 * it's returned verbatim. Otherwise the helper appends ` (2)`, ` (3)`, … until
 * it finds a free slot. The returned title always satisfies {@link
 * validateSheetTitle} — if the base+suffix would exceed 31 chars, the base is
 * truncated to fit.
 *
 * Excel treats sheet names as case-insensitive for uniqueness, so `Data` and
 * `data` collide; the helper applies the same rule.
 *
 * Useful for "duplicate sheet" / "import" flows where you want Excel-like
 * automatic uniqueification ("Sheet1 (2)").
 */
export function pickUniqueSheetTitle(wb: Workbook, base: string): string {
  const reason = validateSheetTitle(base);
  if (reason) {
    throw new OpenXmlSchemaError(`pickUniqueSheetTitle: base "${base}" is not a valid sheet title (${reason})`);
  }
  const used = new Set<string>();
  for (const s of wb.sheets) used.add(s.sheet.title.toLowerCase());
  if (!used.has(base.toLowerCase())) return base;
  for (let n = 2; n < 1000; n++) {
    const suffix = ` (${n})`;
    const room = 31 - suffix.length;
    const truncatedBase = base.length > room ? base.slice(0, room) : base;
    const candidate = `${truncatedBase}${suffix}`;
    if (!used.has(candidate.toLowerCase())) return candidate;
  }
  throw new OpenXmlSchemaError(`pickUniqueSheetTitle: exhausted candidates for "${base}"`);
}

// Excel treats sheet names as case-insensitive for uniqueness — "Data" and
// "data" cannot co-exist in the same workbook. `ignoreIndex` lets renameSheet
// pass its own slot so a case-only rename ("Data" → "data") still succeeds.
const validateUniqueTitle = (wb: Workbook, title: string, ignoreIndex?: number): void => {
  const reason = validateSheetTitle(title);
  if (reason) {
    throw new OpenXmlSchemaError(`Worksheet title "${title}": ${reason}`);
  }
  const lower = title.toLowerCase();
  for (let i = 0; i < wb.sheets.length; i++) {
    if (i === ignoreIndex) continue;
    const s = wb.sheets[i];
    if (s && s.sheet.title.toLowerCase() === lower) {
      throw new OpenXmlSchemaError(`Worksheet title "${title}" is already in use`);
    }
  }
};

const allocateSheetId = (wb: Workbook): number => {
  // sheetId is 1-based and unique. Allocate the smallest unused integer.
  const used = new Set<number>();
  for (const s of wb.sheets) used.add(s.sheetId);
  let n = 1;
  while (used.has(n)) n++;
  return n;
};

/** Add a Worksheet to the Workbook. Returns the sheet for further population. */
export function addWorksheet(wb: Workbook, title: string, opts?: { index?: number; state?: SheetState }): Worksheet {
  validateUniqueTitle(wb, title);
  const sheet = makeWorksheet(title);
  const ref: SheetRef = {
    kind: 'worksheet',
    sheet,
    sheetId: allocateSheetId(wb),
    state: opts?.state ?? 'visible',
  };
  if (opts?.index === undefined) {
    wb.sheets.push(ref);
  } else {
    if (opts.index < 0 || opts.index > wb.sheets.length) {
      throw new OpenXmlSchemaError(`addWorksheet: index ${opts.index} out of range`);
    }
    wb.sheets.splice(opts.index, 0, ref);
  }
  return sheet;
}

/**
 * 0-based tab-strip index of the sheet (worksheet *or* chartsheet) with the
 * given title, or `-1` when not present. Useful when the caller wants to act on
 * the index for `setActiveSheet` / `swapSheets` / similar operations without
 * manually scanning `wb.sheets`.
 */
export function getSheetIndex(wb: Workbook, title: string): number {
  for (let i = 0; i < wb.sheets.length; i++) {
    const ref = wb.sheets[i];
    if (ref && ref.sheet.title === title) return i;
  }
  return -1;
}

/**
 * True iff the workbook has a sheet (worksheet *or* chartsheet) with the given
 * title. Thin shortcut over {@link getSheetIndex}.
 */
export function hasSheet(wb: Workbook, title: string): boolean {
  return getSheetIndex(wb, title) >= 0;
}

/**
 * Count sheets in the workbook, with optional kind/state filters. Mirrors the
 * filter shape of {@link getSheetTitles} but skips the array allocation when
 * the caller only needs the count.
 */
export function countSheets(
  wb: Workbook,
  opts: { kind?: 'worksheet' | 'chartsheet'; state?: SheetState } = {},
): number {
  let n = 0;
  for (const ref of wb.sheets) {
    if (opts.kind !== undefined && ref.kind !== opts.kind) continue;
    if (opts.state !== undefined && ref.state !== opts.state) continue;
    n++;
  }
  return n;
}

/**
 * Sheet titles in tab-strip order. By default returns titles for every sheet
 * (worksheets + chartsheets). Optional filters narrow to one kind
 * (`'worksheet'` / `'chartsheet'`) or one state (`'visible' | 'hidden' |
 * 'veryHidden'`).
 */
export function getSheetTitles(
  wb: Workbook,
  opts: { kind?: 'worksheet' | 'chartsheet'; state?: SheetState } = {},
): string[] {
  const out: string[] = [];
  for (const ref of wb.sheets) {
    if (opts.kind !== undefined && ref.kind !== opts.kind) continue;
    if (opts.state !== undefined && ref.state !== opts.state) continue;
    out.push(ref.sheet.title);
  }
  return out;
}

/**
 * True iff the workbook has a **worksheet** (not a chartsheet) with the given
 * title. Distinct from {@link hasSheet} (matches either kind) and {@link
 * hasChartsheet} (chartsheets only).
 */
export function hasWorksheet(wb: Workbook, title: string): boolean {
  return getSheet(wb, title) !== undefined;
}

/**
 * True iff the workbook has a **chartsheet** (not a worksheet) with the given
 * title. Distinct from {@link hasSheet}, which matches either kind. Use this
 * when the caller needs to discriminate before calling chartsheet-only
 * operations.
 */
export function hasChartsheet(wb: Workbook, title: string): boolean {
  for (const ref of wb.sheets) {
    if (ref.kind === 'chartsheet' && ref.sheet.title === title) return true;
  }
  return false;
}

/** Look up a Worksheet by title. Returns undefined for missing names or chartsheets. */
export function getSheet(wb: Workbook, title: string): Worksheet | undefined {
  for (const s of wb.sheets) {
    if (s.kind === 'worksheet' && s.sheet.title === title) return s.sheet;
  }
  return undefined;
}

/** Look up a Worksheet by index in the sheets array. Returns undefined for chartsheet slots. */
export function getSheetByIndex(wb: Workbook, idx: number): Worksheet | undefined {
  const ref = wb.sheets[idx];
  return ref?.kind === 'worksheet' ? ref.sheet : undefined;
}

/** Look up a Chartsheet by title. Returns undefined for missing names or worksheets. */
export function getChartsheet(wb: Workbook, title: string): Chartsheet | undefined {
  for (const s of wb.sheets) {
    if (s.kind === 'chartsheet' && s.sheet.title === title) return s.sheet;
  }
  return undefined;
}

/** Add a Chartsheet to the Workbook. Returns the chartsheet for further population. */
export function addChartsheet(
  wb: Workbook,
  title: string,
  opts?: { index?: number; state?: SheetState; chart?: ChartReference },
): Chartsheet {
  validateUniqueTitle(wb, title);
  const cs = makeChartsheet(title);
  // When the caller supplies a ChartReference, wrap it in a single-anchor
  // drawing so the writer emits xl/drawings/drawingN.xml + chart part.
  if (opts?.chart) {
    // The drawing wraps the chart in a single absoluteAnchor sized to a
    // standard A4-landscape page (Excel re-flows on open if needed).
    cs.drawing = makeDrawing([
      makeChartDrawingItem(makeAbsoluteAnchor({ x: 0, y: 0, cx: 9144000, cy: 6858000 }), opts.chart),
    ]);
  }
  const ref: SheetRef = {
    kind: 'chartsheet',
    sheet: cs,
    sheetId: allocateSheetId(wb),
    state: opts?.state ?? 'visible',
  };
  if (opts?.index === undefined) {
    wb.sheets.push(ref);
  } else {
    if (opts.index < 0 || opts.index > wb.sheets.length) {
      throw new OpenXmlSchemaError(`addChartsheet: index ${opts.index} out of range`);
    }
    wb.sheets.splice(opts.index, 0, ref);
  }
  return cs;
}

/** All worksheet titles, in display order. */
export function sheetNames(wb: Workbook): string[] {
  return wb.sheets.map((s) => s.sheet.title);
}

/** Remove a sheet by title. No-op if the title is not registered. */
export function removeSheet(wb: Workbook, title: string): void {
  const i = wb.sheets.findIndex((s) => s.sheet.title === title);
  if (i < 0) return;
  wb.sheets.splice(i, 1);
  // Clamp activeSheetIndex within bounds.
  if (wb.activeSheetIndex >= wb.sheets.length) {
    wb.activeSheetIndex = Math.max(0, wb.sheets.length - 1);
  }
}

/** Set the active sheet by title; throws on unknown title. */
export function setActiveSheet(wb: Workbook, title: string): void {
  const i = wb.sheets.findIndex((s) => s.sheet.title === title);
  if (i < 0) throw new OpenXmlSchemaError(`setActiveSheet: no sheet named "${title}"`);
  wb.activeSheetIndex = i;
}

/**
 * Rename a sheet from `oldTitle` to `newTitle`. Throws if no sheet matches
 * `oldTitle`, or if `newTitle` collides with an existing sheet (Excel requires
 * sheet names to be unique within a workbook).
 */
export function renameSheet(wb: Workbook, oldTitle: string, newTitle: string): void {
  const i = wb.sheets.findIndex((s) => s.sheet.title === oldTitle);
  if (i < 0) throw new OpenXmlSchemaError(`renameSheet: no sheet named "${oldTitle}"`);
  if (oldTitle === newTitle) return;
  // Excel allows case-only renames ("Data" → "data"); ignore the current slot
  // when checking uniqueness so the rename validates against everything else.
  validateUniqueTitle(wb, newTitle, i);
  const ref = wb.sheets[i];
  if (ref) ref.sheet.title = newTitle;
}

/**
 * Set the visibility state on a sheet by title. Throws on unknown title.
 * Refuses to hide the last visible sheet: an .xlsx with every sheet hidden
 * fails to open in Excel ("Excel cannot use the object linking and embedding
 * features because no sheet is visible"). Catching it here keeps the
 * workbook recoverable instead of producing a save Excel will reject.
 */
export function setSheetState(wb: Workbook, title: string, state: SheetState): void {
  const ref = wb.sheets.find((s) => s.sheet.title === title);
  if (!ref) throw new OpenXmlSchemaError(`setSheetState: no sheet named "${title}"`);
  if (ref.state === state) return;
  if (state !== 'visible' && ref.state === 'visible') {
    let otherVisible = false;
    for (const candidate of wb.sheets) {
      if (candidate !== ref && candidate.state === 'visible') {
        otherVisible = true;
        break;
      }
    }
    if (!otherVisible) {
      throw new OpenXmlSchemaError(
        `setSheetState: cannot hide "${title}" — it's the last visible sheet,` +
          ' and Excel refuses to open a workbook with every sheet hidden. Make another sheet' +
          ' visible first (or call showSheet()).',
      );
    }
  }
  ref.state = state;
}

/** Look up the current visibility state. Throws on unknown title. */
export function getSheetState(wb: Workbook, title: string): SheetState {
  const ref = wb.sheets.find((s) => s.sheet.title === title);
  if (!ref) throw new OpenXmlSchemaError(`getSheetState: no sheet named "${title}"`);
  return ref.state;
}

/**
 * Hide a sheet (`state: 'hidden'`). Equivalent to right-click → Hide in Excel —
 * the user can re-show it via the Unhide dialog.
 */
export function hideSheet(wb: Workbook, title: string): void {
  setSheetState(wb, title, 'hidden');
}

/**
 * Mark a sheet as very-hidden (`state: 'veryHidden'`). Excel won't surface it
 * in the Unhide dialog — only reachable via VBA / API.
 */
export function veryHideSheet(wb: Workbook, title: string): void {
  setSheetState(wb, title, 'veryHidden');
}

/** Make a hidden / veryHidden sheet visible. */
export function showSheet(wb: Workbook, title: string): void {
  setSheetState(wb, title, 'visible');
}

/**
 * Bulk-update visibility state for many sheets in one call. `entries` is a
 * `Record<title, state>` map; missing titles throw via the underlying
 * `setSheetState`.
 */
export function setSheetStates(wb: Workbook, entries: Record<string, SheetState>): void {
  for (const [title, state] of Object.entries(entries)) {
    setSheetState(wb, title, state);
  }
}

/**
 * Show every hidden / veryHidden worksheet. Returns the count unhidden. Useful
 * for spreadsheet-wide auditing.
 */
export function showAllSheets(wb: Workbook): number {
  let n = 0;
  for (const ref of wb.sheets) {
    if (ref.state !== 'visible') {
      ref.state = 'visible';
      n++;
    }
  }
  return n;
}

/**
 * Move a sheet to a new tab-strip position. `toIndex` is clamped to `[0,
 * sheets.length - 1]`. Adjusts `activeSheetIndex` so the same sheet stays
 * active across the move.
 */
export function moveSheet(wb: Workbook, title: string, toIndex: number): void {
  const from = wb.sheets.findIndex((s) => s.sheet.title === title);
  if (from < 0) throw new OpenXmlSchemaError(`moveSheet: no sheet named "${title}"`);
  if (!Number.isInteger(toIndex)) {
    throw new OpenXmlSchemaError(`moveSheet: toIndex must be an integer; got ${toIndex}`);
  }
  const dest = Math.max(0, Math.min(wb.sheets.length - 1, toIndex));
  if (from === dest) return;
  const wasActive = wb.activeSheetIndex === from;
  const [moved] = wb.sheets.splice(from, 1);
  if (moved) wb.sheets.splice(dest, 0, moved);
  if (wasActive) {
    wb.activeSheetIndex = dest;
  } else {
    // Re-index activeSheetIndex if the move shifted it.
    let cur = wb.activeSheetIndex;
    if (from < cur) cur -= 1;
    if (dest <= cur) cur += 1;
    wb.activeSheetIndex = Math.max(0, Math.min(wb.sheets.length - 1, cur));
  }
}

/**
 * Swap the tab-strip positions of two sheets by title. Both titles must exist;
 * throws otherwise. `activeSheetIndex` follows the moved sheet so the same
 * sheet stays active across the swap.
 */
export function swapSheets(wb: Workbook, titleA: string, titleB: string): void {
  const i = wb.sheets.findIndex((s) => s.sheet.title === titleA);
  const j = wb.sheets.findIndex((s) => s.sheet.title === titleB);
  if (i < 0) throw new OpenXmlSchemaError(`swapSheets: no sheet named "${titleA}"`);
  if (j < 0) throw new OpenXmlSchemaError(`swapSheets: no sheet named "${titleB}"`);
  if (i === j) return;
  const a = wb.sheets[i];
  const b = wb.sheets[j];
  if (!a || !b) return;
  wb.sheets[i] = b;
  wb.sheets[j] = a;
  if (wb.activeSheetIndex === i) wb.activeSheetIndex = j;
  else if (wb.activeSheetIndex === j) wb.activeSheetIndex = i;
}

/**
 * Duplicate a worksheet end-to-end and append it as `newTitle`. Mirrors Excel's
 * "Move or Copy → Create a copy" command. Cells, dimensions, styles (via shared
 * cellXf ids), comments, hyperlinks, conditional formatting, page setup, etc.
 * all carry over verbatim — only fields that must stay workbook-unique get
 * rewritten:
 *
 *  - sheet `title` → `newTitle`
 *  - sheet `sheetId` → freshly allocated
 *  - each table's `id` → max(workbook table ids) + 1
 *  - each table's `displayName` → suffixed with `opts.tableSuffix`
 * (default `"_2"`) so it doesn't collide with the original
 *
 * The new sheet is inserted at the optional `index` (default: appended).
 */
export function duplicateSheet(
  wb: Workbook,
  sourceTitle: string,
  newTitle: string,
  opts: { index?: number; state?: SheetState; tableSuffix?: string } = {},
): Worksheet {
  validateUniqueTitle(wb, newTitle);
  const sourceRef = wb.sheets.find((s) => s.kind === 'worksheet' && s.sheet.title === sourceTitle);
  if (!sourceRef || sourceRef.kind !== 'worksheet') {
    throw new OpenXmlSchemaError(`duplicateSheet: no worksheet named "${sourceTitle}"`);
  }
  const cloned = structuredClone(sourceRef.sheet);
  cloned.title = newTitle;

  // Table id + displayName must stay workbook-unique. Walk every other sheet to
  // find the next free id and renumber/rename in place.
  const suffix = opts.tableSuffix ?? '_2';
  let nextTableId = 0;
  const usedDisplayNames = new Set<string>();
  for (const s of wb.sheets) {
    if (s.kind !== 'worksheet') continue;
    for (const t of s.sheet.tables) {
      if (t.id > nextTableId) nextTableId = t.id;
      usedDisplayNames.add(t.displayName);
    }
  }
  for (const t of cloned.tables) {
    nextTableId += 1;
    t.id = nextTableId;
    let candidate = `${t.displayName}${suffix}`;
    let n = 2;
    while (usedDisplayNames.has(candidate)) {
      candidate = `${t.displayName}${suffix}${n}`;
      n += 1;
    }
    t.displayName = candidate;
    if (t.name === undefined) t.name = candidate;
    usedDisplayNames.add(candidate);
  }

  const ref: SheetRef = {
    kind: 'worksheet',
    sheet: cloned,
    sheetId: allocateSheetId(wb),
    state: opts.state ?? 'visible',
  };
  if (opts.index === undefined) {
    wb.sheets.push(ref);
  } else {
    if (opts.index < 0 || opts.index > wb.sheets.length) {
      throw new OpenXmlSchemaError(`duplicateSheet: index ${opts.index} out of range`);
    }
    wb.sheets.splice(opts.index, 0, ref);
  }
  return cloned;
}

/**
 * Aggregate counts about a workbook's content. Useful for quick QA after large
 * mutations or for surfacing a "what's in this file" banner. All counts walk
 * the typed model — they do **not** save the workbook to bytes — so the cost is
 * O(workbook content).
 */
export interface WorkbookStats {
  /** Total worksheets (excludes chartsheets). */
  worksheetCount: number;
  /** Total chartsheets. */
  chartsheetCount: number;
  /** Sum of populated cells across every worksheet. */
  cellCount: number;
  /** Sum of formula cells. */
  formulaCount: number;
  /** Sum of legacyComments across every worksheet. */
  commentCount: number;
  /** Sum of hyperlinks across every worksheet. */
  hyperlinkCount: number;
  /** Sum of mergedCells ranges. */
  mergedRangeCount: number;
  /** Sum of Excel tables. */
  tableCount: number;
  /** Workbook-level defined names. */
  definedNameCount: number;
  /** Custom-property entry count, 0 when no docProps/custom.xml. */
  customPropertyCount: number;
}

export function getWorkbookStats(wb: Workbook): WorkbookStats {
  let worksheetCount = 0;
  let chartsheetCount = 0;
  let cellCount = 0;
  let formulaCount = 0;
  let commentCount = 0;
  let hyperlinkCount = 0;
  let mergedRangeCount = 0;
  let tableCount = 0;
  for (const ref of wb.sheets) {
    if (ref.kind === 'worksheet') {
      worksheetCount++;
      const ws = ref.sheet;
      for (const rowMap of ws.rows.values()) {
        for (const cell of rowMap.values()) {
          cellCount++;
          if (isFormulaValue(cell.value)) formulaCount++;
        }
      }
      commentCount += ws.legacyComments.length;
      hyperlinkCount += ws.hyperlinks.length;
      mergedRangeCount += ws.mergedCells.length;
      tableCount += ws.tables.length;
    } else {
      chartsheetCount++;
    }
  }
  return {
    worksheetCount,
    chartsheetCount,
    cellCount,
    formulaCount,
    commentCount,
    hyperlinkCount,
    mergedRangeCount,
    tableCount,
    definedNameCount: wb.definedNames.length,
    customPropertyCount: wb.customProperties?.properties.length ?? 0,
  };
}

/**
 * Workbook-wide value-kind histogram. Sums {@link countCellsByKind} across
 * every Worksheet (chartsheets contribute no cells). Buckets have the same
 * shape as the per-worksheet result; an empty workbook returns all-zero counts.
 */
export function getWorkbookCellsByKind(wb: Workbook): CellsByKindCounts {
  const out: CellsByKindCounts = {
    null: 0,
    string: 0,
    number: 0,
    boolean: 0,
    date: 0,
    duration: 0,
    error: 0,
    'rich-text': 0,
    formula: 0,
  };
  for (const ws of iterWorksheets(wb)) {
    const partial = countCellsByKind(ws);
    out.null += partial.null;
    out.string += partial.string;
    out.number += partial.number;
    out.boolean += partial.boolean;
    out.date += partial.date;
    out.duration += partial.duration;
    out.error += partial.error;
    out['rich-text'] += partial['rich-text'];
    out.formula += partial.formula;
  }
  return out;
}

/**
 * Resolve a sheet-qualified A1 address (`'Sheet1!A1'`) to its Cell, or
 * `undefined` when the cell isn't materialised. Throws on malformed addresses,
 * missing sheets, or range inputs.
 */
export function getCellAtAddress(wb: Workbook, address: string): import('../cell/cell').Cell | undefined {
  const { sheet: sheetTitle, range } = parseSheetRange(address);
  if (range.includes(':')) {
    throw new OpenXmlSchemaError(
      `getCellAtAddress: address "${address}" refers to a range, not a single cell`,
    );
  }
  const ws = getSheet(wb, sheetTitle);
  if (!ws) {
    throw new OpenXmlSchemaError(`getCellAtAddress: sheet "${sheetTitle}" not found`);
  }
  const { col, row } = coordinateToTuple(range);
  return getCell(ws, row, col);
}

/**
 * Set a single cell by sheet-qualified A1 address. Throws on malformed
 * addresses, missing sheets, or range inputs.
 */
export function setCellAtAddress(
  wb: Workbook,
  address: string,
  value: CellValue,
): import('../cell/cell').Cell {
  const { sheet: sheetTitle, range } = parseSheetRange(address);
  if (range.includes(':')) {
    throw new OpenXmlSchemaError(
      `setCellAtAddress: address "${address}" refers to a range, not a single cell`,
    );
  }
  const ws = getSheet(wb, sheetTitle);
  if (!ws) {
    throw new OpenXmlSchemaError(`setCellAtAddress: sheet "${sheetTitle}" not found`);
  }
  return setCellByCoord(ws, range, value);
}

/**
 * True iff every Worksheet in the workbook is empty (per {@link
 * isWorksheetEmpty}). Chartsheets carry no cells so they never affect the
 * result. A workbook with zero worksheets is also empty by this definition.
 *
 * Short-circuits on the first non-empty worksheet.
 */
export function isWorkbookEmpty(wb: Workbook): boolean {
  for (const ws of iterWorksheets(wb)) {
    if (!isWorksheetEmpty(ws)) return false;
  }
  return true;
}

/**
 * Per-sheet entry inside {@link WorkbookOverview}. Holds enough metadata to
 * make a "what's in this workbook" panel useful without forcing the caller to
 * walk every worksheet themselves.
 */
export interface WorkbookSheetOverview {
  title: string;
  kind: 'worksheet' | 'chartsheet';
  state: SheetState;
  /** Populated cells in the sheet (0 for chartsheets). */
  cellCount: number;
  /** Populated formula cells (0 for chartsheets). */
  formulaCount: number;
  /** Tables registered on the sheet. */
  tableCount: number;
  /** Drawing items (charts + pictures) on the sheet. */
  drawingItemCount: number;
}

/**
 * High-level "what's in this workbook" snapshot. Combines the aggregate counts
 * from {@link getWorkbookStats} and value-kind histogram from {@link
 * getWorkbookCellsByKind} with per-sheet metadata. JSON-serialisable; suitable
 * for a UI banner / debug dump.
 */
export interface WorkbookOverview {
  worksheetCount: number;
  chartsheetCount: number;
  cellCount: number;
  formulaCount: number;
  commentCount: number;
  hyperlinkCount: number;
  mergedRangeCount: number;
  tableCount: number;
  definedNameCount: number;
  customPropertyCount: number;
  cellsByKind: CellsByKindCounts;
  sheets: WorkbookSheetOverview[];
}

export function describeWorkbook(wb: Workbook): WorkbookOverview {
  // Single-pass aggregation: a cell-by-cell walk dominates this function's
  // cost on million-cell workbooks, so we compute the workbook stats, the
  // value-kind histogram, AND each sheet's cell/formula counts in one sweep
  // instead of triple-scanning.
  let worksheetCount = 0;
  let chartsheetCount = 0;
  let cellCount = 0;
  let formulaCount = 0;
  let commentCount = 0;
  let hyperlinkCount = 0;
  let mergedRangeCount = 0;
  let tableCount = 0;
  const cellsByKind: CellsByKindCounts = {
    null: 0,
    string: 0,
    number: 0,
    boolean: 0,
    date: 0,
    duration: 0,
    error: 0,
    'rich-text': 0,
    formula: 0,
  };
  const sheets: WorkbookSheetOverview[] = wb.sheets.map((ref) => {
    if (ref.kind === 'chartsheet') {
      chartsheetCount++;
      return {
        title: ref.sheet.title,
        kind: 'chartsheet',
        state: ref.state,
        cellCount: 0,
        formulaCount: 0,
        tableCount: 0,
        drawingItemCount: ref.sheet.drawing?.items.length ?? 0,
      };
    }
    worksheetCount++;
    const ws = ref.sheet;
    let sheetCellCount = 0;
    let sheetFormulaCount = 0;
    for (const rowMap of ws.rows.values()) {
      for (const cell of rowMap.values()) {
        sheetCellCount++;
        const bucket = classifyCellValue(cell.value);
        cellsByKind[bucket]++;
        if (isFormulaValue(cell.value)) sheetFormulaCount++;
      }
    }
    cellCount += sheetCellCount;
    formulaCount += sheetFormulaCount;
    commentCount += ws.legacyComments.length;
    hyperlinkCount += ws.hyperlinks.length;
    mergedRangeCount += ws.mergedCells.length;
    tableCount += ws.tables.length;
    return {
      title: ws.title,
      kind: 'worksheet',
      state: ref.state,
      cellCount: sheetCellCount,
      formulaCount: sheetFormulaCount,
      tableCount: ws.tables.length,
      drawingItemCount: ws.drawing?.items.length ?? 0,
    };
  });
  return {
    worksheetCount,
    chartsheetCount,
    cellCount,
    formulaCount,
    commentCount,
    hyperlinkCount,
    mergedRangeCount,
    tableCount,
    definedNameCount: wb.definedNames.length,
    customPropertyCount: wb.customProperties?.properties.length ?? 0,
    cellsByKind,
    sheets,
  };
}

/**
 * Debug-friendly snapshot of everything resolved for a single cell: its value,
 * the full style chain (font / fill / border / alignment / protection /
 * numberFormat), the applied hyperlink + comment, the merged range it sits
 * inside (if any), and the names of any tables / the count of CF / DV blocks
 * that target it.
 *
 * Designed for `console.log`-style introspection — JSON-serialisable and stable
 * in shape regardless of which axes are populated.
 *
 * Throws when `sheetTitle` doesn't resolve. When `ref` is a valid A1 coordinate
 * but no cell exists there, `exists` is `false` and the style chain reflects
 * the workbook defaults.
 */
export interface CellSummary {
  ref: string;
  sheet: string;
  exists: boolean;
  value: CellValue | undefined;
  styleId: number;
  font: Font;
  fill: Fill;
  border: Border;
  alignment: Alignment;
  protection: Protection;
  numberFormat: string;
  hyperlink: Hyperlink | undefined;
  comment: LegacyComment | undefined;
  mergedRange: string | undefined;
  inTables: string[];
  inDataValidations: number;
  inConditionalFormatting: number;
}

export function getCellSummary(wb: Workbook, sheetTitle: string, ref: string): CellSummary {
  const ws = getSheet(wb, sheetTitle);
  if (!ws) throw new OpenXmlSchemaError(`getCellSummary: sheet "${sheetTitle}" not found`);
  const { col, row } = coordinateToTuple(ref);
  const cell = getCell(ws, row, col);
  // Synthesize a placeholder cell so getCell* helpers can resolve defaults even
  // for unmaterialised coordinates.
  const probe = cell ?? { row, col, value: null, styleId: 0 };
  const merged = getMergedRangeAt(ws, row, col);
  const inTables: string[] = [];
  for (const t of ws.tables) {
    if (rangeContainsCell(parseRange(t.ref), row, col)) inTables.push(t.displayName);
  }
  let inDv = 0;
  for (const dv of ws.dataValidations) {
    if (multiCellRangeContainsCell(dv.sqref, row, col)) inDv++;
  }
  let inCf = 0;
  for (const cf of ws.conditionalFormatting) {
    if (multiCellRangeContainsCell(cf.sqref, row, col)) inCf++;
  }
  return {
    ref,
    sheet: sheetTitle,
    exists: cell !== undefined,
    value: cell?.value,
    styleId: probe.styleId,
    font: getCellFont(wb, probe),
    fill: getCellFill(wb, probe),
    border: getCellBorder(wb, probe),
    alignment: getCellAlignment(wb, probe),
    protection: getCellProtection(wb, probe),
    numberFormat: getCellNumberFormat(wb, probe),
    hyperlink: cell ? getCellHyperlink(ws, cell) : undefined,
    comment: cell ? getCellComment(ws, cell) : undefined,
    mergedRange: merged ? rangeToString(merged) : undefined,
    inTables,
    inDataValidations: inDv,
    inConditionalFormatting: inCf,
  };
}

/**
 * Iterate over every Worksheet in the workbook (skips chartsheets). Yields each
 * worksheet in tab-strip order.
 */
export function* iterWorksheets(wb: Workbook): IterableIterator<Worksheet> {
  for (const ref of wb.sheets) {
    if (ref.kind === 'worksheet') yield ref.sheet;
  }
}

/**
 * Iterate only over Worksheets whose tab-strip state is `'visible'`. Hidden /
 * veryHidden sheets are skipped. Useful for reports that should ignore
 * back-office sheets the author has hidden.
 */
export function* iterVisibleWorksheets(wb: Workbook): IterableIterator<Worksheet> {
  for (const ref of wb.sheets) {
    if (ref.kind === 'worksheet' && ref.state === 'visible') yield ref.sheet;
  }
}

/**
 * Iterate Worksheets matching the supplied state. Pass `'hidden'` to skim
 * back-office sheets, `'veryHidden'` to find sheets only accessible via VBA,
 * etc.
 */
export function* iterWorksheetsByState(
  wb: Workbook,
  state: SheetState,
): IterableIterator<Worksheet> {
  for (const ref of wb.sheets) {
    if (ref.kind === 'worksheet' && ref.state === state) yield ref.sheet;
  }
}

/**
 * Iterate every cell across every worksheet in the workbook. Yields `{ sheet,
 * cell }` pairs in tab-strip order, then row-then-column within each sheet.
 * Useful for workbook-wide audits / find-and-replace passes.
 */
export function* iterAllCells(
  wb: Workbook,
): IterableIterator<{ sheet: Worksheet; cell: import('../cell/cell').Cell }> {
  for (const sheet of iterWorksheets(wb)) {
    const rowKeys = [...sheet.rows.keys()].sort((a, b) => a - b);
    for (const r of rowKeys) {
      const rowMap = sheet.rows.get(r);
      if (!rowMap) continue;
      const cols = [...rowMap.keys()].sort((a, b) => a - b);
      for (const c of cols) {
        const cell = rowMap.get(c);
        if (cell !== undefined) yield { sheet, cell };
      }
    }
  }
}

/**
 * Collect every merged range across every worksheet. Each entry carries the
 * merge bounds plus a back-reference to the owning sheet, in tab-strip order.
 * Equivalent to walking `iterWorksheets` and concatenating each sheet's
 * `mergedCells`.
 */
export function getAllMergedRanges(
  wb: Workbook,
): ReadonlyArray<{ sheet: Worksheet; range: import('../worksheet/cell-range').CellRange }> {
  const out: Array<{ sheet: Worksheet; range: import('../worksheet/cell-range').CellRange }> = [];
  for (const sheet of iterWorksheets(wb)) {
    for (const range of sheet.mergedCells) out.push({ sheet, range });
  }
  return out;
}

/**
 * Collect every hyperlink across every worksheet. Each entry pairs the
 * hyperlink with a back-reference to the owning sheet, in tab-strip order.
 */
export function getAllHyperlinks(
  wb: Workbook,
): ReadonlyArray<{ sheet: Worksheet; hyperlink: import('../worksheet/hyperlinks').Hyperlink }> {
  const out: Array<{ sheet: Worksheet; hyperlink: import('../worksheet/hyperlinks').Hyperlink }> = [];
  for (const sheet of iterWorksheets(wb)) {
    for (const h of sheet.hyperlinks) out.push({ sheet, hyperlink: h });
  }
  return out;
}

/**
 * Collect every legacy comment across every worksheet. Each entry pairs the
 * comment with a back-reference to the owning sheet, in tab-strip order.
 */
export function getAllComments(
  wb: Workbook,
): ReadonlyArray<{ sheet: Worksheet; comment: import('../worksheet/comments').LegacyComment }> {
  const out: Array<{ sheet: Worksheet; comment: import('../worksheet/comments').LegacyComment }> = [];
  for (const sheet of iterWorksheets(wb)) {
    for (const c of sheet.legacyComments) out.push({ sheet, comment: c });
  }
  return out;
}

/**
 * Collect every Excel table across every worksheet. Each entry pairs the
 * TableDefinition with a back-reference to the owning sheet, in tab-strip
 * order.
 */
export function getAllTables(
  wb: Workbook,
): ReadonlyArray<{ sheet: Worksheet; table: import('../worksheet/table').TableDefinition }> {
  const out: Array<{ sheet: Worksheet; table: import('../worksheet/table').TableDefinition }> = [];
  for (const sheet of iterWorksheets(wb)) {
    for (const t of sheet.tables) out.push({ sheet, table: t });
  }
  return out;
}

/**
 * Locate an Excel table by `displayName` across the whole workbook. Excel
 * enforces uniqueness at the workbook level, so the first match wins. Returns
 * the owning sheet + the table itself, or `undefined` when nothing matches.
 */
export function findTable(
  wb: Workbook,
  displayName: string,
): { sheet: Worksheet; table: import('../worksheet/table').TableDefinition } | undefined {
  for (const sheet of iterWorksheets(wb)) {
    for (const t of sheet.tables) {
      if (t.displayName === displayName) return { sheet, table: t };
    }
  }
  return undefined;
}

/**
 * First cell across the workbook satisfying `predicate`. Walks every worksheet
 * in tab-strip order, then row-then-column within each sheet (same order as
 * {@link iterAllCells}). Returns `{ sheet, cell }` for the match, or
 * `undefined` when nothing matches.
 */
export function findCellInWorkbook(
  wb: Workbook,
  predicate: (cell: import('../cell/cell').Cell, sheet: Worksheet) => boolean,
): { sheet: Worksheet; cell: import('../cell/cell').Cell } | undefined {
  for (const { sheet, cell } of iterAllCells(wb)) {
    if (predicate(cell, sheet)) return { sheet, cell };
  }
  return undefined;
}

/**
 * Every cell across the workbook satisfying `predicate`. Same iteration order
 * as {@link iterAllCells}. Returns an array of `{ sheet, cell }` matches.
 */
export function findCellsInWorkbook(
  wb: Workbook,
  predicate: (cell: import('../cell/cell').Cell, sheet: Worksheet) => boolean,
): ReadonlyArray<{ sheet: Worksheet; cell: import('../cell/cell').Cell }> {
  const out: Array<{ sheet: Worksheet; cell: import('../cell/cell').Cell }> = [];
  for (const { sheet, cell } of iterAllCells(wb)) {
    if (predicate(cell, sheet)) out.push({ sheet, cell });
  }
  return out;
}

/**
 * Workbook-wide find-and-replace. Same matching rule as `replaceCellValues` but
 * walks every worksheet via {@link iterAllCells}. `search` is either an
 * exact-string match (string-valued cells only) or a predicate `(value, cell,
 * sheet) → boolean`. `replacement` is the new `CellValue`. Returns the count of
 * cells changed across all sheets.
 */
export function replaceCellValuesInWorkbook(
  wb: Workbook,
  search:
    | string
    | ((value: import('../cell/cell').CellValue, cell: import('../cell/cell').Cell, sheet: Worksheet) => boolean),
  replacement: import('../cell/cell').CellValue,
): number {
  let n = 0;
  const matchFn =
    typeof search === 'string'
      ? (v: import('../cell/cell').CellValue) => typeof v === 'string' && v === search
      : (v: import('../cell/cell').CellValue, c: import('../cell/cell').Cell, s: Worksheet) => search(v, c, s);
  for (const { sheet, cell } of iterAllCells(wb)) {
    if (matchFn(cell.value, cell, sheet)) {
      cell.value = replacement;
      n++;
    }
  }
  return n;
}

/**
 * Collect every data-validation block across every worksheet. Each entry pairs
 * the validation with a back-reference to the owning sheet, in tab-strip order.
 */
export function getAllDataValidations(
  wb: Workbook,
): ReadonlyArray<{ sheet: Worksheet; validation: import('../worksheet/data-validations').DataValidation }> {
  const out: Array<{
    sheet: Worksheet;
    validation: import('../worksheet/data-validations').DataValidation;
  }> = [];
  for (const sheet of iterWorksheets(wb)) {
    for (const v of sheet.dataValidations) out.push({ sheet, validation: v });
  }
  return out;
}

/**
 * Collect every image (picture) DrawingItem across every worksheet, each paired
 * with its owning sheet in tab-strip order.
 */
export function getAllImages(
  wb: Workbook,
): ReadonlyArray<{ sheet: Worksheet; item: import('../drawing/drawing').DrawingItem }> {
  const out: Array<{ sheet: Worksheet; item: import('../drawing/drawing').DrawingItem }> = [];
  for (const sheet of iterWorksheets(wb)) {
    if (!sheet.drawing) continue;
    for (const item of sheet.drawing.items) {
      if (item.content.kind === 'picture') out.push({ sheet, item });
    }
  }
  return out;
}

/**
 * Collect every chart DrawingItem across every worksheet, each paired with its
 * owning sheet in tab-strip order.
 */
export function getAllCharts(
  wb: Workbook,
): ReadonlyArray<{ sheet: Worksheet; item: import('../drawing/drawing').DrawingItem }> {
  const out: Array<{ sheet: Worksheet; item: import('../drawing/drawing').DrawingItem }> = [];
  for (const sheet of iterWorksheets(wb)) {
    if (!sheet.drawing) continue;
    for (const item of sheet.drawing.items) {
      if (item.content.kind === 'chart') out.push({ sheet, item });
    }
  }
  return out;
}

/**
 * Collect every conditional-formatting block across every worksheet. Each entry
 * pairs the CF block with a back-reference to the owning sheet, in tab-strip
 * order.
 */
export function getAllConditionalFormatting(
  wb: Workbook,
): ReadonlyArray<{
  sheet: Worksheet;
  formatting: import('../worksheet/conditional-formatting').ConditionalFormatting;
}> {
  const out: Array<{
    sheet: Worksheet;
    formatting: import('../worksheet/conditional-formatting').ConditionalFormatting;
  }> = [];
  for (const sheet of iterWorksheets(wb)) {
    for (const cf of sheet.conditionalFormatting) out.push({ sheet, formatting: cf });
  }
  return out;
}

/**
 * Iterate over every Chartsheet in the workbook. Yields in tab-strip order,
 * skipping regular worksheets.
 */
export function* iterChartsheets(wb: Workbook): IterableIterator<Chartsheet> {
  for (const ref of wb.sheets) {
    if (ref.kind === 'chartsheet') yield ref.sheet;
  }
}

/** Convenience: array of every Worksheet in tab-strip order. */
export function listWorksheets(wb: Workbook): Worksheet[] {
  return [...iterWorksheets(wb)];
}

/** Convenience: array of every Chartsheet in tab-strip order. */
export function listChartsheets(wb: Workbook): Chartsheet[] {
  return [...iterChartsheets(wb)];
}

/** Currently active sheet (worksheet only), or undefined if the active slot is empty or a chartsheet. */
export function getActiveSheet(wb: Workbook): Worksheet | undefined {
  const ref = wb.sheets[wb.activeSheetIndex];
  return ref?.kind === 'worksheet' ? ref.sheet : undefined;
}

/**
 * Title of whichever sheet (worksheet *or* chartsheet) is currently marked
 * active via `wb.activeSheetIndex`. Returns `undefined` for an empty workbook
 * or an out-of-range index.
 *
 * Distinct from {@link getActiveSheet} (which only returns worksheets and
 * yields `undefined` when the active slot is a chartsheet) — this matches
 * `wb.activeSheetIndex` regardless of kind.
 */
export function getActiveSheetTitle(wb: Workbook): string | undefined {
  return wb.sheets[wb.activeSheetIndex]?.sheet.title;
}

/**
 * True iff `title` matches the workbook's currently active sheet (any kind).
 * Empty workbook returns `false` (no active sheet).
 */
export function isActiveSheet(wb: Workbook, title: string): boolean {
  return getActiveSheetTitle(wb) === title;
}

/** Read-only view onto the customXml/* pass-through parts. */
export function listCustomXmlParts(wb: Workbook): Array<{ path: string; content: Uint8Array }> {
  if (!wb.passthrough) return [];
  const out: Array<{ path: string; content: Uint8Array }> = [];
  for (const [path, content] of wb.passthrough) {
    if (path.startsWith('customXml/')) out.push({ path, content });
  }
  return out;
}

/**
 * JSON.stringify replacer that drops the Stylesheet's internal dedup Maps. Use
 * as `JSON.stringify(workbook, jsonReplacer)` when the workbook needs to
 * round-trip through plain JSON (tests, debug dumps). The dedup maps are
 * reconstructed lazily on first add.
 */
export function jsonReplacer(_key: string, value: unknown): unknown {
  if (value instanceof Map) {
    return { __map__: [...value.entries()] };
  }
  return value;
}

/** Companion reviver for {@link jsonReplacer}. */
export function jsonReviver(_key: string, value: unknown): unknown {
  if (typeof value === 'object' && value !== null && Array.isArray((value as { __map__?: unknown[] }).__map__)) {
    return new Map((value as { __map__: Array<[unknown, unknown]> }).__map__);
  }
  return value;
}
