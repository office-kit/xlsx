// Worksheet XML reader.
//
// **Stage 1**: DOM-based reader for `<sheetData>/<row>/<c>` covering the common
// cell-value shapes — number / shared-string / boolean / error / inline string
// / formula. SAX iterparse + dimension / sheetView / mergeCells / hyperlinks
// etc. land in later iterations of the loop. The function signature is stable
// so the SAX swap won't break downstream callers.

import {
  type Cell,
  type ExcelErrorCode,
  type FormulaKind,
  setArrayFormula,
  setDataTableFormula,
  setFormula,
  setSharedFormula,
} from '../cell/cell';
import type { Drawing } from '../drawing/drawing';
import { translateFormula } from '../formula/translate';
import type { Relationships } from '../packaging/relationships';
import { findById } from '../packaging/relationships';
import { coordinateToTuple, tupleToCoordinate } from '../utils/coordinate';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { ERROR_CODES } from '../utils/inference';
import { MARKUP_COMPAT_NS, REL_NS, SHEET_MAIN_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { serializeXml } from '../xml/serializer';
import { findChild, findChildren, type XmlNode } from '../xml/tree';
import type { AutoFilter, FilterColumn } from './auto-filter';
import { parseMultiCellRange, parseRange } from './cell-range';
import type { LegacyComment } from './comments';
import type {
  ConditionalFormatting,
  ConditionalFormattingRule,
  ConditionalFormattingRuleType,
  TimePeriod,
} from './conditional-formatting';
import { makeCfRule, makeConditionalFormatting } from './conditional-formatting';
import type {
  DataValidation,
  DataValidationErrorStyle,
  DataValidationOperator,
  DataValidationType,
} from './data-validations';
import { makeDataValidation } from './data-validations';
import type { ColumnDimension, RowDimension } from './dimensions';
import type { IgnoredError } from './errors';
import type {
  CellCommentMode,
  HeaderFooter,
  PageBreak,
  PageMargins,
  PageOrder,
  PageOrientation,
  PageSetup,
  PrintErrorMode,
  PrintOptions,
} from './page-setup';
import type {
  PhoneticAlignment,
  PhoneticType,
  WorksheetPhoneticProperties,
} from './phonetic';
import type {
  DataConsolidate,
  DataConsolidateFunction,
  DataReference,
} from './data-consolidate';
import type { Scenario, ScenarioInputCell, ScenarioList } from './scenarios';
import type { OutlineProperties, PageSetupProperties, SheetProperties } from './properties';
import type { SheetProtection } from './protection';
import type { ProtectedRange } from './protected-ranges';
import type { SortBy, SortCondition, SortIconSet, SortMethod, SortState } from './sort-state';
import type { FormControl, OleDvAspect, OleObject, OleUpdateMode } from './ole-objects';
import type { CustomSheetView, CustomSheetViewState } from './custom-sheet-views';
import type { WebPublishItem, WorksheetCustomProperty } from './web-publish';
import { makeColor } from '../styles/colors';
import { makeColumnDimension, makeRowDimension } from './dimensions';
import type { Hyperlink } from './hyperlinks';
import type { TableDefinition } from './table';
import type { Pane, PaneState, PaneType, Selection, SheetView, SheetViewMode } from './views';
import { makeSheetView } from './views';
import { makeWorksheet, setCell, type Worksheet } from './worksheet';

const WORKSHEET_TAG = `{${SHEET_MAIN_NS}}worksheet`;
const SHEETDATA_TAG = `{${SHEET_MAIN_NS}}sheetData`;
const ROW_TAG = `{${SHEET_MAIN_NS}}row`;
const C_TAG = `{${SHEET_MAIN_NS}}c`;
const V_TAG = `{${SHEET_MAIN_NS}}v`;
const F_TAG = `{${SHEET_MAIN_NS}}f`;
const IS_TAG = `{${SHEET_MAIN_NS}}is`;
const T_TAG = `{${SHEET_MAIN_NS}}t`;
const MERGE_CELLS_TAG = `{${SHEET_MAIN_NS}}mergeCells`;
const MERGE_CELL_TAG = `{${SHEET_MAIN_NS}}mergeCell`;
const SHEET_VIEWS_TAG = `{${SHEET_MAIN_NS}}sheetViews`;
const SHEET_VIEW_TAG = `{${SHEET_MAIN_NS}}sheetView`;
const PANE_TAG = `{${SHEET_MAIN_NS}}pane`;
const SELECTION_TAG = `{${SHEET_MAIN_NS}}selection`;
const COLS_TAG = `{${SHEET_MAIN_NS}}cols`;
const COL_TAG = `{${SHEET_MAIN_NS}}col`;
const SHEET_FORMAT_PR_TAG = `{${SHEET_MAIN_NS}}sheetFormatPr`;
const HYPERLINKS_TAG = `{${SHEET_MAIN_NS}}hyperlinks`;
const HYPERLINK_TAG = `{${SHEET_MAIN_NS}}hyperlink`;
const DATA_VALIDATIONS_TAG = `{${SHEET_MAIN_NS}}dataValidations`;
const DATA_VALIDATION_TAG = `{${SHEET_MAIN_NS}}dataValidation`;
const FORMULA1_TAG = `{${SHEET_MAIN_NS}}formula1`;
const FORMULA2_TAG = `{${SHEET_MAIN_NS}}formula2`;
const AUTOFILTER_TAG = `{${SHEET_MAIN_NS}}autoFilter`;
const FILTER_COLUMN_TAG = `{${SHEET_MAIN_NS}}filterColumn`;
const FILTERS_TAG = `{${SHEET_MAIN_NS}}filters`;
const FILTER_TAG = `{${SHEET_MAIN_NS}}filter`;
const TABLE_PARTS_TAG = `{${SHEET_MAIN_NS}}tableParts`;
const TABLE_PART_TAG = `{${SHEET_MAIN_NS}}tablePart`;
const CONDITIONAL_FORMATTING_TAG = `{${SHEET_MAIN_NS}}conditionalFormatting`;
const CF_RULE_TAG = `{${SHEET_MAIN_NS}}cfRule`;
const FORMULA_TAG = `{${SHEET_MAIN_NS}}formula`;
const DRAWING_TAG = `{${SHEET_MAIN_NS}}drawing`;
const CELL_WATCHES_TAG = `{${SHEET_MAIN_NS}}cellWatches`;
const CELL_WATCH_TAG = `{${SHEET_MAIN_NS}}cellWatch`;
const IGNORED_ERRORS_TAG = `{${SHEET_MAIN_NS}}ignoredErrors`;
const IGNORED_ERROR_TAG = `{${SHEET_MAIN_NS}}ignoredError`;
const SHEET_PR_TAG = `{${SHEET_MAIN_NS}}sheetPr`;
const TAB_COLOR_TAG = `{${SHEET_MAIN_NS}}tabColor`;
const OUTLINE_PR_TAG = `{${SHEET_MAIN_NS}}outlinePr`;
const PAGE_SETUP_PR_TAG = `{${SHEET_MAIN_NS}}pageSetUpPr`;
const SHEET_PROTECTION_TAG = `{${SHEET_MAIN_NS}}sheetProtection`;
const PROTECTED_RANGES_TAG = `{${SHEET_MAIN_NS}}protectedRanges`;
const PROTECTED_RANGE_TAG = `{${SHEET_MAIN_NS}}protectedRange`;
const SORT_STATE_TAG = `{${SHEET_MAIN_NS}}sortState`;
const SORT_CONDITION_TAG = `{${SHEET_MAIN_NS}}sortCondition`;
const PICTURE_TAG = `{${SHEET_MAIN_NS}}picture`;
const LEGACY_DRAWING_TAG = `{${SHEET_MAIN_NS}}legacyDrawing`;
const LEGACY_DRAWING_HF_TAG = `{${SHEET_MAIN_NS}}legacyDrawingHF`;
const MC_ALTERNATE_CONTENT_TAG = `{${MARKUP_COMPAT_NS}}AlternateContent`;
const MC_CHOICE_TAG = `{${MARKUP_COMPAT_NS}}Choice`;
const CUSTOM_SHEET_VIEWS_TAG = `{${SHEET_MAIN_NS}}customSheetViews`;
const CUSTOM_SHEET_VIEW_TAG = `{${SHEET_MAIN_NS}}customSheetView`;
const OLE_OBJECTS_TAG = `{${SHEET_MAIN_NS}}oleObjects`;
const OLE_OBJECT_TAG = `{${SHEET_MAIN_NS}}oleObject`;
const OBJECT_PR_TAG = `{${SHEET_MAIN_NS}}objectPr`;
const CONTROLS_TAG = `{${SHEET_MAIN_NS}}controls`;
const CONTROL_TAG = `{${SHEET_MAIN_NS}}control`;
const CONTROL_PR_TAG = `{${SHEET_MAIN_NS}}controlPr`;
const SMART_TAGS_TAG = `{${SHEET_MAIN_NS}}smartTags`;
const CELL_SMART_TAGS_TAG = `{${SHEET_MAIN_NS}}cellSmartTags`;
const CELL_SMART_TAG_TAG = `{${SHEET_MAIN_NS}}cellSmartTag`;
const CELL_SMART_TAG_PR_TAG = `{${SHEET_MAIN_NS}}cellSmartTagPr`;
const PRINT_OPTIONS_TAG = `{${SHEET_MAIN_NS}}printOptions`;
const PAGE_MARGINS_TAG = `{${SHEET_MAIN_NS}}pageMargins`;
const PAGE_SETUP_TAG = `{${SHEET_MAIN_NS}}pageSetup`;
const HEADER_FOOTER_TAG = `{${SHEET_MAIN_NS}}headerFooter`;
const ROW_BREAKS_TAG = `{${SHEET_MAIN_NS}}rowBreaks`;
const COL_BREAKS_TAG = `{${SHEET_MAIN_NS}}colBreaks`;
const BRK_TAG = `{${SHEET_MAIN_NS}}brk`;
const CUSTOM_PROPERTIES_TAG = `{${SHEET_MAIN_NS}}customProperties`;
const CUSTOM_PROPERTY_TAG = `{${SHEET_MAIN_NS}}customProperty`;
const WEB_PUBLISH_ITEMS_TAG = `{${SHEET_MAIN_NS}}webPublishItems`;
const WEB_PUBLISH_ITEM_TAG = `{${SHEET_MAIN_NS}}webPublishItem`;
const PHONETIC_PR_TAG = `{${SHEET_MAIN_NS}}phoneticPr`;
const DATA_CONSOLIDATE_TAG = `{${SHEET_MAIN_NS}}dataConsolidate`;
const DATA_REFS_TAG = `{${SHEET_MAIN_NS}}dataRefs`;
const DATA_REF_TAG = `{${SHEET_MAIN_NS}}dataRef`;
const SCENARIOS_TAG = `{${SHEET_MAIN_NS}}scenarios`;
const SCENARIO_TAG = `{${SHEET_MAIN_NS}}scenario`;
const INPUT_CELLS_TAG = `{${SHEET_MAIN_NS}}inputCells`;
const ODD_HEADER_TAG = `{${SHEET_MAIN_NS}}oddHeader`;
const ODD_FOOTER_TAG = `{${SHEET_MAIN_NS}}oddFooter`;
const EVEN_HEADER_TAG = `{${SHEET_MAIN_NS}}evenHeader`;
const EVEN_FOOTER_TAG = `{${SHEET_MAIN_NS}}evenFooter`;
const FIRST_HEADER_TAG = `{${SHEET_MAIN_NS}}firstHeader`;
const FIRST_FOOTER_TAG = `{${SHEET_MAIN_NS}}firstFooter`;

/** Inputs the worksheet reader needs from the surrounding workbook context. */
export interface WorksheetReadContext {
  /** Resolved shared-strings table. Pass `[]` when no sst is present. */
  sharedStrings: ReadonlyArray<string>;
  /** This worksheet's `_rels/sheetN.xml.rels`. Used to resolve external hyperlink targets and table parts. */
  rels?: Relationships;
  /** Resolves a worksheet-rels rId pointing at xl/tables/tableN.xml into a parsed TableDefinition. */
  loadTable?: (relId: string) => TableDefinition | undefined;
  /**
   * Loader for comments parts (rels Type=…/comments). Called once per matching
   * rel; the reader appends every returned LegacyComment onto
   * `ws.legacyComments`.
   */
  loadComments?: (relId: string) => ReadonlyArray<LegacyComment> | undefined;
  /**
   * Loader for the worksheet's drawing part. Called once when the worksheet
   * inline carries `<drawing r:id="...">`.
   */
  loadDrawing?: (relId: string) => Drawing | undefined;
}

/** Per-worksheet state for shared-formula expansion. */
interface SharedFormulaCache {
  origin: string;
  formula: string;
}

/**
 * Parse a `xl/worksheets/sheetN.xml` payload into a fully-populated Worksheet.
 * The returned worksheet's `title` matches the `title` argument; the XML
 * doesn't carry the sheet name (it lives in `workbook.xml`).
 */
export function parseWorksheetXml(bytes: Uint8Array | string, title: string, ctx: WorksheetReadContext): Worksheet {
  const root = parseXml(bytes);
  if (root.name !== WORKSHEET_TAG) {
    throw new OpenXmlSchemaError(`parseWorksheetXml: root is "${root.name}", expected worksheet`);
  }
  const ws = makeWorksheet(title);

  // <sheetPr codeName="…" filterMode="0" ...> <tabColor rgb="FF0070C0"/>
  // <outlinePr summaryBelow="1" .../> <pageSetUpPr fitToPage="1"/> </sheetPr>
  const sheetPrEl = findChild(root, SHEET_PR_TAG);
  if (sheetPrEl) {
    const props = parseSheetProperties(sheetPrEl);
    if (props) ws.sheetProperties = props;
  }

  // <sheetFormatPr> defaults — recorded so dimension-less sheets still reflect
  // any non-default workbook-wide row height / column width.
  const sheetFormatEl = findChild(root, SHEET_FORMAT_PR_TAG);
  if (sheetFormatEl) {
    const defaultColumnWidth = parseFloatAttr(sheetFormatEl.attrs['defaultColWidth']);
    if (defaultColumnWidth !== undefined) ws.defaultColumnWidth = defaultColumnWidth;
    const defaultRowHeight = parseFloatAttr(sheetFormatEl.attrs['defaultRowHeight']);
    if (defaultRowHeight !== undefined) ws.defaultRowHeight = defaultRowHeight;
    const outlineLevelRow = parseIntegerAttr(sheetFormatEl.attrs['outlineLevelRow']);
    if (outlineLevelRow !== undefined) ws.outlineLevelRow = outlineLevelRow;
    const outlineLevelCol = parseIntegerAttr(sheetFormatEl.attrs['outlineLevelCol']);
    if (outlineLevelCol !== undefined) ws.outlineLevelCol = outlineLevelCol;
    const customHeight = parseBoolXmlAttr(sheetFormatEl.attrs['customHeight']);
    if (customHeight !== undefined) ws.customHeight = customHeight;
    const zeroHeight = parseBoolXmlAttr(sheetFormatEl.attrs['zeroHeight']);
    if (zeroHeight !== undefined) ws.zeroHeight = zeroHeight;
    const thickTop = parseBoolXmlAttr(sheetFormatEl.attrs['thickTop']);
    if (thickTop !== undefined) ws.thickTop = thickTop;
    const thickBottom = parseBoolXmlAttr(sheetFormatEl.attrs['thickBottom']);
    if (thickBottom !== undefined) ws.thickBottom = thickBottom;
    const baseColWidth = parseIntegerAttr(sheetFormatEl.attrs['baseColWidth']);
    if (baseColWidth !== undefined) ws.baseColWidth = baseColWidth;
  }

  // <cols> column dimensions — preserve runs verbatim (one entry per <col>).
  const colsEl = findChild(root, COLS_TAG);
  if (colsEl) {
    for (const c of findChildren(colsEl, COL_TAG)) {
      const dim = parseColumnDimension(c);
      ws.columnDimensions.set(dim.min, dim);
    }
  }

  const sheetData = findChild(root, SHEETDATA_TAG);
  if (sheetData) {
    const sharedFormulas = new Map<number, SharedFormulaCache>();
    for (const rowNode of findChildren(sheetData, ROW_TAG)) {
      const rowIdx = parseRowIndex(rowNode);
      maybeRecordRowDimension(ws, rowNode, rowIdx);
      let nextCol = 1;
      for (const cNode of findChildren(rowNode, C_TAG)) {
        const coord = parseCellCoord(cNode, rowIdx, nextCol);
        readCell(ws, cNode, coord, ctx, sharedFormulas);
        nextCol = coord.col + 1;
      }
    }
  }

  // <sheetProtection> sits between sheetData and mergeCells per ECMA-376
  // §18.3.1.85. Parse all 16 boolean lock flags + optional password hash
  // fields.
  const protectionEl = findChild(root, SHEET_PROTECTION_TAG);
  if (protectionEl) {
    ws.sheetProtection = parseSheetProtection(protectionEl);
  }

  // <sortState ref=… columnSort=… caseSensitive=… sortMethod=…> <sortCondition
  // ref=… descending=… sortBy=… .../> </sortState>
  const ssEl = findChild(root, SORT_STATE_TAG);
  if (ssEl) {
    const ss = parseSortState(ssEl);
    if (ss) ws.sortState = ss;
  }

  // <protectedRanges><protectedRange sqref=… name=… [hash
  // quad]/></protectedRanges>
  const prsEl = findChild(root, PROTECTED_RANGES_TAG);
  if (prsEl) {
    for (const pr of findChildren(prsEl, PROTECTED_RANGE_TAG)) {
      const sqref = pr.attrs['sqref'];
      const name = pr.attrs['name'];
      if (!sqref || !name) continue;
      const out: ProtectedRange = { sqref: parseMultiCellRange(sqref), name };
      if (pr.attrs['password'] !== undefined) out.password = pr.attrs['password'];
      if (pr.attrs['securityDescriptor'] !== undefined)
        out.securityDescriptor = pr.attrs['securityDescriptor'];
      if (pr.attrs['algorithmName'] !== undefined) out.algorithmName = pr.attrs['algorithmName'];
      if (pr.attrs['hashValue'] !== undefined) out.hashValue = pr.attrs['hashValue'];
      if (pr.attrs['saltValue'] !== undefined) out.saltValue = pr.attrs['saltValue'];
      if (pr.attrs['spinCount'] !== undefined) {
        const n = Number.parseInt(pr.attrs['spinCount'], 10);
        if (Number.isInteger(n)) out.spinCount = n;
      }
      ws.protectedRanges.push(out);
    }
  }

  // <mergeCells> sits as a sibling of <sheetData>; pull each <mergeCell
  // ref="…"/> straight onto the worksheet's mergedCells list. We don't go
  // through mergeCells() here because its overlap check is for *new* merges;
  // the source xlsx is assumed valid.
  const mergeCellsEl = findChild(root, MERGE_CELLS_TAG);
  if (mergeCellsEl) {
    for (const m of findChildren(mergeCellsEl, MERGE_CELL_TAG)) {
      const ref = m.attrs['ref'];
      if (!ref) throw new OpenXmlSchemaError('worksheet: <mergeCell> missing @ref');
      ws.mergedCells.push(parseRange(ref));
    }
  }

  // <sheetViews> sits early in the worksheet but our load order doesn't matter
  // — reading it after sheetData keeps the loops separable.
  const sheetViewsEl = findChild(root, SHEET_VIEWS_TAG);
  if (sheetViewsEl) {
    for (const v of findChildren(sheetViewsEl, SHEET_VIEW_TAG)) {
      ws.views.push(parseSheetView(v));
    }
  }

  // <hyperlinks> — relations to external URLs come from the worksheet rels.
  const hyperlinksEl = findChild(root, HYPERLINKS_TAG);
  if (hyperlinksEl) {
    for (const h of findChildren(hyperlinksEl, HYPERLINK_TAG)) {
      ws.hyperlinks.push(parseHyperlink(h, ctx.rels));
    }
  }

  // <dataValidations> — list / range / formula constraints, sqref-scoped.
  const dvWrap = findChild(root, DATA_VALIDATIONS_TAG);
  if (dvWrap) {
    for (const d of findChildren(dvWrap, DATA_VALIDATION_TAG)) {
      ws.dataValidations.push(parseDataValidation(d));
    }
  }

  // <autoFilter ref="..."> with optional <filterColumn> children.
  const autoFilterEl = findChild(root, AUTOFILTER_TAG);
  if (autoFilterEl) {
    const filter = parseAutoFilter(autoFilterEl);
    if (filter) ws.autoFilter = filter;
  }

  // <conditionalFormatting sqref="…"><cfRule .../></conditionalFormatting> —
  // multiple per sheet, each with its own sqref + rules array.
  for (const cfEl of findChildren(root, CONDITIONAL_FORMATTING_TAG)) {
    const cf = parseConditionalFormatting(cfEl);
    if (cf) ws.conditionalFormatting.push(cf);
  }

  // <tableParts><tablePart r:id="rIdN"/></tableParts> — resolved through the
  // ctx.loadTable callback (load.ts threads the archive in).
  const tablePartsEl = findChild(root, TABLE_PARTS_TAG);
  if (tablePartsEl && ctx.loadTable) {
    for (const tp of findChildren(tablePartsEl, TABLE_PART_TAG)) {
      const rId = tp.attrs[`{${REL_NS}}id`];
      if (!rId) continue;
      const table = ctx.loadTable(rId);
      if (table) {
        table.rId = rId;
        ws.tables.push(table);
      }
    }
  }

  // Comments don't have a sheet-inline anchor — they're discovered via the rels
  // file (Type ending in /comments). Walk those and let load.ts read each
  // commentsN.xml part.
  if (ctx.rels && ctx.loadComments) {
    for (const rel of ctx.rels.rels) {
      if (rel.type === `${REL_NS}/comments`) {
        const list = ctx.loadComments(rel.id);
        if (list) ws.legacyComments.push(...list);
      }
    }
  }

  // <picture r:id="rIdN"/> — sheet background image.
  const pictureEl = findChild(root, PICTURE_TAG);
  if (pictureEl) {
    const rId = pictureEl.attrs[`{${REL_NS}}id`];
    if (rId) ws.backgroundPictureRId = rId;
  }

  // <legacyDrawing r:id="rIdN"/> — comment / form-control VML. Whether the
  // rId is kept or superseded by a regenerated comment VML is decided by the
  // caller, which can see the VML part itself.
  const ldEl = findChild(root, LEGACY_DRAWING_TAG);
  if (ldEl) {
    const rId = ldEl.attrs[`{${REL_NS}}id`];
    if (rId) ws.legacyDrawingRId = rId;
  }

  // <legacyDrawingHF r:id="rIdN"/> — header/footer background VML.
  const lhfEl = findChild(root, LEGACY_DRAWING_HF_TAG);
  if (lhfEl) {
    const rId = lhfEl.attrs[`{${REL_NS}}id`];
    if (rId) ws.legacyDrawingHFRId = rId;
  }

  // <customSheetViews><customSheetView guid="…" scale="…"
  // …>...</customSheetView></customSheetViews>
  const csvWrap = findChild(root, CUSTOM_SHEET_VIEWS_TAG);
  if (csvWrap) {
    for (const v of findChildren(csvWrap, CUSTOM_SHEET_VIEW_TAG)) {
      const view = parseCustomSheetView(v);
      if (view) ws.customSheetViews.push(view);
    }
  }

  // <oleObjects><oleObject ...><objectPr.../></oleObject></oleObjects>
  for (const ooWrap of findChildrenThroughMc(root, OLE_OBJECTS_TAG)) {
    for (const oo of findChildrenThroughMc(ooWrap, OLE_OBJECT_TAG)) {
      const parsed = parseOleObject(oo);
      if (parsed) ws.oleObjects.push(parsed);
    }
  }

  // <controls><control ...><controlPr.../></control></controls>
  for (const cWrap of findChildrenThroughMc(root, CONTROLS_TAG)) {
    for (const c of findChildrenThroughMc(cWrap, CONTROL_TAG)) {
      const parsed = parseFormControl(c);
      if (parsed) ws.controls.push(parsed);
    }
  }

  // <smartTags><cellSmartTags r="A1"><cellSmartTag
  // type=…><cellSmartTagPr/>…</cellSmartTag></cellSmartTags></smartTags>
  const stEl = findChild(root, SMART_TAGS_TAG);
  if (stEl) {
    for (const cstNode of findChildren(stEl, CELL_SMART_TAGS_TAG)) {
      const ref = cstNode.attrs['r'];
      if (!ref) continue;
      const tags: import('./smart-tags').CellSmartTag[] = [];
      for (const tagNode of findChildren(cstNode, CELL_SMART_TAG_TAG)) {
        const typeRaw = tagNode.attrs['type'];
        if (typeRaw === undefined) continue;
        const type = Number.parseInt(typeRaw, 10);
        if (!Number.isInteger(type)) continue;
        const tag: import('./smart-tags').CellSmartTag = { type, properties: [] };
        const deleted = parseBoolXmlAttr(tagNode.attrs['deleted']);
        if (deleted !== undefined) tag.deleted = deleted;
        const xmlBased = parseBoolXmlAttr(tagNode.attrs['xmlBased']);
        if (xmlBased !== undefined) tag.xmlBased = xmlBased;
        for (const prop of findChildren(tagNode, CELL_SMART_TAG_PR_TAG)) {
          const key = prop.attrs['key'];
          const val = prop.attrs['val'];
          if (key !== undefined && val !== undefined) tag.properties.push({ key, val });
        }
        tags.push(tag);
      }
      ws.smartTags.push({ ref, tags });
    }
  }

  // <drawing r:id="rIdN"/> — at most one per sheet. Resolve via loadDrawing.
  const drawingEl = findChild(root, DRAWING_TAG);
  if (drawingEl && ctx.loadDrawing) {
    const rId = drawingEl.attrs[`{${REL_NS}}id`];
    if (rId) {
      const d = ctx.loadDrawing(rId);
      if (d) ws.drawing = d;
    }
  }

  // <printOptions> / <pageMargins> / <pageSetup> / <headerFooter> — page-setup
  // typed model (B6). Sit between <hyperlinks> and the legacy drawing block per
  // ECMA-376.
  const poEl = findChild(root, PRINT_OPTIONS_TAG);
  if (poEl) {
    const po = parsePrintOptions(poEl);
    if (po) ws.printOptions = po;
  }
  const pmEl = findChild(root, PAGE_MARGINS_TAG);
  if (pmEl) {
    const pm = parsePageMargins(pmEl);
    if (pm) ws.pageMargins = pm;
  }
  const psEl = findChild(root, PAGE_SETUP_TAG);
  if (psEl) {
    const ps = parsePageSetup(psEl);
    if (ps) ws.pageSetup = ps;
  }
  const hfEl = findChild(root, HEADER_FOOTER_TAG);
  if (hfEl) {
    const hf = parseHeaderFooter(hfEl);
    if (hf) ws.headerFooter = hf;
  }

  // <rowBreaks count="…"><brk id="…" man="1"/></rowBreaks> + <colBreaks…>
  const rbEl = findChild(root, ROW_BREAKS_TAG);
  if (rbEl) {
    for (const brk of findChildren(rbEl, BRK_TAG)) ws.rowBreaks.push(parsePageBreak(brk));
  }
  const cbEl = findChild(root, COL_BREAKS_TAG);
  if (cbEl) {
    for (const brk of findChildren(cbEl, BRK_TAG)) ws.colBreaks.push(parsePageBreak(brk));
  }

  // <customProperties><customProperty name="…" r:id="rIdN"/></customProperties>
  const cpEl = findChild(root, CUSTOM_PROPERTIES_TAG);
  if (cpEl) {
    for (const cp of findChildren(cpEl, CUSTOM_PROPERTY_TAG)) {
      const name = cp.attrs['name'];
      if (!name) continue;
      const entry: WorksheetCustomProperty = { name };
      const rId = cp.attrs[`{${REL_NS}}id`];
      if (rId) entry.rId = rId;
      ws.customProperties.push(entry);
    }
  }

  // <webPublishItems><webPublishItem .../></webPublishItems>
  const wpEl = findChild(root, WEB_PUBLISH_ITEMS_TAG);
  if (wpEl) {
    for (const wp of findChildren(wpEl, WEB_PUBLISH_ITEM_TAG)) {
      const item = parseWebPublishItem(wp);
      if (item) ws.webPublishItems.push(item);
    }
  }

  // <phoneticPr fontId="…" type="…" alignment="…"/>
  const ppEl = findChild(root, PHONETIC_PR_TAG);
  if (ppEl) {
    const pp = parsePhoneticPr(ppEl);
    if (pp) ws.phoneticPr = pp;
  }

  // <dataConsolidate function="sum"
  // topLabels="1"…><dataRefs>…</dataRefs></dataConsolidate>
  const dcEl = findChild(root, DATA_CONSOLIDATE_TAG);
  if (dcEl) {
    const dc = parseDataConsolidate(dcEl);
    if (dc) ws.dataConsolidate = dc;
  }

  // <scenarios current="…" show="…" sqref="…"><scenario name="…"> <inputCells
  // r="…" val="…"/> </scenario></scenarios>
  const scEl = findChild(root, SCENARIOS_TAG);
  if (scEl) {
    const sl = parseScenarioList(scEl);
    if (sl) ws.scenarios = sl;
  }

  // <cellWatches><cellWatch r="…"/></cellWatches>
  const cwWrap = findChild(root, CELL_WATCHES_TAG);
  if (cwWrap) {
    for (const w of findChildren(cwWrap, CELL_WATCH_TAG)) {
      const ref = w.attrs['r'];
      if (ref) ws.cellWatches.push({ ref });
    }
  }

  // <ignoredErrors><ignoredError sqref="…" evalError="1" .../></ignoredErrors>
  const ieWrap = findChild(root, IGNORED_ERRORS_TAG);
  if (ieWrap) {
    for (const ie of findChildren(ieWrap, IGNORED_ERROR_TAG)) {
      ws.ignoredErrors.push(parseIgnoredError(ie));
    }
  }

  captureWorksheetBodyExtras(root, ws);
  return ws;
}

const SORT_BY_VALUES: ReadonlyArray<SortBy> = ['value', 'cellColor', 'fontColor', 'icon'];
const SORT_METHODS: ReadonlyArray<SortMethod> = ['stroke', 'pinYin'];
const SORT_ICON_SETS: ReadonlyArray<SortIconSet> = [
  '3Arrows',
  '3ArrowsGray',
  '3Flags',
  '3TrafficLights1',
  '3TrafficLights2',
  '3Signs',
  '3Symbols',
  '3Symbols2',
  '4Arrows',
  '4ArrowsGray',
  '4RedToBlack',
  '4Rating',
  '4TrafficLights',
  '5Arrows',
  '5ArrowsGray',
  '5Rating',
  '5Quarters',
];

const parseSortState = (node: XmlNode): SortState | undefined => {
  const ref = node.attrs['ref'];
  if (!ref) return undefined;
  const out: SortState = { ref, conditions: [] };
  const cs = parseBoolXmlAttr(node.attrs['columnSort']);
  if (cs !== undefined) out.columnSort = cs;
  const cse = parseBoolXmlAttr(node.attrs['caseSensitive']);
  if (cse !== undefined) out.caseSensitive = cse;
  const sm = node.attrs['sortMethod'];
  if (sm && SORT_METHODS.includes(sm as SortMethod)) out.sortMethod = sm as SortMethod;

  for (const sc of findChildren(node, SORT_CONDITION_TAG)) {
    const cRef = sc.attrs['ref'];
    if (!cRef) continue;
    const c: SortCondition = { ref: cRef };
    const desc = parseBoolXmlAttr(sc.attrs['descending']);
    if (desc !== undefined) c.descending = desc;
    const sb = sc.attrs['sortBy'];
    if (sb && SORT_BY_VALUES.includes(sb as SortBy)) c.sortBy = sb as SortBy;
    if (sc.attrs['customList'] !== undefined) c.customList = sc.attrs['customList'];
    if (sc.attrs['dxfId'] !== undefined) {
      const n = Number.parseInt(sc.attrs['dxfId'], 10);
      if (Number.isInteger(n)) c.dxfId = n;
    }
    const is = sc.attrs['iconSet'];
    if (is && SORT_ICON_SETS.includes(is as SortIconSet)) c.iconSet = is as SortIconSet;
    if (sc.attrs['iconId'] !== undefined) {
      const n = Number.parseInt(sc.attrs['iconId'], 10);
      if (Number.isInteger(n)) c.iconId = n;
    }
    out.conditions.push(c);
  }
  return out;
};

const parseScenarioList = (node: XmlNode): ScenarioList | undefined => {
  const out: ScenarioList = { scenarios: [] };
  const current = parseIntegerAttr(node.attrs['current']);
  if (current !== undefined) out.current = current;
  const show = parseIntegerAttr(node.attrs['show']);
  if (show !== undefined) out.show = show;
  const sqref = node.attrs['sqref'];
  if (sqref) out.sqref = parseMultiCellRange(sqref);

  for (const sNode of findChildren(node, SCENARIO_TAG)) {
    const s = parseScenario(sNode);
    if (s) out.scenarios.push(s);
  }
  return out.scenarios.length > 0 || out.current !== undefined || out.show !== undefined || out.sqref !== undefined
    ? out
    : undefined;
};

const parseScenario = (node: XmlNode): Scenario | undefined => {
  const name = node.attrs['name'];
  if (!name) return undefined;
  const out: Scenario = { name, inputCells: [] };
  const locked = parseBoolXmlAttr(node.attrs['locked']);
  if (locked !== undefined) out.locked = locked;
  const hidden = parseBoolXmlAttr(node.attrs['hidden']);
  if (hidden !== undefined) out.hidden = hidden;
  if (node.attrs['user'] !== undefined) out.user = node.attrs['user'];
  if (node.attrs['comment'] !== undefined) out.comment = node.attrs['comment'];

  for (const ic of findChildren(node, INPUT_CELLS_TAG)) {
    const cell = parseScenarioInputCell(ic);
    if (cell) out.inputCells.push(cell);
  }
  return out;
};

const parseScenarioInputCell = (node: XmlNode): ScenarioInputCell | undefined => {
  const ref = node.attrs['r'];
  const val = node.attrs['val'];
  if (!ref || val === undefined) return undefined;
  const out: ScenarioInputCell = { ref, val };
  const deleted = parseBoolXmlAttr(node.attrs['deleted']);
  if (deleted !== undefined) out.deleted = deleted;
  const undone = parseBoolXmlAttr(node.attrs['undone']);
  if (undone !== undefined) out.undone = undone;
  const numFmtId = parseIntegerAttr(node.attrs['numFmtId']);
  if (numFmtId !== undefined) out.numFmtId = numFmtId;
  return out;
};

const DATA_CONSOLIDATE_FUNCTIONS: ReadonlyArray<DataConsolidateFunction> = [
  'average',
  'count',
  'countNums',
  'max',
  'min',
  'product',
  'stdDev',
  'stdDevp',
  'sum',
  'var',
  'varp',
];

const parseDataConsolidate = (node: XmlNode): DataConsolidate | undefined => {
  const out: DataConsolidate = {};
  const f = node.attrs['function'];
  if (f && DATA_CONSOLIDATE_FUNCTIONS.includes(f as DataConsolidateFunction)) {
    out.function = f as DataConsolidateFunction;
  }
  const topLabels = parseBoolXmlAttr(node.attrs['topLabels']);
  if (topLabels !== undefined) out.topLabels = topLabels;
  const leftLabels = parseBoolXmlAttr(node.attrs['leftLabels']);
  if (leftLabels !== undefined) out.leftLabels = leftLabels;
  const link = parseBoolXmlAttr(node.attrs['link']);
  if (link !== undefined) out.link = link;
  if (node.attrs['startLabels'] !== undefined) out.startLabels = node.attrs['startLabels'];

  const refsEl = findChild(node, DATA_REFS_TAG);
  if (refsEl) {
    const refs: DataReference[] = [];
    for (const ref of findChildren(refsEl, DATA_REF_TAG)) {
      const entry: DataReference = {};
      if (ref.attrs['name'] !== undefined) entry.name = ref.attrs['name'];
      if (ref.attrs['ref'] !== undefined) entry.ref = ref.attrs['ref'];
      if (ref.attrs['sheet'] !== undefined) entry.sheet = ref.attrs['sheet'];
      const rId = ref.attrs[`{${REL_NS}}id`];
      if (rId) entry.rId = rId;
      refs.push(entry);
    }
    if (refs.length > 0) out.dataRefs = refs;
  }
  return Object.keys(out).length > 0 ? out : undefined;
};

const PHONETIC_TYPES: ReadonlyArray<PhoneticType> = [
  'halfwidthKatakana',
  'fullwidthKatakana',
  'Hiragana',
  'noConversion',
];
const PHONETIC_ALIGNMENTS: ReadonlyArray<PhoneticAlignment> = [
  'noControl',
  'left',
  'center',
  'distributed',
];

const parsePhoneticPr = (node: XmlNode): WorksheetPhoneticProperties | undefined => {
  const out: WorksheetPhoneticProperties = {};
  const fontId = parseIntegerAttr(node.attrs['fontId']);
  if (fontId !== undefined) out.fontId = fontId;
  const t = node.attrs['type'];
  if (t && PHONETIC_TYPES.includes(t as PhoneticType)) out.type = t as PhoneticType;
  const a = node.attrs['alignment'];
  if (a && PHONETIC_ALIGNMENTS.includes(a as PhoneticAlignment)) out.alignment = a as PhoneticAlignment;
  return Object.keys(out).length > 0 ? out : undefined;
};

const VALID_WP_SOURCE_TYPES: ReadonlySet<WebPublishItem['sourceType']> = new Set([
  'sheet',
  'printArea',
  'autoFilter',
  'range',
  'chart',
  'pivotTable',
  'query',
  'label',
]);

export const parseWebPublishItem = (node: XmlNode): WebPublishItem | undefined => {
  const idRaw = node.attrs['id'];
  if (!idRaw) return undefined;
  const id = Number.parseInt(idRaw, 10);
  if (!Number.isInteger(id)) return undefined;
  const divId = node.attrs['divId'];
  const sourceType = node.attrs['sourceType'] as WebPublishItem['sourceType'] | undefined;
  const destinationFile = node.attrs['destinationFile'];
  if (!divId || !sourceType || !VALID_WP_SOURCE_TYPES.has(sourceType) || !destinationFile) return undefined;
  const out: WebPublishItem = { id, divId, sourceType, destinationFile };
  if (node.attrs['sourceRef'] !== undefined) out.sourceRef = node.attrs['sourceRef'];
  if (node.attrs['sourceObject'] !== undefined) out.sourceObject = node.attrs['sourceObject'];
  if (node.attrs['title'] !== undefined) out.title = node.attrs['title'];
  const auto = parseBoolXmlAttr(node.attrs['autoRepublish']);
  if (auto !== undefined) out.autoRepublish = auto;
  return out;
};

const CUSTOM_SHEET_VIEW_STATES: ReadonlyArray<CustomSheetViewState> = [
  'visible',
  'hidden',
  'veryHidden',
];

const parseCustomSheetView = (node: XmlNode): CustomSheetView | undefined => {
  const guid = node.attrs['guid'];
  if (!guid) return undefined;
  const out: CustomSheetView = { guid };
  const intAttr = (k: string): number | undefined => {
    const raw = node.attrs[k];
    if (raw === undefined) return undefined;
    const n = Number.parseInt(raw, 10);
    return Number.isInteger(n) ? n : undefined;
  };
  const scale = intAttr('scale');
  if (scale !== undefined) out.scale = scale;
  const colorId = intAttr('colorId');
  if (colorId !== undefined) out.colorId = colorId;

  const boolKeys = [
    'showPageBreaks',
    'showFormulas',
    'showGridLines',
    'showRowCol',
    'outlineSymbols',
    'zeroValues',
    'fitToPage',
    'printArea',
    'filter',
    'showAutoFilter',
    'hiddenRows',
    'hiddenColumns',
    'filterUnique',
    'showRuler',
  ] as const satisfies ReadonlyArray<keyof CustomSheetView>;
  for (const k of boolKeys) {
    const v = parseBoolXmlAttr(node.attrs[k]);
    if (v !== undefined) out[k] = v;
  }
  const stateRaw = node.attrs['state'];
  if (stateRaw && CUSTOM_SHEET_VIEW_STATES.includes(stateRaw as CustomSheetViewState)) {
    out.state = stateRaw as CustomSheetViewState;
  }
  const viewRaw = node.attrs['view'];
  if (viewRaw && SHEET_VIEW_MODES.includes(viewRaw as SheetViewMode)) {
    out.view = viewRaw as SheetViewMode;
  }
  if (node.attrs['topLeftCell'] !== undefined) out.topLeftCell = node.attrs['topLeftCell'];

  // Children — reuse existing parsers.
  const paneEl = findChild(node, PANE_TAG);
  if (paneEl) out.pane = parsePane(paneEl);
  const selections: Selection[] = [];
  for (const s of findChildren(node, SELECTION_TAG)) selections.push(parseSelection(s));
  if (selections.length > 0) out.selections = selections;

  const rbEl = findChild(node, ROW_BREAKS_TAG);
  if (rbEl) {
    const arr: PageBreak[] = [];
    for (const brk of findChildren(rbEl, BRK_TAG)) arr.push(parsePageBreak(brk));
    if (arr.length > 0) out.rowBreaks = arr;
  }
  const cbEl = findChild(node, COL_BREAKS_TAG);
  if (cbEl) {
    const arr: PageBreak[] = [];
    for (const brk of findChildren(cbEl, BRK_TAG)) arr.push(parsePageBreak(brk));
    if (arr.length > 0) out.colBreaks = arr;
  }

  const pmEl = findChild(node, PAGE_MARGINS_TAG);
  if (pmEl) {
    const pm = parsePageMargins(pmEl);
    if (pm) out.pageMargins = pm;
  }
  const poEl = findChild(node, PRINT_OPTIONS_TAG);
  if (poEl) {
    const po = parsePrintOptions(poEl);
    if (po) out.printOptions = po;
  }
  const psEl = findChild(node, PAGE_SETUP_TAG);
  if (psEl) {
    const ps = parsePageSetup(psEl);
    if (ps) out.pageSetup = ps;
  }
  const hfEl = findChild(node, HEADER_FOOTER_TAG);
  if (hfEl) {
    const hf = parseHeaderFooter(hfEl);
    if (hf) out.headerFooter = hf;
  }
  return out;
};

const OLE_DV_ASPECTS: ReadonlyArray<OleDvAspect> = ['DVASPECT_CONTENT', 'DVASPECT_ICON'];
const OLE_UPDATE_MODES: ReadonlyArray<OleUpdateMode> = ['OLEUPDATE_ALWAYS', 'OLEUPDATE_ONCALL'];

const parseOleObject = (node: XmlNode): OleObject | undefined => {
  const shapeIdRaw = node.attrs['shapeId'];
  if (shapeIdRaw === undefined) return undefined;
  const shapeId = Number.parseInt(shapeIdRaw, 10);
  if (!Number.isInteger(shapeId)) return undefined;
  const out: OleObject = { shapeId };
  const rId = node.attrs[`{${REL_NS}}id`];
  if (rId) out.rId = rId;
  if (node.attrs['progId'] !== undefined) out.progId = node.attrs['progId'];
  const dvAspect = node.attrs['dvAspect'];
  if (dvAspect && OLE_DV_ASPECTS.includes(dvAspect as OleDvAspect)) out.dvAspect = dvAspect as OleDvAspect;
  if (node.attrs['link'] !== undefined) out.link = node.attrs['link'];
  const oleUpdate = node.attrs['oleUpdate'];
  if (oleUpdate && OLE_UPDATE_MODES.includes(oleUpdate as OleUpdateMode))
    out.oleUpdate = oleUpdate as OleUpdateMode;
  const autoLoad = parseBoolXmlAttr(node.attrs['autoLoad']);
  if (autoLoad !== undefined) out.autoLoad = autoLoad;
  const objectPr = findChild(node, OBJECT_PR_TAG);
  if (objectPr) out.objectPr = objectPr;
  return out;
};

const parseFormControl = (node: XmlNode): FormControl | undefined => {
  const shapeIdRaw = node.attrs['shapeId'];
  if (shapeIdRaw === undefined) return undefined;
  const shapeId = Number.parseInt(shapeIdRaw, 10);
  if (!Number.isInteger(shapeId)) return undefined;
  const out: FormControl = { shapeId };
  const rId = node.attrs[`{${REL_NS}}id`];
  if (rId) out.rId = rId;
  if (node.attrs['name'] !== undefined) out.name = node.attrs['name'];
  const controlPr = findChild(node, CONTROL_PR_TAG);
  if (controlPr) out.controlPr = controlPr;
  return out;
};

const parsePageBreak = (node: XmlNode): PageBreak => {
  const out: PageBreak = {};
  const id = parseIntegerAttr(node.attrs['id']);
  if (id !== undefined) out.id = id;
  const min = parseIntegerAttr(node.attrs['min']);
  if (min !== undefined) out.min = min;
  const max = parseIntegerAttr(node.attrs['max']);
  if (max !== undefined) out.max = max;
  const man = parseBoolXmlAttr(node.attrs['man']);
  if (man !== undefined) out.man = man;
  const pt = parseBoolXmlAttr(node.attrs['pt']);
  if (pt !== undefined) out.pt = pt;
  return out;
};

const parseBoolFlag = (raw: string | undefined): boolean | undefined => {
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  return undefined;
};

const parsePrintOptions = (node: XmlNode): PrintOptions | undefined => {
  const out: PrintOptions = {};
  const hc = parseBoolFlag(node.attrs['horizontalCentered']);
  if (hc !== undefined) out.horizontalCentered = hc;
  const vc = parseBoolFlag(node.attrs['verticalCentered']);
  if (vc !== undefined) out.verticalCentered = vc;
  const headings = parseBoolFlag(node.attrs['headings']);
  if (headings !== undefined) out.headings = headings;
  const gl = parseBoolFlag(node.attrs['gridLines']);
  if (gl !== undefined) out.gridLines = gl;
  const gls = parseBoolFlag(node.attrs['gridLinesSet']);
  if (gls !== undefined) out.gridLinesSet = gls;
  return Object.keys(out).length > 0 ? out : undefined;
};

export const parsePageMargins = (node: XmlNode): PageMargins | undefined => {
  const left = parseFloatAttr(node.attrs['left']);
  const right = parseFloatAttr(node.attrs['right']);
  const top = parseFloatAttr(node.attrs['top']);
  const bottom = parseFloatAttr(node.attrs['bottom']);
  const header = parseFloatAttr(node.attrs['header']);
  const footer = parseFloatAttr(node.attrs['footer']);
  if (
    left === undefined ||
    right === undefined ||
    top === undefined ||
    bottom === undefined ||
    header === undefined ||
    footer === undefined
  ) {
    return undefined;
  }
  return { left, right, top, bottom, header, footer };
};

const PAGE_ORIENTATIONS: ReadonlyArray<PageOrientation> = ['default', 'portrait', 'landscape'];
const PAGE_ORDERS: ReadonlyArray<PageOrder> = ['downThenOver', 'overThenDown'];
const CELL_COMMENT_MODES: ReadonlyArray<CellCommentMode> = ['none', 'asDisplayed', 'atEnd'];
const PRINT_ERROR_MODES: ReadonlyArray<PrintErrorMode> = ['displayed', 'blank', 'dash', 'NA'];

export const parsePageSetup = (node: XmlNode): PageSetup | undefined => {
  const out: PageSetup = {};
  const intAttr = (k: string): void => {
    const v = parseIntegerAttr(node.attrs[k]);
    if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  };
  intAttr('paperSize');
  intAttr('scale');
  intAttr('firstPageNumber');
  intAttr('fitToWidth');
  intAttr('fitToHeight');
  intAttr('horizontalDpi');
  intAttr('verticalDpi');
  intAttr('copies');

  const ord = node.attrs['pageOrder'];
  if (ord && PAGE_ORDERS.includes(ord as PageOrder)) out.pageOrder = ord as PageOrder;
  const ori = node.attrs['orientation'];
  if (ori && PAGE_ORIENTATIONS.includes(ori as PageOrientation)) out.orientation = ori as PageOrientation;
  const cc = node.attrs['cellComments'];
  if (cc && CELL_COMMENT_MODES.includes(cc as CellCommentMode)) out.cellComments = cc as CellCommentMode;
  const errs = node.attrs['errors'];
  if (errs && PRINT_ERROR_MODES.includes(errs as PrintErrorMode)) out.errors = errs as PrintErrorMode;

  const upd = parseBoolFlag(node.attrs['usePrinterDefaults']);
  if (upd !== undefined) out.usePrinterDefaults = upd;
  const bw = parseBoolFlag(node.attrs['blackAndWhite']);
  if (bw !== undefined) out.blackAndWhite = bw;
  const draft = parseBoolFlag(node.attrs['draft']);
  if (draft !== undefined) out.draft = draft;
  const ufpn = parseBoolFlag(node.attrs['useFirstPageNumber']);
  if (ufpn !== undefined) out.useFirstPageNumber = ufpn;

  if (node.attrs['paperWidth']) out.paperWidth = node.attrs['paperWidth'];
  if (node.attrs['paperHeight']) out.paperHeight = node.attrs['paperHeight'];
  const rId = node.attrs[`{${REL_NS}}id`];
  if (rId) out.rId = rId;

  return Object.keys(out).length > 0 ? out : undefined;
};

export const parseHeaderFooter = (node: XmlNode): HeaderFooter | undefined => {
  const out: HeaderFooter = {};
  const df = parseBoolFlag(node.attrs['differentFirst']);
  if (df !== undefined) out.differentFirst = df;
  const doe = parseBoolFlag(node.attrs['differentOddEven']);
  if (doe !== undefined) out.differentOddEven = doe;
  const swd = parseBoolFlag(node.attrs['scaleWithDoc']);
  if (swd !== undefined) out.scaleWithDoc = swd;
  const awm = parseBoolFlag(node.attrs['alignWithMargins']);
  if (awm !== undefined) out.alignWithMargins = awm;

  const text = (tag: string): string | undefined => {
    const child = findChild(node, tag);
    if (!child) return undefined;
    return child.text;
  };
  const oh = text(ODD_HEADER_TAG);
  if (oh !== undefined) out.oddHeader = oh;
  const of = text(ODD_FOOTER_TAG);
  if (of !== undefined) out.oddFooter = of;
  const eh = text(EVEN_HEADER_TAG);
  if (eh !== undefined) out.evenHeader = eh;
  const ef = text(EVEN_FOOTER_TAG);
  if (ef !== undefined) out.evenFooter = ef;
  const fh = text(FIRST_HEADER_TAG);
  if (fh !== undefined) out.firstHeader = fh;
  const ff = text(FIRST_FOOTER_TAG);
  if (ff !== undefined) out.firstFooter = ff;

  return Object.keys(out).length > 0 ? out : undefined;
};

const parseSheetProtection = (node: XmlNode): SheetProtection => {
  const out: SheetProtection = {};
  const flag = (k: string): void => {
    const raw = node.attrs[k];
    if (raw === '1' || raw === 'true') (out as Record<string, unknown>)[k] = true;
    else if (raw === '0' || raw === 'false') (out as Record<string, unknown>)[k] = false;
  };
  for (const k of [
    'sheet',
    'objects',
    'scenarios',
    'formatCells',
    'formatColumns',
    'formatRows',
    'insertColumns',
    'insertRows',
    'insertHyperlinks',
    'deleteColumns',
    'deleteRows',
    'selectLockedCells',
    'selectUnlockedCells',
    'sort',
    'autoFilter',
    'pivotTables',
  ]) {
    flag(k);
  }
  if (node.attrs['saltValue'] !== undefined) out.saltValue = node.attrs['saltValue'];
  if (node.attrs['spinCount'] !== undefined) {
    const n = Number.parseInt(node.attrs['spinCount'], 10);
    if (Number.isInteger(n)) out.spinCount = n;
  }
  if (node.attrs['algorithmName'] !== undefined) out.algorithmName = node.attrs['algorithmName'];
  if (node.attrs['hashValue'] !== undefined) out.hashValue = node.attrs['hashValue'];
  return out;
};

const parseSheetProperties = (node: XmlNode): SheetProperties | undefined => {
  const out: SheetProperties = {};
  const flag = (raw: string | undefined): boolean | undefined => {
    if (raw === undefined) return undefined;
    if (raw === '1' || raw === 'true') return true;
    if (raw === '0' || raw === 'false') return false;
    return undefined;
  };

  if (node.attrs['codeName'] !== undefined) out.codeName = node.attrs['codeName'];
  const efcc = flag(node.attrs['enableFormatConditionsCalculation']);
  if (efcc !== undefined) out.enableFormatConditionsCalculation = efcc;
  const fm = flag(node.attrs['filterMode']);
  if (fm !== undefined) out.filterMode = fm;
  const pub = flag(node.attrs['published']);
  if (pub !== undefined) out.published = pub;
  const sh = flag(node.attrs['syncHorizontal']);
  if (sh !== undefined) out.syncHorizontal = sh;
  if (node.attrs['syncRef'] !== undefined) out.syncRef = node.attrs['syncRef'];
  const sv = flag(node.attrs['syncVertical']);
  if (sv !== undefined) out.syncVertical = sv;
  const te = flag(node.attrs['transitionEvaluation']);
  if (te !== undefined) out.transitionEvaluation = te;
  const tre = flag(node.attrs['transitionEntry']);
  if (tre !== undefined) out.transitionEntry = tre;

  const tabColorEl = findChild(node, TAB_COLOR_TAG);
  if (tabColorEl) {
    const tcOpts: { rgb?: string; indexed?: number; theme?: number; auto?: boolean; tint?: number } = {};
    if (tabColorEl.attrs['rgb'] !== undefined) tcOpts.rgb = tabColorEl.attrs['rgb'];
    if (tabColorEl.attrs['indexed'] !== undefined) {
      const n = Number.parseInt(tabColorEl.attrs['indexed'], 10);
      if (Number.isInteger(n)) tcOpts.indexed = n;
    }
    if (tabColorEl.attrs['theme'] !== undefined) {
      const n = Number.parseInt(tabColorEl.attrs['theme'], 10);
      if (Number.isInteger(n)) tcOpts.theme = n;
    }
    const auto = flag(tabColorEl.attrs['auto']);
    if (auto !== undefined) tcOpts.auto = auto;
    if (tabColorEl.attrs['tint'] !== undefined) {
      const n = Number.parseFloat(tabColorEl.attrs['tint']);
      if (Number.isFinite(n)) tcOpts.tint = n;
    }
    out.tabColor = makeColor(tcOpts);
  }

  const outlineEl = findChild(node, OUTLINE_PR_TAG);
  if (outlineEl) {
    const op: OutlineProperties = {};
    const aS = flag(outlineEl.attrs['applyStyles']);
    if (aS !== undefined) op.applyStyles = aS;
    const sB = flag(outlineEl.attrs['summaryBelow']);
    if (sB !== undefined) op.summaryBelow = sB;
    const sR = flag(outlineEl.attrs['summaryRight']);
    if (sR !== undefined) op.summaryRight = sR;
    const sOS = flag(outlineEl.attrs['showOutlineSymbols']);
    if (sOS !== undefined) op.showOutlineSymbols = sOS;
    out.outlinePr = op;
  }

  const psEl = findChild(node, PAGE_SETUP_PR_TAG);
  if (psEl) {
    const ps: PageSetupProperties = {};
    const apb = flag(psEl.attrs['autoPageBreaks']);
    if (apb !== undefined) ps.autoPageBreaks = apb;
    const ftp = flag(psEl.attrs['fitToPage']);
    if (ftp !== undefined) ps.fitToPage = ftp;
    out.pageSetUpPr = ps;
  }

  return Object.keys(out).length > 0 ? out : undefined;
};

const parseIgnoredError = (node: XmlNode): IgnoredError => {
  const sqref = node.attrs['sqref'];
  if (!sqref) throw new OpenXmlSchemaError('worksheet: <ignoredError> missing @sqref');
  const out: IgnoredError = { sqref: parseMultiCellRange(sqref) };
  const flag = (raw: string | undefined): boolean | undefined => {
    if (raw === undefined) return undefined;
    if (raw === '1' || raw === 'true') return true;
    if (raw === '0' || raw === 'false') return false;
    return undefined;
  };
  const evalError = flag(node.attrs['evalError']);
  if (evalError !== undefined) out.evalError = evalError;
  const twoDigitTextYear = flag(node.attrs['twoDigitTextYear']);
  if (twoDigitTextYear !== undefined) out.twoDigitTextYear = twoDigitTextYear;
  const numberStoredAsText = flag(node.attrs['numberStoredAsText']);
  if (numberStoredAsText !== undefined) out.numberStoredAsText = numberStoredAsText;
  const formula = flag(node.attrs['formula']);
  if (formula !== undefined) out.formula = formula;
  const formulaRange = flag(node.attrs['formulaRange']);
  if (formulaRange !== undefined) out.formulaRange = formulaRange;
  const unlockedFormula = flag(node.attrs['unlockedFormula']);
  if (unlockedFormula !== undefined) out.unlockedFormula = unlockedFormula;
  const emptyCellReference = flag(node.attrs['emptyCellReference']);
  if (emptyCellReference !== undefined) out.emptyCellReference = emptyCellReference;
  const listDataValidation = flag(node.attrs['listDataValidation']);
  if (listDataValidation !== undefined) out.listDataValidation = listDataValidation;
  const calculatedColumn = flag(node.attrs['calculatedColumn']);
  if (calculatedColumn !== undefined) out.calculatedColumn = calculatedColumn;
  return out;
};

/**
 * Pick up every top-level `<worksheet>` child we don't model (e.g. `<sheetPr>`,
 * `<printOptions>`, `<pageMargins>`, `<pageSetup>`, `<headerFooter>`,
 * `<rowBreaks>`, `<colBreaks>`, `<oleObjects>`, `<controls>`, `<picture>`,
 * `<legacyDrawingHF>`, `<extLst>`) so the writer can re-emit them around the
 * modeled blocks. Anything before `<sheetData>` (in document order) goes into
 * the early bucket; the rest goes into the late bucket.
 */
function captureWorksheetBodyExtras(root: XmlNode, ws: Worksheet): void {
  const beforeSheetData: XmlNode[] = [];
  const afterSheetData: XmlNode[] = [];
  let seenSheetData = false;
  for (const child of root.children) {
    if (MODELED_WORKSHEET_TAGS.has(child.name) || wrapsModeledElement(child)) {
      if (child.name === SHEETDATA_TAG) seenSheetData = true;
      continue;
    }
    if (seenSheetData) afterSheetData.push(child);
    else beforeSheetData.push(child);
  }
  if (beforeSheetData.length > 0 || afterSheetData.length > 0) {
    ws.bodyExtras = { beforeSheetData, afterSheetData };
  }
}

/**
 * Excel 2010+ gates `<controls>` / `<oleObjects>` (and each entry inside them)
 * behind `<mc:AlternateContent><mc:Choice Requires="x14">` so Excel 2007
 * skips the `<controlPr>` / `<objectPr>` children it predates. Collect the
 * named children whether they sit directly under `parent` or inside such a
 * wrapper. `mc:Fallback` only carries a degraded copy of the same entry, so it
 * is not consulted.
 */
const findChildrenThroughMc = (parent: XmlNode, name: string): XmlNode[] => {
  const out: XmlNode[] = [];
  for (const child of parent.children) {
    if (child.name === name) {
      out.push(child);
    } else if (child.name === MC_ALTERNATE_CONTENT_TAG) {
      for (const choice of findChildren(child, MC_CHOICE_TAG)) out.push(...findChildren(choice, name));
    }
  }
  return out;
};

/** The only worksheet children the reader collects through an `mc:AlternateContent` wrapper. */
const MC_WRAPPED_TAGS: ReadonlySet<string> = new Set([OLE_OBJECTS_TAG, CONTROLS_TAG]);

/**
 * True for an `mc:AlternateContent` whose Choice holds one of `MC_WRAPPED_TAGS`
 * — those are re-emitted from the typed model, so the wrapper must not also
 * land in bodyExtras. A wrapper around anything else is still a body extra.
 */
const wrapsModeledElement = (node: XmlNode): boolean =>
  node.name === MC_ALTERNATE_CONTENT_TAG &&
  findChildren(node, MC_CHOICE_TAG).some((choice) => choice.children.some((c) => MC_WRAPPED_TAGS.has(c.name)));

const MODELED_WORKSHEET_TAGS: ReadonlySet<string> = new Set([
  `{${SHEET_MAIN_NS}}dimension`,
  SHEETDATA_TAG,
  MERGE_CELLS_TAG,
  SHEET_VIEWS_TAG,
  SHEET_FORMAT_PR_TAG,
  COLS_TAG,
  HYPERLINKS_TAG,
  DATA_VALIDATIONS_TAG,
  AUTOFILTER_TAG,
  TABLE_PARTS_TAG,
  CONDITIONAL_FORMATTING_TAG,
  DRAWING_TAG,
  CELL_WATCHES_TAG,
  IGNORED_ERRORS_TAG,
  SHEET_PR_TAG,
  SHEET_PROTECTION_TAG,
  PROTECTED_RANGES_TAG,
  SORT_STATE_TAG,
  PICTURE_TAG,
  LEGACY_DRAWING_HF_TAG,
  OLE_OBJECTS_TAG,
  CONTROLS_TAG,
  SMART_TAGS_TAG,
  CUSTOM_SHEET_VIEWS_TAG,
  PRINT_OPTIONS_TAG,
  PAGE_MARGINS_TAG,
  PAGE_SETUP_TAG,
  HEADER_FOOTER_TAG,
  ROW_BREAKS_TAG,
  COL_BREAKS_TAG,
  CUSTOM_PROPERTIES_TAG,
  WEB_PUBLISH_ITEMS_TAG,
  PHONETIC_PR_TAG,
  DATA_CONSOLIDATE_TAG,
  SCENARIOS_TAG,
  LEGACY_DRAWING_TAG,
]);

const PANE_TYPES: ReadonlyArray<PaneType> = ['bottomRight', 'topRight', 'bottomLeft', 'topLeft'];
const PANE_STATES: ReadonlyArray<PaneState> = ['split', 'frozen', 'frozenSplit'];
const SHEET_VIEW_MODES: ReadonlyArray<SheetViewMode> = ['normal', 'pageBreakPreview', 'pageLayout'];

const parseFloatAttr = (raw: string | undefined): number | undefined => {
  if (raw === undefined) return undefined;
  const n = Number.parseFloat(raw);
  return Number.isFinite(n) ? n : undefined;
};

const parseIntegerAttr = (raw: string | undefined): number | undefined => {
  if (raw === undefined) return undefined;
  const n = Number.parseInt(raw, 10);
  return Number.isInteger(n) ? n : undefined;
};

const parseBoolXmlAttr = (raw: string | undefined): boolean | undefined => {
  if (raw === undefined) return undefined;
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  return undefined;
};

const parseSheetView = (node: XmlNode): SheetView => {
  const opts: Partial<SheetView> = {
    workbookViewId: parseIntegerAttr(node.attrs['workbookViewId']) ?? 0,
  };
  const tabSelected = parseBoolXmlAttr(node.attrs['tabSelected']);
  if (tabSelected !== undefined) opts.tabSelected = tabSelected;
  const showGridLines = parseBoolXmlAttr(node.attrs['showGridLines']);
  if (showGridLines !== undefined) opts.showGridLines = showGridLines;
  const showRowColHeaders = parseBoolXmlAttr(node.attrs['showRowColHeaders']);
  if (showRowColHeaders !== undefined) opts.showRowColHeaders = showRowColHeaders;
  const showFormulas = parseBoolXmlAttr(node.attrs['showFormulas']);
  if (showFormulas !== undefined) opts.showFormulas = showFormulas;
  const showZeros = parseBoolXmlAttr(node.attrs['showZeros']);
  if (showZeros !== undefined) opts.showZeros = showZeros;
  const rightToLeft = parseBoolXmlAttr(node.attrs['rightToLeft']);
  if (rightToLeft !== undefined) opts.rightToLeft = rightToLeft;
  const view = node.attrs['view'];
  if (view !== undefined && (SHEET_VIEW_MODES as ReadonlyArray<string>).includes(view)) {
    opts.view = view as SheetViewMode;
  }
  if (node.attrs['topLeftCell']) opts.topLeftCell = node.attrs['topLeftCell'];
  const zoomScale = parseIntegerAttr(node.attrs['zoomScale']);
  if (zoomScale !== undefined) opts.zoomScale = zoomScale;
  const zoomScaleNormal = parseIntegerAttr(node.attrs['zoomScaleNormal']);
  if (zoomScaleNormal !== undefined) opts.zoomScaleNormal = zoomScaleNormal;

  const paneEl = findChild(node, PANE_TAG);
  if (paneEl) opts.pane = parsePane(paneEl);
  const selectionEl = findChild(node, SELECTION_TAG);
  if (selectionEl) opts.selection = parseSelection(selectionEl);
  return makeSheetView(opts);
};

const parsePane = (node: XmlNode): Pane => {
  const stateRaw = node.attrs['state'];
  const state: PaneState =
    stateRaw && (PANE_STATES as ReadonlyArray<string>).includes(stateRaw) ? (stateRaw as PaneState) : 'split';
  const pane: Pane = { state };
  const xSplit = parseFloatAttr(node.attrs['xSplit']);
  if (xSplit !== undefined) pane.xSplit = xSplit;
  const ySplit = parseFloatAttr(node.attrs['ySplit']);
  if (ySplit !== undefined) pane.ySplit = ySplit;
  if (node.attrs['topLeftCell']) pane.topLeftCell = node.attrs['topLeftCell'];
  const activePaneRaw = node.attrs['activePane'];
  if (activePaneRaw && (PANE_TYPES as ReadonlyArray<string>).includes(activePaneRaw)) {
    pane.activePane = activePaneRaw as PaneType;
  }
  return pane;
};

const parseSelection = (node: XmlNode): Selection => {
  const sel: Selection = {};
  const paneRaw = node.attrs['pane'];
  if (paneRaw && (PANE_TYPES as ReadonlyArray<string>).includes(paneRaw)) sel.pane = paneRaw as PaneType;
  if (node.attrs['activeCell']) sel.activeCell = node.attrs['activeCell'];
  if (node.attrs['sqref']) sel.sqref = node.attrs['sqref'];
  return sel;
};

const parseRowIndex = (rowNode: XmlNode): number => {
  const rAttr = rowNode.attrs['r'];
  if (rAttr === undefined) {
    throw new OpenXmlSchemaError('worksheet: <row> missing required @r');
  }
  const r = Number.parseInt(rAttr, 10);
  if (!Number.isInteger(r) || r < 1) {
    throw new OpenXmlSchemaError(`worksheet: <row r="${rAttr}"> is not a positive integer`);
  }
  return r;
};

const parseCellCoord = (cNode: XmlNode, rowIdx: number, fallbackCol: number): { row: number; col: number } => {
  const rAttr = cNode.attrs['r'];
  if (rAttr === undefined) {
    // Cells without @r take the next column slot in the current row.
    return { row: rowIdx, col: fallbackCol };
  }
  const t = coordinateToTuple(rAttr);
  if (t.row !== rowIdx) {
    throw new OpenXmlSchemaError(`worksheet: <c r="${rAttr}"> row ${t.row} disagrees with <row r="${rowIdx}">`);
  }
  return { row: t.row, col: t.col };
};

const readCell = (
  ws: Worksheet,
  cNode: XmlNode,
  coord: { row: number; col: number },
  ctx: WorksheetReadContext,
  sharedFormulas: Map<number, SharedFormulaCache>,
): void => {
  const t = cNode.attrs['t'] ?? 'n';
  const styleAttr = cNode.attrs['s'];
  const styleId = styleAttr === undefined ? 0 : parseStyleId(styleAttr);

  const fNode = findChild(cNode, F_TAG);
  const vNode = findChild(cNode, V_TAG);
  const isNode = findChild(cNode, IS_TAG);

  // ---- formula path ------------------------------------------------------
  if (fNode) {
    const cell = setCell(ws, coord.row, coord.col, null, styleId);
    // `<v/>` and a missing `<v>` are different states: Excel writes the former
    // for a formula whose cached result is the empty string (common for links
    // into a workbook it cannot reach), and that cached value is the only
    // thing it has to display until the link resolves.
    const cachedRaw = vNode === undefined ? undefined : (vNode.text ?? '');
    const cached = decodeCachedValue(cachedRaw, t);
    handleFormula(cell, fNode, coord, cached, sharedFormulas);
    return;
  }

  // ---- non-formula values ------------------------------------------------
  let value: number | string | boolean | { kind: 'error'; code: ExcelErrorCode } | null = null;
  switch (t) {
    case 'n':
      value = vNode?.text !== undefined && vNode.text !== '' ? Number.parseFloat(vNode.text) : null;
      break;
    case 's': {
      if (vNode?.text === undefined) {
        throw new OpenXmlSchemaError('worksheet: <c t="s"> missing <v>');
      }
      const idx = Number.parseInt(vNode.text, 10);
      if (!Number.isInteger(idx) || idx < 0) {
        throw new OpenXmlSchemaError(`worksheet: <c t="s"><v>${vNode.text}</v> is not a valid index`);
      }
      const sst = ctx.sharedStrings[idx];
      if (sst === undefined) {
        throw new OpenXmlSchemaError(
          `worksheet: shared-string index ${idx} out of range [0, ${ctx.sharedStrings.length})`,
        );
      }
      value = sst;
      break;
    }
    case 'b':
      value = vNode?.text === '1';
      break;
    case 'e': {
      const code = vNode?.text;
      if (code === undefined || !ERROR_CODES.has(code)) {
        throw new OpenXmlSchemaError(`worksheet: unknown error code "${code}" in <c t="e">`);
      }
      value = { kind: 'error', code: code as ExcelErrorCode };
      break;
    }
    case 'str':
      value = vNode?.text ?? '';
      break;
    case 'inlineStr':
      value = readInlineString(isNode);
      break;
    default:
      throw new OpenXmlSchemaError(`worksheet: unknown cell type t="${t}"`);
  }
  setCell(ws, coord.row, coord.col, value, styleId);
};

const parseStyleId = (raw: string): number => {
  const n = Number.parseInt(raw, 10);
  if (!Number.isInteger(n) || n < 0) {
    throw new OpenXmlSchemaError(`worksheet: <c s="${raw}"> is not a non-negative integer`);
  }
  return n;
};

/** Inline-string body — concatenates `<is>/<t>` text runs. Rich runs lose formatting in stage-1. */
const readInlineString = (isNode: XmlNode | undefined): string => {
  if (!isNode) return '';
  const direct = findChild(isNode, T_TAG);
  if (direct) return direct.text ?? '';
  // Rich inline string — concatenate every <r>/<t>.
  let out = '';
  for (const child of isNode.children) {
    const t = findChild(child, T_TAG);
    if (t?.text) out += t.text;
  }
  return out;
};

const decodeCachedValue = (raw: string | undefined, t: string): number | string | boolean | undefined => {
  if (raw === undefined) return undefined;
  switch (t) {
    case 'n':
      // An empty `<v/>` under the (default) numeric type carries no number.
      return raw === '' ? undefined : Number.parseFloat(raw);
    case 'b':
      return raw === '' ? undefined : raw === '1';
    case 'str':
      return raw;
    case 'e':
      return raw;
    case 's':
      // Cached value of a formula resolving to a shared string is rare; keep
      // as-is.
      return raw;
    default:
      return raw;
  }
};

const handleFormula = (
  cell: Cell,
  fNode: XmlNode,
  coord: { row: number; col: number },
  cached: number | string | boolean | undefined,
  sharedFormulas: Map<number, SharedFormulaCache>,
): void => {
  const tAttr = fNode.attrs['t'] ?? 'normal';
  // Keep the formula text verbatim, including the `_xlfn.` / `_xlfn._xlws.`
  // prefixes Excel writes for future-functions (SCAN, LAMBDA, XLOOKUP, …).
  // The prefix is part of the stored grammar — a bare `SCAN(...)` is an
  // unknown name that Excel renders as #NAME?. Stripping it here and then
  // writing the model back verbatim corrupted every dynamic-array formula on
  // a load → save round-trip. openpyxl surfaces the prefix verbatim too.
  const formula = fNode.text ?? '';
  const opts = cached !== undefined ? { cachedValue: cached } : undefined;
  switch (tAttr as FormulaKind) {
    case 'normal':
      setFormula(cell, formula, opts);
      return;
    case 'array': {
      const ref = fNode.attrs['ref'];
      if (!ref) {
        throw new OpenXmlSchemaError('worksheet: <f t="array"> missing @ref');
      }
      setArrayFormula(cell, ref, formula, opts);
      return;
    }
    case 'shared': {
      const siRaw = fNode.attrs['si'];
      if (siRaw === undefined) {
        throw new OpenXmlSchemaError('worksheet: <f t="shared"> missing @si');
      }
      const si = Number.parseInt(siRaw, 10);
      if (!Number.isInteger(si) || si < 0) {
        throw new OpenXmlSchemaError(`worksheet: <f t="shared" si="${siRaw}"> is not a valid index`);
      }
      const ref = fNode.attrs['ref'];
      if (formula.length > 0) {
        // Origin / first occurrence of the shared formula. The ref is required
        // by Excel but we accept its absence for resilience.
        sharedFormulas.set(si, { origin: tupleToCoordinate(coord.col, coord.row), formula });
        setSharedFormula(cell, si, formula, ref, opts);
        return;
      }
      // Subsequent reference — translate the cached origin formula.
      const cache = sharedFormulas.get(si);
      if (!cache) {
        throw new OpenXmlSchemaError(`worksheet: <f t="shared" si="${si}"/> with no preceding origin formula`);
      }
      const dest = tupleToCoordinate(coord.col, coord.row);
      // OOXML shared-formula text omits the leading '='; the translator treats
      // unprefixed input as a LITERAL and skips ref shifting, so we re-prefix
      // before translating and strip again on the way out.
      const translated = translateFormula(`=${cache.formula}`, cache.origin, { dest });
      const stripped = translated.startsWith('=') ? translated.slice(1) : translated;
      setSharedFormula(cell, si, stripped, undefined, opts);
      return;
    }
    case 'dataTable': {
      // Data-Table formula (What-if Analysis output). Preserve every
      // dt-specific attribute so the writer re-emits the exact same <f
      // t="dataTable" r1="…" /> shape and Excel keeps treating the cell as a
      // Data Table cell.
      const ref = fNode.attrs['ref'];
      if (!ref) {
        throw new OpenXmlSchemaError('worksheet: <f t="dataTable"> missing @ref');
      }
      const dtOpts: import('../cell/cell').DataTableFormulaOpts = {
        ref,
        ...(cached !== undefined ? { cachedValue: cached } : {}),
        ...(fNode.attrs['r1'] !== undefined ? { r1: fNode.attrs['r1'] } : {}),
        ...(fNode.attrs['r2'] !== undefined ? { r2: fNode.attrs['r2'] } : {}),
        ...(parseDataTableBool(fNode.attrs['dt2D']) ? { dt2D: true } : {}),
        ...(parseDataTableBool(fNode.attrs['dtr']) ? { dtr: true } : {}),
        ...(parseDataTableBool(fNode.attrs['del1']) ? { del1: true } : {}),
        ...(parseDataTableBool(fNode.attrs['del2']) ? { del2: true } : {}),
        ...(parseDataTableBool(fNode.attrs['aca']) ? { aca: true } : {}),
        ...(parseDataTableBool(fNode.attrs['ca']) ? { ca: true } : {}),
      };
      setDataTableFormula(cell, formula, dtOpts);
      return;
    }
    default:
      throw new OpenXmlSchemaError(`worksheet: <f t="${tAttr}"> unknown formula kind`);
  }
};

const parseDataTableBool = (raw: string | undefined): boolean => {
  // Excel emits these as the OOXML truthy values "1" or "true".
  if (raw === undefined) return false;
  return raw === '1' || raw === 'true';
};

const parseColumnDimension = (node: XmlNode): ColumnDimension => {
  const minRaw = node.attrs['min'];
  const maxRaw = node.attrs['max'];
  if (minRaw === undefined || maxRaw === undefined) {
    throw new OpenXmlSchemaError('worksheet: <col> missing required @min/@max');
  }
  const min = Number.parseInt(minRaw, 10);
  const max = Number.parseInt(maxRaw, 10);
  if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min) {
    throw new OpenXmlSchemaError(`worksheet: <col min="${minRaw}" max="${maxRaw}"> not a valid column run`);
  }
  const opts: Partial<Omit<ColumnDimension, 'min' | 'max'>> = {};
  const width = parseFloatAttr(node.attrs['width']);
  if (width !== undefined) opts.width = width;
  const customWidth = parseBoolXmlAttr(node.attrs['customWidth']);
  if (customWidth !== undefined) opts.customWidth = customWidth;
  const hidden = parseBoolXmlAttr(node.attrs['hidden']);
  if (hidden !== undefined) opts.hidden = hidden;
  const bestFit = parseBoolXmlAttr(node.attrs['bestFit']);
  if (bestFit !== undefined) opts.bestFit = bestFit;
  const outlineLevel = parseIntegerAttr(node.attrs['outlineLevel']);
  if (outlineLevel !== undefined) opts.outlineLevel = outlineLevel;
  const style = parseIntegerAttr(node.attrs['style']);
  if (style !== undefined) opts.style = style;
  const collapsed = parseBoolXmlAttr(node.attrs['collapsed']);
  if (collapsed !== undefined) opts.collapsed = collapsed;
  // makeColumnDimension fills min=col=max; rebuild with both range ends.
  return { ...makeColumnDimension(min, opts), max };
};

const DV_TYPES: ReadonlyArray<DataValidationType> = [
  'whole',
  'decimal',
  'list',
  'date',
  'time',
  'textLength',
  'custom',
];
const DV_OPERATORS: ReadonlyArray<DataValidationOperator> = [
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
];
const DV_ERROR_STYLES: ReadonlyArray<DataValidationErrorStyle> = ['stop', 'warning', 'information'];

const parseDataValidation = (node: XmlNode): DataValidation => {
  const typeRaw = node.attrs['type'];
  const type: DataValidationType =
    typeRaw && (DV_TYPES as ReadonlyArray<string>).includes(typeRaw) ? (typeRaw as DataValidationType) : 'custom';
  const sqrefRaw = node.attrs['sqref'];
  if (!sqrefRaw) throw new OpenXmlSchemaError('worksheet: <dataValidation> missing @sqref');
  const opts: Partial<DataValidation> & { type: DataValidationType; sqref: ReturnType<typeof parseMultiCellRange> } = {
    type,
    sqref: parseMultiCellRange(sqrefRaw),
  };
  const operator = node.attrs['operator'];
  if (operator && (DV_OPERATORS as ReadonlyArray<string>).includes(operator)) {
    opts.operator = operator as DataValidationOperator;
  }
  const allowBlank = parseBoolXmlAttr(node.attrs['allowBlank']);
  if (allowBlank !== undefined) opts.allowBlank = allowBlank;
  const showInputMessage = parseBoolXmlAttr(node.attrs['showInputMessage']);
  if (showInputMessage !== undefined) opts.showInputMessage = showInputMessage;
  const showErrorMessage = parseBoolXmlAttr(node.attrs['showErrorMessage']);
  if (showErrorMessage !== undefined) opts.showErrorMessage = showErrorMessage;
  const showDropDown = parseBoolXmlAttr(node.attrs['showDropDown']);
  if (showDropDown !== undefined) opts.showDropDown = showDropDown;
  const errorTitle = node.attrs['errorTitle'];
  if (errorTitle !== undefined) opts.errorTitle = errorTitle;
  const error = node.attrs['error'];
  if (error !== undefined) opts.error = error;
  const errorStyle = node.attrs['errorStyle'];
  if (errorStyle && (DV_ERROR_STYLES as ReadonlyArray<string>).includes(errorStyle)) {
    opts.errorStyle = errorStyle as DataValidationErrorStyle;
  }
  const promptTitle = node.attrs['promptTitle'];
  if (promptTitle !== undefined) opts.promptTitle = promptTitle;
  const prompt = node.attrs['prompt'];
  if (prompt !== undefined) opts.prompt = prompt;
  const f1 = findChild(node, FORMULA1_TAG);
  if (f1?.text !== undefined) opts.formula1 = f1.text;
  const f2 = findChild(node, FORMULA2_TAG);
  if (f2?.text !== undefined) opts.formula2 = f2.text;
  return makeDataValidation(opts);
};

const CF_RULE_TYPES: ReadonlyArray<ConditionalFormattingRuleType> = [
  'expression',
  'cellIs',
  'colorScale',
  'dataBar',
  'iconSet',
  'top10',
  'aboveAverage',
  'uniqueValues',
  'duplicateValues',
  'containsText',
  'notContainsText',
  'beginsWith',
  'endsWith',
  'containsBlanks',
  'notContainsBlanks',
  'containsErrors',
  'notContainsErrors',
  'timePeriod',
];
const VISUAL_RULE_TYPES = new Set<ConditionalFormattingRuleType>(['colorScale', 'dataBar', 'iconSet']);

const parseConditionalFormatting = (node: XmlNode): ConditionalFormatting | undefined => {
  const sqrefRaw = node.attrs['sqref'];
  if (!sqrefRaw) return undefined;
  const rules: ConditionalFormattingRule[] = [];
  for (const r of findChildren(node, CF_RULE_TAG)) {
    const rule = parseCfRule(r);
    if (rule) rules.push(rule);
  }
  const opts: Parameters<typeof makeConditionalFormatting>[0] = {
    sqref: parseMultiCellRange(sqrefRaw),
    rules,
  };
  const pivot = parseBoolXmlAttr(node.attrs['pivot']);
  if (pivot !== undefined) opts.pivot = pivot;
  return makeConditionalFormatting(opts);
};

const parseCfRule = (node: XmlNode): ConditionalFormattingRule | undefined => {
  const typeRaw = node.attrs['type'];
  const priorityRaw = node.attrs['priority'];
  if (!typeRaw || priorityRaw === undefined) return undefined;
  if (!(CF_RULE_TYPES as ReadonlyArray<string>).includes(typeRaw)) return undefined;
  const priority = Number.parseInt(priorityRaw, 10);
  if (!Number.isInteger(priority)) return undefined;
  const type = typeRaw as ConditionalFormattingRuleType;
  const opts: Parameters<typeof makeCfRule>[0] = { type, priority };
  const dxf = parseIntegerAttr(node.attrs['dxfId']);
  if (dxf !== undefined) opts.dxfId = dxf;
  const stop = parseBoolXmlAttr(node.attrs['stopIfTrue']);
  if (stop !== undefined) opts.stopIfTrue = stop;
  if (node.attrs['operator']) opts.operator = node.attrs['operator'];
  if (node.attrs['text'] !== undefined) opts.text = node.attrs['text'];
  const percent = parseBoolXmlAttr(node.attrs['percent']);
  if (percent !== undefined) opts.percent = percent;
  const bottom = parseBoolXmlAttr(node.attrs['bottom']);
  if (bottom !== undefined) opts.bottom = bottom;
  const rank = parseIntegerAttr(node.attrs['rank']);
  if (rank !== undefined) opts.rank = rank;
  const aboveAverage = parseBoolXmlAttr(node.attrs['aboveAverage']);
  if (aboveAverage !== undefined) opts.aboveAverage = aboveAverage;
  const equalAverage = parseBoolXmlAttr(node.attrs['equalAverage']);
  if (equalAverage !== undefined) opts.equalAverage = equalAverage;
  const stdDev = parseIntegerAttr(node.attrs['stdDev']);
  if (stdDev !== undefined) opts.stdDev = stdDev;
  if (node.attrs['timePeriod']) opts.timePeriod = node.attrs['timePeriod'] as TimePeriod;

  const formulas: string[] = [];
  for (const f of findChildren(node, FORMULA_TAG)) formulas.push(f.text ?? '');
  if (formulas.length > 0) opts.formulas = formulas;

  if (VISUAL_RULE_TYPES.has(type)) {
    // Round-trip every non-formula child verbatim. Re-serialise to bytes then
    // back to a string so the writer can emit the same markup.
    const inner: string[] = [];
    for (const child of node.children) {
      if (child.name === FORMULA_TAG) continue;
      inner.push(new TextDecoder().decode(serializeXml(child, { xmlDeclaration: false })));
    }
    if (inner.length > 0) opts.innerXml = inner.join('');
  }
  return makeCfRule(opts);
};

const parseAutoFilter = (node: XmlNode): AutoFilter | undefined => {
  const ref = node.attrs['ref'];
  if (!ref) return undefined;
  const filterColumns: FilterColumn[] = [];
  for (const fc of findChildren(node, FILTER_COLUMN_TAG)) {
    const colIdRaw = fc.attrs['colId'];
    const colId = colIdRaw !== undefined ? Number.parseInt(colIdRaw, 10) : -1;
    if (!Number.isInteger(colId) || colId < 0) continue;
    const filtersEl = findChild(fc, FILTERS_TAG);
    if (!filtersEl) continue;
    const values: string[] = [];
    for (const f of findChildren(filtersEl, FILTER_TAG)) {
      const v = f.attrs['val'];
      if (v !== undefined) values.push(v);
    }
    const blank = parseBoolXmlAttr(filtersEl.attrs['blank']);
    const fc2: FilterColumn = { kind: 'filters', colId, values };
    if (blank !== undefined) fc2.blank = blank;
    filterColumns.push(fc2);
  }
  return { ref, filterColumns };
};

const parseHyperlink = (node: XmlNode, rels: Relationships | undefined): Hyperlink => {
  const ref = node.attrs['ref'];
  if (!ref) throw new OpenXmlSchemaError('worksheet: <hyperlink> missing @ref');
  const ridAttr = node.attrs[`{${REL_NS}}id`];
  const opts: Partial<Hyperlink> = { ref };
  if (ridAttr) {
    opts.rId = ridAttr;
    const rel = rels ? findById(rels, ridAttr) : undefined;
    if (rel) opts.target = rel.target;
  }
  if (node.attrs['location']) opts.location = node.attrs['location'];
  if (node.attrs['tooltip']) opts.tooltip = node.attrs['tooltip'];
  if (node.attrs['display']) opts.display = node.attrs['display'];
  // Deliberately not `makeHyperlink`: that validates the target, and a file
  // whose rels already carry a broken one still has to load so it can be
  // inspected and repaired. The check belongs where the link is authored.
  return opts as Hyperlink;
};

const maybeRecordRowDimension = (ws: Worksheet, node: XmlNode, rowIdx: number): void => {
  const opts: Partial<RowDimension> = {};
  const ht = parseFloatAttr(node.attrs['ht']);
  if (ht !== undefined) opts.height = ht;
  const customHeight = parseBoolXmlAttr(node.attrs['customHeight']);
  if (customHeight !== undefined) opts.customHeight = customHeight;
  const hidden = parseBoolXmlAttr(node.attrs['hidden']);
  if (hidden !== undefined) opts.hidden = hidden;
  const outlineLevel = parseIntegerAttr(node.attrs['outlineLevel']);
  if (outlineLevel !== undefined) opts.outlineLevel = outlineLevel;
  const collapsed = parseBoolXmlAttr(node.attrs['collapsed']);
  if (collapsed !== undefined) opts.collapsed = collapsed;
  const style = parseIntegerAttr(node.attrs['s']);
  if (style !== undefined) opts.style = style;
  if (Object.keys(opts).length === 0) return;
  ws.rowDimensions.set(rowIdx, makeRowDimension(opts));
};
