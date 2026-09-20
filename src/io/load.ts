// Public `loadWorkbook` entry point.
//
// **This is the minimum-skeleton stage**: open zip → parse manifest → resolve
// workbook part path → parse the `<sheets>` list → for each sheet, allocate an
// empty Worksheet (title + sheetId + state). Reading the actual cell content /
// styles / sharedStrings / theme / docProps happens in the next iterations of
// the loop.
//
// The skeleton is enough to round-trip through openpyxl's `genuine/empty.xlsx`
// fixture (3 empty sheets) and to give the rest of phase 3 a stable scaffolding
// to layer onto.

import { normalizeStrictArchive } from './strict.js';
import { findUserShapesRId, parseChartXml } from '../chart/chart-xml.js';
import { isChartExBytes, parseChartExXml } from '../chart/cx/chartex-xml.js';
import { parseUserShapesXml } from '../chart/user-shapes-xml.js';
import { parseChartsheetXml } from '../chartsheet/chartsheet-xml.js';
import { collectRawRelIds, parseDrawingXml } from '../drawing/drawing-xml.js';
import { loadImage } from '../drawing/image.js';
import type { XlsxSource } from '../io/source.js';
import { corePropsFromBytes } from '../packaging/core.js';
import { customPropsFromBytes } from '../packaging/custom.js';
import { extendedPropsFromBytes } from '../packaging/extended.js';
import { manifestFromBytes } from '../packaging/manifest.js';
import {
  findById,
  findByType,
  indexRelsById,
  makeRelationships,
  type Relationship,
  type Relationships,
  relsFromBytes,
} from '../packaging/relationships.js';
import { parseStylesheetXml } from '../styles/stylesheet-reader.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import type { DefinedName } from '../workbook/defined-names.js';
import { makeDefinedName } from '../workbook/defined-names.js';
import { parseSharedStringsXml, type SharedStringsTable } from '../workbook/shared-strings.js';
import { createWorkbook, type SheetRef, type SheetState, type Workbook } from '../workbook/workbook.js';
import { parseCommentsXml } from '../worksheet/comments-xml.js';
import {
  type ContentLimits,
  makeContentBudget,
  resolveContentLimits,
  type ResolvedContentLimits,
} from '../worksheet/content-budget.js';
import { parseWorksheetXml } from '../worksheet/reader.js';
import { parseTableXml } from '../worksheet/table-xml.js';
import {
  ARC_APP,
  ARC_CONTENT_TYPES,
  ARC_CORE,
  ARC_CUSTOM,
  ARC_ROOT_RELS,
  ARC_SHARED_STRINGS,
  ARC_STYLE,
  ARC_THEME,
  ARC_WORKBOOK,
  MARKUP_COMPAT_NS,
  REL_NS,
  SHEET_MAIN_NS,
} from '../xml/namespaces.js';
import { type ParsedDocument, parseXmlDocument } from '../xml/parser.js';
import { assertNotStrictRelTypes, assertNotStrictRoot } from '../xml/strict-package.js';
import { findChild, findChildren, type XmlNode } from '../xml/tree.js';
import type { DecompressionLimits } from '../zip/decompression-guard.js';
import { openZip, type ZipArchive } from '../zip/reader.js';

/**
 * Options for {@link loadWorkbook}. Earlier drafts exposed `readOnly` /
 * `keepLinks` / `keepVba` / `dataOnly` / `richText` placeholders that the
 * loader silently ignored; those were removed to keep the surface honest.
 * Future toggles land here once their behavior is implemented.
 */
export interface LoadOptions {
  /**
   * Decompression-bomb safeguards applied while inflating zip entries.
   * Defaults to limits that fit any legitimate xlsx; pass `false` to disable
   * (only safe for fully trusted sources). See
   * {@link DecompressionLimits} for the individual knobs.
   */
  decompressionLimits?: DecompressionLimits | false;
  /**
   * Caps on how much content the load will model. `decompressionLimits` bounds
   * the bytes an archive inflates to; this bounds the cells and rows those
   * bytes turn into, which is what the load spends its time on and what the
   * Workbook it returns is made of. A small upload can inflate to a
   * `<sheetData>` of a few hundred MB without tripping any byte limit, and the
   * model was built for every cell of it before the caller got a chance to look
   * at anything.
   *
   * This does not bound peak memory by itself. `loadWorkbook` inflates a
   * worksheet part and decodes it to a string before the first cell of it is
   * charged, so those bytes stay bounded by `decompressionLimits` alone; what
   * the cap removes is the Workbook built on top of them, which outlives the
   * part and is the larger of the two.
   *
   * Unlimited by default. Both counts cover one pass over the content: every
   * worksheet of the workbook here, and one row-iteration in
   * `loadWorkbookStream`, which materialises a row at a time and can be
   * iterated more than once.
   *
   * A service that accepts spreadsheets from strangers wants a ceiling here as
   * well as on `decompressionLimits`. Cost is close to linear in cells, so size
   * the cap from the heap the process can spare and halve it to halve the worst
   * case. As a starting profile for ordinary business workbooks:
   *
   * ```ts
   * const wb = await loadWorkbook(source, {
   *   decompressionLimits: { maxTotalUncompressedBytes: 256 * 1024 * 1024 },
   *   contentLimits: { maxCells: 1_000_000, maxRows: 100_000 },
   * });
   * ```
   *
   * Exceeding either cap throws an `OpenXmlContentLimitError`, which names the
   * cap and the cell or row that reached it. See `SECURITY.md` for the threat
   * model both options belong to.
   */
  contentLimits?: ContentLimits;
}

/** Office Document relationship type: the package-root pointer to `xl/workbook.xml`. */
export const OFFICE_DOC_REL_TYPE = `${REL_NS}/officeDocument`;

/** Root element of `xl/workbook.xml`. Shared with the streaming reader. */
export const WORKBOOK_TAG = `{${SHEET_MAIN_NS}}workbook`;

/**
 * Resolve an OPC relationship target against its source part path.
 *
 * - Targets starting with `/` are package-absolute.
 * - Otherwise the target is relative to the source part's parent directory.
 * - `..` segments collapse normally.
 */
export function resolveRelTarget(sourcePartPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const lastSlash = sourcePartPath.lastIndexOf('/');
  const parentDir = lastSlash >= 0 ? sourcePartPath.slice(0, lastSlash + 1) : '';
  const joined = parentDir + target;
  return normalizePath(joined);
}

/**
 * Some producers store literal spaces in ZIP entry names. Decode only as a
 * fallback after looking for the exact package name, so encoded names retain
 * their identity and existing relationship resolution stays compatible.
 */
function decodeTarget(target: string): string {
  if (!target.includes('%')) return target;
  try {
    return decodeURIComponent(target);
  } catch {
    return target;
  }
}

function normalizePath(path: string): string {
  const segments = path.split('/');
  const out: string[] = [];
  for (const seg of segments) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') {
      // Pop the last accumulated segment when one exists; when `out` is
      // empty (a relative target with more `..` than ancestors) the pop is
      // a no-op so the climb is silently absorbed at the package root. The
      // archive.has() check on the resolved path catches any escape attempt
      // because the entry simply won't exist outside the package — there's
      // no filesystem traversal to worry about, only a missing-entry error.
      out.pop();
      continue;
    }
    out.push(seg);
  }
  return out.join('/');
}

/** A workbook-level part a package carries at most one of. */
interface OptionalWorkbookPart {
  /** How the part is named in error messages. */
  name: string;
  /** Relationship type that binds the part to the workbook. */
  relType: string;
  /** Where Excel writes the part. A convention, not a rule. */
  conventionalPath: string;
}

export const SHARED_STRINGS_PART: OptionalWorkbookPart = {
  name: 'sharedStrings',
  relType: `${REL_NS}/sharedStrings`,
  conventionalPath: ARC_SHARED_STRINGS,
};

export const STYLES_PART: OptionalWorkbookPart = {
  name: 'styles',
  relType: `${REL_NS}/styles`,
  conventionalPath: ARC_STYLE,
};

const THEME_PART: OptionalWorkbookPart = {
  name: 'theme',
  relType: `${REL_NS}/theme`,
  conventionalPath: ARC_THEME,
};

/**
 * Read an optional workbook-level part, following the relationship that binds
 * it to the workbook.
 *
 * The relationship decides. `xl/sharedStrings.xml` and its siblings are only
 * where Excel happens to write these parts, and a producer may point the rel
 * at any part in the package, so the conventional path serves the packages
 * that ship the part with no relationship at all.
 *
 * A relationship that resolves to nothing is a malformed package rather than
 * an absent part, so it throws instead of reading back as an empty table.
 */
export function readOptionalWorkbookPart(
  archive: ZipArchive,
  workbookPath: string,
  wbRels: Relationships,
  part: OptionalWorkbookPart,
): Uint8Array | undefined {
  const rel = findByType(wbRels, part.relType);
  if (!rel) {
    return archive.has(part.conventionalPath) ? archive.read(part.conventionalPath) : undefined;
  }
  if (rel.targetMode === 'External') {
    throw new OpenXmlSchemaError(
      `workbook rels: the ${part.name} relationship is external ("${rel.target}"), and has to name a part inside the package`,
    );
  }
  let path = resolveRelTarget(workbookPath, rel.target);
  if (!archive.has(path)) path = resolveRelTarget(workbookPath, decodeTarget(rel.target));
  if (!archive.has(path)) {
    throw new OpenXmlSchemaError(
      `workbook rels: the ${part.name} relationship targets "${path}", which the package does not contain`,
    );
  }
  return archive.read(path);
}

/** Sibling rels-part path for a given part. `xl/workbook.xml` → `xl/_rels/workbook.xml.rels`. */
function relsPathFor(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  if (i < 0) return `_rels/${partPath}.rels`;
  return `${partPath.slice(0, i)}/_rels/${partPath.slice(i + 1)}.rels`;
}

interface SheetEntry {
  /** Display name (`sheet/@name`). */
  name: string;
  /** Workbook-scope sheetId (`sheet/@sheetId`). */
  sheetId: number;
  /** Workbook rels Id (`sheet/@r:id`). */
  rId: string;
  /** Visibility state — defaults to `'visible'` when the attribute is absent. */
  state: SheetState;
}

const SHEET_TAG = `{${SHEET_MAIN_NS}}sheet`;
const SHEETS_TAG = `{${SHEET_MAIN_NS}}sheets`;
const DEFINED_NAMES_TAG = `{${SHEET_MAIN_NS}}definedNames`;
const DEFINED_NAME_TAG = `{${SHEET_MAIN_NS}}definedName`;
const RID_ATTR = `{${REL_NS}}id`;

/** Extract the `<definedNames>/<definedName>` entries from a parsed `xl/workbook.xml`. */
function parseDefinedNames(workbookRoot: XmlNode): DefinedName[] {
  const wrapper = findChild(workbookRoot, DEFINED_NAMES_TAG);
  if (!wrapper) return [];
  const out: DefinedName[] = [];
  for (const node of findChildren(wrapper, DEFINED_NAME_TAG)) {
    const name = node.attrs['name'];
    if (!name) throw new OpenXmlSchemaError("workbook.xml: <definedName> is missing 'name'");
    const value = node.text ?? '';
    const opts: Partial<DefinedName> & { name: string; value: string } = { name, value };
    const scopeAttr = node.attrs['localSheetId'];
    if (scopeAttr !== undefined) {
      const scope = Number.parseInt(scopeAttr, 10);
      if (Number.isInteger(scope) && scope >= 0) opts.scope = scope;
    }
    if (node.attrs['hidden'] === '1' || node.attrs['hidden'] === 'true') opts.hidden = true;
    if (node.attrs['comment'] !== undefined) opts.comment = node.attrs['comment'];
    out.push(makeDefinedName(opts));
  }
  return out;
}

const WORKBOOK_PR_TAG = `{${SHEET_MAIN_NS}}workbookPr`;
const WORKBOOK_PROTECTION_TAG = `{${SHEET_MAIN_NS}}workbookProtection`;
const BOOK_VIEWS_TAG = `{${SHEET_MAIN_NS}}bookViews`;
const WORKBOOK_VIEW_TAG = `{${SHEET_MAIN_NS}}workbookView`;
const CUSTOM_WORKBOOK_VIEWS_TAG = `{${SHEET_MAIN_NS}}customWorkbookViews`;
const CUSTOM_WORKBOOK_VIEW_TAG = `{${SHEET_MAIN_NS}}customWorkbookView`;
const CALC_PR_TAG = `{${SHEET_MAIN_NS}}calcPr`;
const FILE_VERSION_TAG = `{${SHEET_MAIN_NS}}fileVersion`;
const FILE_SHARING_TAG = `{${SHEET_MAIN_NS}}fileSharing`;
const OLE_SIZE_TAG = `{${SHEET_MAIN_NS}}oleSize`;
const FILE_RECOVERY_PR_TAG = `{${SHEET_MAIN_NS}}fileRecoveryPr`;
const PIVOT_CACHES_TAG = `{${SHEET_MAIN_NS}}pivotCaches`;
const PIVOT_CACHE_TAG = `{${SHEET_MAIN_NS}}pivotCache`;
const EXTERNAL_REFERENCES_TAG = `{${SHEET_MAIN_NS}}externalReferences`;
const EXTERNAL_REFERENCE_TAG = `{${SHEET_MAIN_NS}}externalReference`;
const SMART_TAG_PR_TAG = `{${SHEET_MAIN_NS}}smartTagPr`;
const SMART_TAG_TYPES_TAG = `{${SHEET_MAIN_NS}}smartTagTypes`;
const SMART_TAG_TYPE_TAG = `{${SHEET_MAIN_NS}}smartTagType`;
const FUNCTION_GROUPS_TAG = `{${SHEET_MAIN_NS}}functionGroups`;
const FUNCTION_GROUP_TAG = `{${SHEET_MAIN_NS}}functionGroup`;

/**
 * Parse the `<workbookPr date1904>` flag. Mac-origin workbooks set
 * `date1904="true"`; everything else uses the Windows 1900 epoch. The value
 * drives Date serialization and lets readers select the epoch when converting
 * numeric date serials.
 */
export function parseDate1904(workbookRoot: XmlNode): boolean {
  const pr = findChild(workbookRoot, WORKBOOK_PR_TAG);
  if (!pr) return false;
  const v = pr.attrs['date1904'];
  return v === '1' || v === 'true';
}

/** Extract the `<sheets>/<sheet>` entries from a parsed `xl/workbook.xml`. */
export function parseSheetEntries(workbookRoot: XmlNode): SheetEntry[] {
  const sheets = findChild(workbookRoot, SHEETS_TAG);
  if (!sheets) return [];
  const out: SheetEntry[] = [];
  for (const node of findChildren(sheets, SHEET_TAG)) {
    const name = node.attrs['name'];
    const sheetIdAttr = node.attrs['sheetId'];
    const rId = node.attrs[RID_ATTR];
    if (!name) throw new OpenXmlSchemaError("workbook.xml: <sheet> is missing 'name'");
    if (!sheetIdAttr) throw new OpenXmlSchemaError(`workbook.xml: <sheet name="${name}"> is missing 'sheetId'`);
    if (!rId) throw new OpenXmlSchemaError(`workbook.xml: <sheet name="${name}"> is missing 'r:id'`);
    const sheetId = Number.parseInt(sheetIdAttr, 10);
    if (!Number.isInteger(sheetId) || sheetId < 1) {
      throw new OpenXmlSchemaError(
        `workbook.xml: <sheet name="${name}"> sheetId "${sheetIdAttr}" is not a positive integer`,
      );
    }
    const stateAttr = node.attrs['state'];
    let state: SheetState = 'visible';
    if (stateAttr === 'hidden' || stateAttr === 'veryHidden') state = stateAttr;
    out.push({ name, sheetId, rId, state });
  }
  return out;
}

/**
 * Load a workbook from any {@link XlsxSource}. Currently produces a scaffold
 * Workbook: each Worksheet is empty (no cells / styles / shared strings / theme
 * yet). The next phase-3 iterations layer those in atop the same skeleton.
 */
export async function loadWorkbook(source: XlsxSource, opts: LoadOptions = {}): Promise<Workbook> {
  // Settled before the source is opened, so a cap that cannot mean anything is
  // reported where the caller passed it rather than partway through a read.
  const contentLimits = resolveContentLimits(opts.contentLimits);
  const archive = await openZip(
    source,
    opts.decompressionLimits === undefined ? {} : { decompressionLimits: opts.decompressionLimits },
  );
  try {
    return loadWorkbookFromArchive(normalizeStrictArchive(archive), contentLimits);
  } finally {
    archive.close();
  }
}

/** Internal: same as {@link loadWorkbook} but operating on an already-opened archive. */
function loadWorkbookFromArchive(archive: ZipArchive, contentLimits: ResolvedContentLimits): Workbook {
  // One budget for the whole load: a per-sheet cap would let a workbook of a
  // thousand small sheets through a ceiling meant to bound the whole read.
  const contentBudget = makeContentBudget(contentLimits);
  // 1. Manifest — resolves which override entries the package declares.
  if (!archive.has(ARC_CONTENT_TYPES)) {
    throw new OpenXmlSchemaError(`loadWorkbook: missing "${ARC_CONTENT_TYPES}"`);
  }
  const manifest = manifestFromBytes(archive.read(ARC_CONTENT_TYPES));

  // 2. Root rels → resolve the office-document relationship to the workbook part path.
  if (!archive.has(ARC_ROOT_RELS)) {
    throw new OpenXmlSchemaError(`loadWorkbook: missing "${ARC_ROOT_RELS}"`);
  }
  const rootRels = relsFromBytes(archive.read(ARC_ROOT_RELS));
  const officeRel = rootRels.rels.find((r) => r.type === OFFICE_DOC_REL_TYPE);
  if (!officeRel) {
    assertNotStrictRelTypes(rootRels.rels.map((r) => r.type));
    throw new OpenXmlSchemaError('loadWorkbook: root rels missing officeDocument relationship');
  }
  const workbookPath = resolveRelTarget('', officeRel.target);
  if (workbookPath !== ARC_WORKBOOK) {
    // Most xlsx files put the workbook at xl/workbook.xml. We accept any path
    // the rels point at as long as the archive holds it.
    if (!archive.has(workbookPath)) {
      throw new OpenXmlSchemaError(`loadWorkbook: workbook part "${workbookPath}" not found in archive`);
    }
  }

  // 3. workbook.xml — parse to extract sheet metadata only.
  const wbDoc = parseXmlDocument(archive.read(workbookPath));
  const wbRoot = wbDoc.root;
  // Compared with its namespace, not on local name alone: every `<sheets>`
  // lookup below is a transitional QName, so a root in some other namespace
  // would read as a workbook with no sheets rather than be rejected.
  if (wbRoot.name !== WORKBOOK_TAG) {
    assertNotStrictRoot(wbRoot.name);
    throw new OpenXmlSchemaError(`loadWorkbook: ${workbookPath} root is "${wbRoot.name}", expected workbook`);
  }
  const sheetEntries = parseSheetEntries(wbRoot);
  const definedNamesFromXml = parseDefinedNames(wbRoot);

  // 4. workbook.xml.rels — needed to resolve each sheet's rId to a part path.
  const wbRelsPath = relsPathFor(workbookPath);
  // openpyxl tolerates a missing workbook.xml.rels (it implies no sheets); in
  // practice every Excel file has one. We require it so a malformed package
  // surfaces as an OpenXmlSchemaError, not a silently-empty workbook.
  if (sheetEntries.length > 0 && !archive.has(wbRelsPath)) {
    throw new OpenXmlSchemaError(`loadWorkbook: workbook has sheets but rels part "${wbRelsPath}" is missing`);
  }
  const wbRels = archive.has(wbRelsPath) ? relsFromBytes(archive.read(wbRelsPath)) : makeRelationships();

  // 4b. sharedStrings.xml: optional, wherever the workbook rels point.
  const sstBytes = readOptionalWorkbookPart(archive, workbookPath, wbRels, SHARED_STRINGS_PART);
  const sharedStrings: SharedStringsTable | undefined =
    sstBytes === undefined ? undefined : parseSharedStringsXml(sstBytes);
  // Entries reach the worksheet reader as-is: a rich-text `<si>` becomes a
  // rich-text cell value, so the per-run formatting Excel stored survives the
  // round-trip instead of collapsing into the concatenated text body.
  const sst = sharedStrings?.entries ?? [];

  // 4c. styles.xml: optional. Same lookup as sst.
  const stylesBytes = readOptionalWorkbookPart(archive, workbookPath, wbRels, STYLES_PART);
  const styles: ReturnType<typeof parseStylesheetXml> | undefined =
    stylesBytes === undefined ? undefined : parseStylesheetXml(stylesBytes);

  // 4d. docProps/{core,app,custom}.xml — package-level metadata. Each part is
  // optional; absent ones leave the matching Workbook field undefined. We walk
  // both the canonical path and the root rels so non-default layouts (rare but
  // legal) still resolve.
  const properties = archive.has(ARC_CORE) ? corePropsFromBytes(archive.read(ARC_CORE)) : undefined;
  const appProperties = archive.has(ARC_APP) ? extendedPropsFromBytes(archive.read(ARC_APP)) : undefined;
  const customProperties = archive.has(ARC_CUSTOM) ? customPropsFromBytes(archive.read(ARC_CUSTOM)) : undefined;

  // 4e. xl/theme/theme1.xml — kept verbatim. Excel renders with this exact
  // payload; round-tripping the bytes avoids drift.
  const themeXml = readOptionalWorkbookPart(archive, workbookPath, wbRels, THEME_PART);

  // 5. Build the Workbook. We bypass `addWorksheet` because that allocates
  // sheetIds via `allocateSheetId`; load preserves the IDs from XML.
  const wb = createWorkbook({ date1904: parseDate1904(wbRoot) });
  if (styles) wb.styles = styles;
  if (properties) wb.properties = properties;
  if (appProperties) wb.appProperties = appProperties;
  if (customProperties) wb.customProperties = customProperties;
  if (themeXml) wb.themeXml = themeXml;
  if (definedNamesFromXml.length > 0) wb.definedNames = definedNamesFromXml;
  const seenTitles = new Set<string>();
  const passthroughRoots: string[] = [];
  const vml = makeVmlPartCache(archive);
  for (const entry of sheetEntries) {
    if (seenTitles.has(entry.name)) {
      throw new OpenXmlSchemaError(`loadWorkbook: duplicate sheet name "${entry.name}"`);
    }
    seenTitles.add(entry.name);
    const rel = findById(wbRels, entry.rId);
    if (!rel) {
      throw new OpenXmlSchemaError(`loadWorkbook: sheet "${entry.name}" rId "${entry.rId}" has no matching rels entry`);
    }
    const sheetPath = resolveRelTarget(workbookPath, rel.target);
    if (!archive.has(sheetPath)) {
      throw new OpenXmlSchemaError(`loadWorkbook: sheet part "${sheetPath}" not found in archive`);
    }
    const sheetRelsPath = relsPathFor(sheetPath);
    const sheetRels = archive.has(sheetRelsPath) ? relsFromBytes(archive.read(sheetRelsPath)) : undefined;
    // Build an id → rel index once; the loadTable / loadComments / loadDrawing
    // callbacks are invoked once per cross-reference in the worksheet body and
    // a linear find() inside each would be O(rels × refs). Sheets with many
    // hyperlinks / tables / drawings (e.g. dashboard workbooks) make that
    // accumulation visible.
    const sheetRelsById = sheetRels ? indexRelsById(sheetRels) : undefined;
    const loadTable = sheetRelsById
      ? (relId: string) => {
          const tRel = sheetRelsById.get(relId);
          if (!tRel) return undefined;
          const tablePath = resolveRelTarget(sheetPath, tRel.target);
          if (!archive.has(tablePath)) return undefined;
          return parseTableXml(archive.read(tablePath));
        }
      : undefined;
    const loadComments = sheetRelsById
      ? (relId: string) => {
          const cRel = sheetRelsById.get(relId);
          if (!cRel) return undefined;
          const cPath = resolveRelTarget(sheetPath, cRel.target);
          if (!archive.has(cPath)) return undefined;
          return parseCommentsXml(archive.read(cPath));
        }
      : undefined;
    const loadDrawing = sheetRelsById
      ? (relId: string) => {
          const dRel = sheetRelsById.get(relId);
          if (!dRel) return undefined;
          const dPath = resolveRelTarget(sheetPath, dRel.target);
          if (!archive.has(dPath)) return undefined;
          const drawing = parseDrawingXml(archive.read(dPath));
          // Phase-2: resolve drawing-rels to populate chart payloads.
          const dRelsPath = relsPathFor(dPath);
          if (archive.has(dRelsPath)) {
            const dRels = relsFromBytes(archive.read(dRelsPath));
            const dRelsById = indexRelsById(dRels);
            for (const item of drawing.items) {
              if (item.content.kind === 'chart') {
                const chartRId = item.content.chart.rId;
                if (!chartRId) continue;
                const chartRel = dRelsById.get(chartRId);
                if (!chartRel) continue;
                const chartPath = resolveRelTarget(dPath, chartRel.target);
                if (archive.has(chartPath)) {
                  const chartBytes = archive.read(chartPath);
                  if (isChartExBytes(chartBytes)) {
                    item.content.chart.cxSpace = parseChartExXml(chartBytes);
                  } else {
                    const space = parseChartXml(chartBytes);
                    // Resolve <c:userShapes r:id="..."> via the chart's own
                    // rels file (xl/charts/_rels/chartN.xml.rels).
                    const userShapesRId = findUserShapesRId(chartBytes);
                    if (userShapesRId) {
                      const chartRelsPath = relsPathFor(chartPath);
                      if (archive.has(chartRelsPath)) {
                        const chartRelsObj = relsFromBytes(archive.read(chartRelsPath));
                        const usRel = indexRelsById(chartRelsObj).get(userShapesRId);
                        if (usRel) {
                          const usPath = resolveRelTarget(chartPath, usRel.target);
                          if (archive.has(usPath)) {
                            try {
                              space.userShapes = parseUserShapesXml(archive.read(usPath));
                            } catch {
                              // Tolerate parse failures (Excel sometimes emits
                              // chartDrawing parts with namespaces / shapes
                              // outside our model).
                            }
                          }
                        }
                      }
                    }
                    item.content.chart.space = space;
                  }
                }
              } else if (item.content.kind === 'picture') {
                const picRId = item.content.picture.rId;
                if (!picRId) continue;
                const picRel = dRelsById.get(picRId);
                if (!picRel) continue;
                const imgPath = resolveRelTarget(dPath, picRel.target);
                if (archive.has(imgPath)) {
                  try {
                    item.content.picture.image = loadImage(archive.read(imgPath));
                  } catch {
                    // Unknown / unsupported format — leave bytes-less; callers
                    // can read via the rId + archive directly if they need the
                    // raw payload.
                  }
                }
              }
            }
            // Anything kept verbatim (shapes, groups, mc:AlternateContent
            // wrappers) may still reference the drawing's rels. Carry those
            // entries under their original ids and hand their targets to the
            // passthrough sweep so the references still resolve on re-save.
            const rawRelIds = collectRawRelIds(drawing);
            if (rawRelIds.size > 0) {
              const carried = makeRelationships();
              for (const id of rawRelIds) {
                const rawRel = dRelsById.get(id);
                if (!rawRel) continue;
                carried.rels.push(rawRel);
                if (rawRel.targetMode !== 'External') {
                  passthroughRoots.push(resolveRelTarget(dPath, rawRel.target));
                }
              }
              if (carried.rels.length > 0) drawing.rawRels = carried;
            }
          }
          return drawing;
        }
      : undefined;
    // Distinguish worksheet vs chartsheet by inspecting the workbook-rels
    // entry's relationship type.
    const isChartsheet = rel.type === `${REL_NS}/chartsheet`;
    if (isChartsheet) {
      const chartsheet = parseChartsheetXml(archive.read(sheetPath), entry.name);
      // Inline drawing reference from the chartsheet XML.
      if (sheetRels) {
        // Find the drawing rel and resolve it the same way worksheets do.
        const drawingRel = sheetRels.rels.find((r) => r.type === `${REL_NS}/drawing`);
        if (drawingRel && loadDrawing) {
          const d = loadDrawing(drawingRel.id);
          if (d) chartsheet.drawing = d;
        }
      }
      const ref: SheetRef = {
        kind: 'chartsheet',
        sheet: chartsheet,
        sheetId: entry.sheetId,
        state: entry.state,
        rId: entry.rId,
      };
      wb.sheets.push(ref);
      continue;
    }
    const ws = parseWorksheetXml(archive.read(sheetPath), entry.name, {
      sharedStrings: sst,
      contentBudget,
      ...(sheetRels ? { rels: sheetRels } : {}),
      ...(loadTable ? { loadTable } : {}),
      ...(loadComments ? { loadComments } : {}),
      ...(loadDrawing ? { loadDrawing } : {}),
    });
    if (sheetRels) {
      const extras = captureSheetRelsExtras(sheetRels, sheetPath, vml);
      if (extras.length > 0) {
        ws.relsExtras = extras;
        // The writer re-emits these rels verbatim, so whatever they point at
        // has to be carried over as passthrough or the output dangles.
        for (const e of extras) {
          if (e.targetMode !== 'External') passthroughRoots.push(resolveRelTarget(sheetPath, e.target));
        }
      }
    }
    // A <legacyDrawing> pointing at the comment VML (or at nothing) has no
    // surviving rel; the writer regenerates the comment link from
    // ws.legacyComments, so re-emitting this rId would dangle.
    if (ws.legacyDrawingRId !== undefined && !ws.relsExtras?.some((e) => e.id === ws.legacyDrawingRId)) {
      delete ws.legacyDrawingRId;
    }
    const ref: SheetRef = {
      kind: 'worksheet',
      sheet: ws,
      sheetId: entry.sheetId,
      state: entry.state,
      rId: entry.rId,
    };
    wb.sheets.push(ref);
  }

  captureWorkbookXmlExtras(wbRoot, wb);
  captureWorkbookXmlRoot(wbDoc, wb);
  captureWorkbookRelsExtras(wbRels, wb);

  // Pass-through: capture parts we don't model (VBA / pivot / activeX / OLE /
  // customUI / customXml / etc.) so re-saving doesn't drop them.
  capturePassthrough(archive, manifest, wb, passthroughRoots, vml);
  return wb;
}

const SHEET_MODELED_REL_TYPES: ReadonlySet<string> = new Set([
  `${REL_NS}/hyperlink`,
  `${REL_NS}/table`,
  `${REL_NS}/comments`,
  `${REL_NS}/drawing`,
]);
const VML_DRAWING_REL = `${REL_NS}/vmlDrawing`;

/**
 * Capture per-sheet rels entries that don't match a modeled type. The writer
 * re-emits these verbatim alongside the freshly allocated modeled rels so
 * captured passthrough parts (pivotTable / queryTable / slicer /
 * printerSettings / oleObject / customProperty / threadedComment) remain
 * reachable from the worksheet after a round-trip.
 *
 * `vmlDrawing` rels are split by part content rather than by type: only the
 * comments VML is regenerated (from `ws.legacyComments`), so its rel is
 * dropped; header/footer and form-control VML is re-emitted by its original
 * r:id and rides along as an extra.
 */
function captureSheetRelsExtras(
  sheetRels: import('../packaging/relationships.js').Relationships,
  sheetPath: string,
  vml: VmlPartCache,
): Relationship[] {
  const extras: Relationship[] = [];
  for (const rel of sheetRels.rels) {
    if (SHEET_MODELED_REL_TYPES.has(rel.type)) continue;
    if (rel.type === VML_DRAWING_REL && vml.isCommentVml(resolveRelTarget(sheetPath, rel.target))) continue;
    extras.push(rel);
  }
  return extras;
}

/**
 * Walk top-level children of `<workbook>` and split anything that isn't
 * `<sheets>` or `<definedNames>` into the before/after halves the writer
 * inserts around the modeled elements. Order is preserved within each half so
 * things like `<fileVersion>`, `<workbookPr>`, `<bookViews>`, `<calcPr>`,
 * `<pivotCaches>`, `<extLst>` round-trip in document order.
 */
/**
 * Keep the `<workbook>` root's namespace declarations and `mc:Ignorable`. Excel
 * writes `mc:Ignorable="x15 xr xr6 xr10 xr2"` naming prefixes declared only
 * there; dropping either half — or renaming the prefixes — makes Excel report
 * the file as corrupt. The default namespace and `r` are re-emitted by the
 * writer itself, so they are not carried here.
 */
function captureWorkbookXmlRoot(doc: ParsedDocument, wb: Workbook): void {
  const namespaces = doc.rootNamespaces.filter((d) => d.prefix !== '' && d.prefix !== 'r');
  const ignorable = doc.root.attrs[`{${MARKUP_COMPAT_NS}}Ignorable`];
  if (namespaces.length === 0 && ignorable === undefined) return;
  wb.workbookXmlRoot = {
    namespaces,
    ...(ignorable !== undefined ? { ignorable } : {}),
  };
}

function captureWorkbookXmlExtras(wbRoot: XmlNode, wb: Workbook): void {
  const beforeSheets: XmlNode[] = [];
  const afterSheets: XmlNode[] = [];
  let seenSheets = false;
  for (const child of wbRoot.children) {
    if (child.name === SHEETS_TAG) {
      seenSheets = true;
      continue;
    }
    if (child.name === DEFINED_NAMES_TAG) continue;
    // Lift <workbookProtection> into the typed workbook field instead of
    // stashing it as a passthrough XmlNode (B5 partial).
    if (child.name === WORKBOOK_PROTECTION_TAG) {
      wb.workbookProtection = parseWorkbookProtection(child);
      continue;
    }
    // Lift <workbookPr> into the typed workbook field. The date1904 attribute
    // is already mirrored onto wb.date1904; everything else stops leaking into
    // bodyExtras.
    if (child.name === WORKBOOK_PR_TAG) {
      const wp = parseWorkbookProperties(child);
      if (wp) wb.workbookProperties = wp;
      continue;
    }
    // Lift <fileSharing> into the typed workbook field.
    if (child.name === FILE_SHARING_TAG) {
      const fs: import('../workbook/file-sharing.js').FileSharing = {};
      const a = child.attrs;
      const flag = (raw: string | undefined): boolean | undefined => {
        if (raw === '1' || raw === 'true') return true;
        if (raw === '0' || raw === 'false') return false;
        return undefined;
      };
      const ror = flag(a['readOnlyRecommended']);
      if (ror !== undefined) fs.readOnlyRecommended = ror;
      if (a['userName'] !== undefined) fs.userName = a['userName'];
      if (a['reservationPassword'] !== undefined) fs.reservationPassword = a['reservationPassword'];
      if (a['algorithmName'] !== undefined) fs.algorithmName = a['algorithmName'];
      if (a['hashValue'] !== undefined) fs.hashValue = a['hashValue'];
      if (a['saltValue'] !== undefined) fs.saltValue = a['saltValue'];
      if (a['spinCount'] !== undefined) {
        const n = Number.parseInt(a['spinCount'], 10);
        if (Number.isInteger(n)) fs.spinCount = n;
      }
      if (Object.keys(fs).length > 0) wb.fileSharing = fs;
      continue;
    }
    // Lift <fileVersion> into the typed workbook field.
    if (child.name === FILE_VERSION_TAG) {
      const fv: import('../workbook/file-version.js').FileVersion = {};
      if (child.attrs['appName'] !== undefined) fv.appName = child.attrs['appName'];
      if (child.attrs['lastEdited'] !== undefined) fv.lastEdited = child.attrs['lastEdited'];
      if (child.attrs['lowestEdited'] !== undefined) fv.lowestEdited = child.attrs['lowestEdited'];
      if (child.attrs['rupBuild'] !== undefined) fv.rupBuild = child.attrs['rupBuild'];
      if (child.attrs['codeName'] !== undefined) fv.codeName = child.attrs['codeName'];
      if (Object.keys(fv).length > 0) wb.fileVersion = fv;
      continue;
    }
    // Lift <bookViews> into the typed workbook field.
    if (child.name === BOOK_VIEWS_TAG) {
      const views: import('../workbook/views.js').WorkbookView[] = [];
      for (const v of findChildren(child, WORKBOOK_VIEW_TAG)) views.push(parseWorkbookView(v));
      if (views.length > 0) wb.bookViews = views;
      continue;
    }
    // Lift <customWorkbookViews> into the typed workbook field.
    if (child.name === CUSTOM_WORKBOOK_VIEWS_TAG) {
      const cws: import('../workbook/views.js').CustomWorkbookView[] = [];
      for (const v of findChildren(child, CUSTOM_WORKBOOK_VIEW_TAG)) {
        const parsed = parseCustomWorkbookView(v);
        if (parsed) cws.push(parsed);
      }
      if (cws.length > 0) wb.customWorkbookViews = cws;
      continue;
    }
    // Lift <calcPr> into the typed workbook field.
    if (child.name === CALC_PR_TAG) {
      const cp = parseCalcProperties(child);
      if (cp) wb.calcProperties = cp;
      continue;
    }
    // Lift <oleSize ref="…"/> as a single typed string attribute.
    if (child.name === OLE_SIZE_TAG) {
      const ref = child.attrs['ref'];
      if (ref) wb.oleSize = ref;
      continue;
    }
    // Lift <smartTagPr embed="1" show="all"/>.
    if (child.name === SMART_TAG_PR_TAG) {
      const out: import('../workbook/smart-tags.js').SmartTagProperties = {};
      const a = child.attrs;
      if (a['embed'] === '1' || a['embed'] === 'true') out.embed = true;
      else if (a['embed'] === '0' || a['embed'] === 'false') out.embed = false;
      if (a['show'] === 'all' || a['show'] === 'noIndicator') out.show = a['show'];
      if (Object.keys(out).length > 0) wb.smartTagPr = out;
      continue;
    }
    // Lift <smartTagTypes><smartTagType .../></smartTagTypes>.
    if (child.name === SMART_TAG_TYPES_TAG) {
      const tags: import('../workbook/smart-tags.js').SmartTagType[] = [];
      for (const t of findChildren(child, SMART_TAG_TYPE_TAG)) {
        const entry: import('../workbook/smart-tags.js').SmartTagType = {};
        if (t.attrs['namespaceUri'] !== undefined) entry.namespaceUri = t.attrs['namespaceUri'];
        if (t.attrs['name'] !== undefined) entry.name = t.attrs['name'];
        if (t.attrs['url'] !== undefined) entry.url = t.attrs['url'];
        if (Object.keys(entry).length > 0) tags.push(entry);
      }
      if (tags.length > 0) wb.smartTagTypes = tags;
      continue;
    }
    // Lift <functionGroups builtInGroupCount=…><functionGroup
    // name=…/></functionGroups>.
    if (child.name === FUNCTION_GROUPS_TAG) {
      const fg: import('../workbook/function-groups.js').FunctionGroups = { groups: [] };
      const bicgRaw = child.attrs['builtInGroupCount'];
      if (bicgRaw !== undefined) {
        const n = Number.parseInt(bicgRaw, 10);
        if (Number.isInteger(n)) fg.builtInGroupCount = n;
      }
      for (const g of findChildren(child, FUNCTION_GROUP_TAG)) {
        const name = g.attrs['name'];
        if (name) fg.groups.push({ name });
      }
      if (fg.groups.length > 0 || fg.builtInGroupCount !== undefined) wb.functionGroups = fg;
      continue;
    }
    // Lift <externalReferences><externalReference
    // r:id=…/></externalReferences>.
    if (child.name === EXTERNAL_REFERENCES_TAG) {
      const refs: Array<{ rId: string }> = [];
      for (const er of findChildren(child, EXTERNAL_REFERENCE_TAG)) {
        const rId = er.attrs[`{${REL_NS}}id`];
        if (rId) refs.push({ rId });
      }
      if (refs.length > 0) wb.externalReferences = refs;
      continue;
    }
    // Lift <pivotCaches><pivotCache cacheId=… r:id=…/></pivotCaches>.
    if (child.name === PIVOT_CACHES_TAG) {
      const caches: Array<{ cacheId: number; rId: string }> = [];
      for (const pc of findChildren(child, PIVOT_CACHE_TAG)) {
        const cacheIdAttr = pc.attrs['cacheId'];
        const rId = pc.attrs[`{${REL_NS}}id`];
        if (cacheIdAttr === undefined || !rId) continue;
        const cacheId = Number.parseInt(cacheIdAttr, 10);
        if (!Number.isInteger(cacheId)) continue;
        caches.push({ cacheId, rId });
      }
      if (caches.length > 0) wb.pivotCaches = caches;
      continue;
    }
    // Lift <fileRecoveryPr> into the typed workbook field.
    if (child.name === FILE_RECOVERY_PR_TAG) {
      const fp: import('../workbook/file-recovery.js').FileRecoveryProperties = {};
      const a = child.attrs;
      const flag = (raw: string | undefined): boolean | undefined => {
        if (raw === '1' || raw === 'true') return true;
        if (raw === '0' || raw === 'false') return false;
        return undefined;
      };
      const ar = flag(a['autoRecover']);
      if (ar !== undefined) fp.autoRecover = ar;
      const cs = flag(a['crashSave']);
      if (cs !== undefined) fp.crashSave = cs;
      const del = flag(a['dataExtractLoad']);
      if (del !== undefined) fp.dataExtractLoad = del;
      const rl = flag(a['repairLoad']);
      if (rl !== undefined) fp.repairLoad = rl;
      if (Object.keys(fp).length > 0) wb.fileRecoveryPr = fp;
      continue;
    }
    if (seenSheets) afterSheets.push(child);
    else beforeSheets.push(child);
  }
  if (beforeSheets.length > 0 || afterSheets.length > 0) {
    wb.workbookXmlExtras = { beforeSheets, afterSheets };
  }
}

const SHOW_OBJECTS_MODES: ReadonlyArray<import('../workbook/workbook-properties.js').ShowObjectsMode> = [
  'all',
  'placeholders',
  'none',
];
const UPDATE_LINKS_MODES: ReadonlyArray<import('../workbook/workbook-properties.js').UpdateLinksMode> = [
  'userSet',
  'never',
  'always',
];

const parseWorkbookProperties = (
  node: XmlNode,
): import('../workbook/workbook-properties.js').WorkbookProperties | undefined => {
  const out: import('../workbook/workbook-properties.js').WorkbookProperties = {};
  const a = node.attrs;
  const flag = (raw: string | undefined): boolean | undefined => {
    if (raw === '1' || raw === 'true') return true;
    if (raw === '0' || raw === 'false') return false;
    return undefined;
  };
  const intAttr = (k: string): number | undefined => {
    if (a[k] === undefined) return undefined;
    const n = Number.parseInt(a[k], 10);
    return Number.isInteger(n) ? n : undefined;
  };

  const bools = [
    'date1904',
    'dateCompatibility',
    'showBorderUnselectedTables',
    'filterPrivacy',
    'promptedSolutions',
    'showInkAnnotation',
    'backupFile',
    'saveExternalLinkValues',
    'hidePivotFieldList',
    'showPivotChartFilter',
    'allowRefreshQuery',
    'publishItems',
    'checkCompatibility',
    'autoCompressPictures',
    'refreshAllConnections',
  ] as const satisfies ReadonlyArray<keyof import('../workbook/workbook-properties.js').WorkbookProperties>;
  for (const k of bools) {
    const v = flag(a[k]);
    if (v !== undefined) out[k] = v;
  }

  const showObjects = a['showObjects'];
  if (showObjects && SHOW_OBJECTS_MODES.includes(showObjects as import('../workbook/workbook-properties.js').ShowObjectsMode)) {
    out.showObjects = showObjects as import('../workbook/workbook-properties.js').ShowObjectsMode;
  }
  const updateLinks = a['updateLinks'];
  if (updateLinks && UPDATE_LINKS_MODES.includes(updateLinks as import('../workbook/workbook-properties.js').UpdateLinksMode)) {
    out.updateLinks = updateLinks as import('../workbook/workbook-properties.js').UpdateLinksMode;
  }
  if (a['codeName'] !== undefined) out.codeName = a['codeName'];
  const dtv = intAttr('defaultThemeVersion');
  if (dtv !== undefined) out.defaultThemeVersion = dtv;

  return Object.keys(out).length > 0 ? out : undefined;
};

const CALC_MODES: ReadonlyArray<import('../workbook/calc-properties.js').CalcMode> = [
  'manual',
  'auto',
  'autoNoTable',
];
const REF_MODES: ReadonlyArray<import('../workbook/calc-properties.js').RefMode> = ['A1', 'R1C1'];

const parseCalcProperties = (
  node: XmlNode,
): import('../workbook/calc-properties.js').CalcProperties | undefined => {
  const out: import('../workbook/calc-properties.js').CalcProperties = {};
  const a = node.attrs;
  const flag = (raw: string | undefined): boolean | undefined => {
    if (raw === '1' || raw === 'true') return true;
    if (raw === '0' || raw === 'false') return false;
    return undefined;
  };
  const intAttr = (k: string): number | undefined => {
    if (a[k] === undefined) return undefined;
    const n = Number.parseInt(a[k], 10);
    return Number.isInteger(n) ? n : undefined;
  };
  const floatAttr = (k: string): number | undefined => {
    if (a[k] === undefined) return undefined;
    const n = Number.parseFloat(a[k]);
    return Number.isFinite(n) ? n : undefined;
  };

  const calcId = intAttr('calcId');
  if (calcId !== undefined) out.calcId = calcId;
  const calcMode = a['calcMode'];
  if (calcMode && CALC_MODES.includes(calcMode as import('../workbook/calc-properties.js').CalcMode)) {
    out.calcMode = calcMode as import('../workbook/calc-properties.js').CalcMode;
  }
  const fcol = flag(a['fullCalcOnLoad']);
  if (fcol !== undefined) out.fullCalcOnLoad = fcol;
  const refMode = a['refMode'];
  if (refMode && REF_MODES.includes(refMode as import('../workbook/calc-properties.js').RefMode)) {
    out.refMode = refMode as import('../workbook/calc-properties.js').RefMode;
  }
  const iterate = flag(a['iterate']);
  if (iterate !== undefined) out.iterate = iterate;
  const iterateCount = intAttr('iterateCount');
  if (iterateCount !== undefined) out.iterateCount = iterateCount;
  const iterateDelta = floatAttr('iterateDelta');
  if (iterateDelta !== undefined) out.iterateDelta = iterateDelta;
  const fullPrecision = flag(a['fullPrecision']);
  if (fullPrecision !== undefined) out.fullPrecision = fullPrecision;
  const calcCompleted = flag(a['calcCompleted']);
  if (calcCompleted !== undefined) out.calcCompleted = calcCompleted;
  const calcOnSave = flag(a['calcOnSave']);
  if (calcOnSave !== undefined) out.calcOnSave = calcOnSave;
  const concurrentCalc = flag(a['concurrentCalc']);
  if (concurrentCalc !== undefined) out.concurrentCalc = concurrentCalc;
  const concurrentManualCount = intAttr('concurrentManualCount');
  if (concurrentManualCount !== undefined) out.concurrentManualCount = concurrentManualCount;
  const forceFullCalc = flag(a['forceFullCalc']);
  if (forceFullCalc !== undefined) out.forceFullCalc = forceFullCalc;

  return Object.keys(out).length > 0 ? out : undefined;
};

const SHOW_COMMENTS_MODES: ReadonlyArray<import('../workbook/views.js').CustomViewShowComments> = [
  'commNone',
  'commIndicator',
  'commIndAndComment',
];
const SHOW_OBJECTS_CV_MODES: ReadonlyArray<import('../workbook/views.js').CustomViewShowObjects> = [
  'all',
  'placeholders',
  'none',
];

const parseCustomWorkbookView = (
  node: XmlNode,
): import('../workbook/views.js').CustomWorkbookView | undefined => {
  const a = node.attrs;
  const name = a['name'];
  const guid = a['guid'];
  if (!name || !guid) return undefined;
  const flag = (raw: string | undefined): boolean | undefined => {
    if (raw === '1' || raw === 'true') return true;
    if (raw === '0' || raw === 'false') return false;
    return undefined;
  };
  const intAttr = (k: string): number | undefined => {
    if (a[k] === undefined) return undefined;
    const n = Number.parseInt(a[k], 10);
    return Number.isInteger(n) ? n : undefined;
  };
  const ww = intAttr('windowWidth') ?? 0;
  const wh = intAttr('windowHeight') ?? 0;
  const asid = intAttr('activeSheetId') ?? 0;
  const out: import('../workbook/views.js').CustomWorkbookView = {
    name,
    guid,
    windowWidth: ww,
    windowHeight: wh,
    activeSheetId: asid,
  };

  const boolKeys = [
    'autoUpdate',
    'changesSavedWin',
    'onlySync',
    'personalView',
    'includePrintSettings',
    'includeHiddenRowCol',
    'maximized',
    'minimized',
    'showHorizontalScroll',
    'showVerticalScroll',
    'showSheetTabs',
    'showFormulaBar',
    'showStatusbar',
  ] as const satisfies ReadonlyArray<keyof import('../workbook/views.js').CustomWorkbookView>;
  for (const k of boolKeys) {
    const v = flag(a[k]);
    if (v !== undefined) out[k] = v;
  }
  const intKeys = [
    'mergeInterval',
    'xWindow',
    'yWindow',
    'tabRatio',
  ] as const satisfies ReadonlyArray<keyof import('../workbook/views.js').CustomWorkbookView>;
  for (const k of intKeys) {
    const v = intAttr(k);
    if (v !== undefined) out[k] = v;
  }

  const sc = a['showComments'];
  if (sc && SHOW_COMMENTS_MODES.includes(sc as import('../workbook/views.js').CustomViewShowComments)) {
    out.showComments = sc as import('../workbook/views.js').CustomViewShowComments;
  }
  const so = a['showObjects'];
  if (so && SHOW_OBJECTS_CV_MODES.includes(so as import('../workbook/views.js').CustomViewShowObjects)) {
    out.showObjects = so as import('../workbook/views.js').CustomViewShowObjects;
  }
  return out;
};

const VISIBILITIES: ReadonlyArray<import('../workbook/views.js').WorkbookViewVisibility> = [
  'visible',
  'hidden',
  'veryHidden',
];

const parseWorkbookView = (node: XmlNode): import('../workbook/views.js').WorkbookView => {
  const out: import('../workbook/views.js').WorkbookView = {};
  const a = node.attrs;
  const flag = (raw: string | undefined): boolean | undefined => {
    if (raw === '1' || raw === 'true') return true;
    if (raw === '0' || raw === 'false') return false;
    return undefined;
  };
  const intAttr = (k: string): number | undefined => {
    if (a[k] === undefined) return undefined;
    const n = Number.parseInt(a[k], 10);
    return Number.isInteger(n) ? n : undefined;
  };

  const visibility = a['visibility'];
  if (visibility && VISIBILITIES.includes(visibility as import('../workbook/views.js').WorkbookViewVisibility)) {
    out.visibility = visibility as import('../workbook/views.js').WorkbookViewVisibility;
  }
  const minimized = flag(a['minimized']);
  if (minimized !== undefined) out.minimized = minimized;
  const shScroll = flag(a['showHorizontalScroll']);
  if (shScroll !== undefined) out.showHorizontalScroll = shScroll;
  const svScroll = flag(a['showVerticalScroll']);
  if (svScroll !== undefined) out.showVerticalScroll = svScroll;
  const sst = flag(a['showSheetTabs']);
  if (sst !== undefined) out.showSheetTabs = sst;
  const xWindow = intAttr('xWindow');
  if (xWindow !== undefined) out.xWindow = xWindow;
  const yWindow = intAttr('yWindow');
  if (yWindow !== undefined) out.yWindow = yWindow;
  const ww = intAttr('windowWidth');
  if (ww !== undefined) out.windowWidth = ww;
  const wh = intAttr('windowHeight');
  if (wh !== undefined) out.windowHeight = wh;
  const tr = intAttr('tabRatio');
  if (tr !== undefined) out.tabRatio = tr;
  const fs = intAttr('firstSheet');
  if (fs !== undefined) out.firstSheet = fs;
  const at = intAttr('activeTab');
  if (at !== undefined) out.activeTab = at;
  // `xr2:uid` and friends: anything namespaced is outside CT_BookView, so keep
  // it verbatim rather than dropping it.
  const extAttrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(a)) {
    if (k.startsWith('{')) extAttrs[k] = v;
  }
  if (Object.keys(extAttrs).length > 0) out.extAttrs = extAttrs;
  const adg = flag(a['autoFilterDateGrouping']);
  if (adg !== undefined) out.autoFilterDateGrouping = adg;
  return out;
};

const parseWorkbookProtection = (node: XmlNode): import('../workbook/protection.js').WorkbookProtection => {
  const out: import('../workbook/protection.js').WorkbookProtection = {};
  const a = node.attrs;
  const flag = (raw: string | undefined): boolean | undefined => {
    if (raw === '1' || raw === 'true') return true;
    if (raw === '0' || raw === 'false') return false;
    return undefined;
  };
  if (a['workbookPassword'] !== undefined) out.workbookPassword = a['workbookPassword'];
  if (a['workbookPasswordCharacterSet'] !== undefined)
    out.workbookPasswordCharacterSet = a['workbookPasswordCharacterSet'];
  if (a['workbookAlgorithmName'] !== undefined) out.workbookAlgorithmName = a['workbookAlgorithmName'];
  if (a['workbookHashValue'] !== undefined) out.workbookHashValue = a['workbookHashValue'];
  if (a['workbookSaltValue'] !== undefined) out.workbookSaltValue = a['workbookSaltValue'];
  if (a['workbookSpinCount'] !== undefined) {
    const n = Number.parseInt(a['workbookSpinCount'], 10);
    if (Number.isInteger(n)) out.workbookSpinCount = n;
  }
  if (a['revisionsPassword'] !== undefined) out.revisionsPassword = a['revisionsPassword'];
  if (a['revisionsPasswordCharacterSet'] !== undefined)
    out.revisionsPasswordCharacterSet = a['revisionsPasswordCharacterSet'];
  if (a['revisionsAlgorithmName'] !== undefined) out.revisionsAlgorithmName = a['revisionsAlgorithmName'];
  if (a['revisionsHashValue'] !== undefined) out.revisionsHashValue = a['revisionsHashValue'];
  if (a['revisionsSaltValue'] !== undefined) out.revisionsSaltValue = a['revisionsSaltValue'];
  if (a['revisionsSpinCount'] !== undefined) {
    const n = Number.parseInt(a['revisionsSpinCount'], 10);
    if (Number.isInteger(n)) out.revisionsSpinCount = n;
  }
  const ls = flag(a['lockStructure']);
  if (ls !== undefined) out.lockStructure = ls;
  const lw = flag(a['lockWindows']);
  if (lw !== undefined) out.lockWindows = lw;
  const lr = flag(a['lockRevision']);
  if (lr !== undefined) out.lockRevision = lr;
  return out;
};

/**
 * Capture workbook-rels entries that don't match a modeled type so the writer
 * can re-emit them with their original Id (and any captured `<pivotCaches
 * r:id="…"/>` etc. still resolves after a round-trip). Modeled non-sheet rels
 * (sst / styles / theme / vbaProject) keep their original Id via
 * `wb.workbookRelOriginalIds` so the writer can prefer those over freshly
 * allocated ones.
 */
function captureWorkbookRelsExtras(
  wbRels: import('../packaging/relationships.js').Relationships,
  wb: Workbook,
): void {
  const SHEET_RELS = new Set([`${REL_NS}/worksheet`, `${REL_NS}/chartsheet`]);
  const original: NonNullable<Workbook['workbookRelOriginalIds']> = {};
  const extras: Array<{ id: string; type: string; target: string }> = [];
  for (const rel of wbRels.rels) {
    if (SHEET_RELS.has(rel.type)) continue;
    if (rel.type === `${REL_NS}/sharedStrings`) {
      original.sharedStrings = rel.id;
      continue;
    }
    if (rel.type === `${REL_NS}/styles`) {
      original.styles = rel.id;
      continue;
    }
    if (rel.type === `${REL_NS}/theme`) {
      original.theme = rel.id;
      continue;
    }
    if (rel.type === `${REL_NS}/vbaProject`) {
      original.vbaProject = rel.id;
      continue;
    }
    extras.push({ id: rel.id, type: rel.type, target: rel.target });
  }
  if (Object.keys(original).length > 0) wb.workbookRelOriginalIds = original;
  if (extras.length > 0) wb.workbookRelsExtras = extras;
}

const PASSTHROUGH_PREFIXES: ReadonlyArray<string> = [
  'xl/activeX/',
  'xl/ctrlProps/',
  'xl/embeddings/',
  'xl/externalLinks/',
  // xl/model/ — Power Pivot data model (`xl/model/item.data` etc.).
  'xl/model/',
  'xl/persons/',
  'xl/pivotCache/',
  'xl/pivotTables/',
  'xl/printerSettings/',
  'xl/queryTables/',
  'xl/richData/',
  'xl/slicerCaches/',
  'xl/slicers/',
  'xl/threadedComments/',
  'xl/timelineCaches/',
  'xl/timelines/',
  'xl/workbookCache/',
  'customUI/',
  'customXml/',
];

/**
 * Excel emits both form-control VMLs and comment VMLs at
 * `xl/drawings/vmlDrawingN.vml`. Filename alone can't tell them apart, but
 * ECMA-376 §17.18.51 requires comment shapes to carry `<x:ClientData
 * ObjectType="Note">`, so a byte-search for that marker decides which path the
 * file belongs on:
 *
 *  - With marker → comment VML; the comments writer regenerates
 * these from `Worksheet.legacyComments`, so we must not capture them as
 * passthrough (would duplicate the entry on save).
 *  - Without marker → control / OLE / shape VML; capture as
 * passthrough so form controls survive load → save → load.
 */
const COMMENT_VML_MARKER = 'ObjectType="Note"';
// Latin-1 is a 1-to-1 byte-to-character mapping for the 0–255 range, so a
// String.indexOf scan on the decoded view returns the same answer as a byte
// search but takes advantage of V8's native StringPrototypeIndexOf — orders
// of magnitude faster than the JS loop the previous implementation used on
// multi-megabyte VML drawings.
const LATIN1_DECODER = new TextDecoder('latin1');

const isVmlDrawing = (path: string): boolean =>
  path.startsWith('xl/drawings/') && path.endsWith('.vml');

const containsCommentMarker = (bytes: Uint8Array): boolean =>
  LATIN1_DECODER.decode(bytes).includes(COMMENT_VML_MARKER);

interface VmlPartCache {
  /** Inflated VML bytes, read once per part. */
  read(path: string): Uint8Array;
  /** True when the part at `path` exists and is a comment overlay. */
  isCommentVml(path: string): boolean;
}

/**
 * Both the per-sheet rels capture and the passthrough sweep need to classify
 * VML parts; share one inflate per part between them.
 */
const makeVmlPartCache = (archive: ZipArchive): VmlPartCache => {
  const bytesByPath = new Map<string, Uint8Array>();
  const read = (path: string): Uint8Array => {
    let bytes = bytesByPath.get(path);
    if (bytes === undefined) {
      bytes = archive.read(path);
      bytesByPath.set(path, bytes);
    }
    return bytes;
  };
  return {
    read,
    isCommentVml: (path) => archive.has(path) && containsCommentMarker(read(path)),
  };
};

/**
 * Top-level xl/*.xml files that aren't modeled but Excel relies on (or
 * harmlessly preserves). Captured by exact path; their content types come
 * through the manifest Override map.
 *
 * - `xl/calcChain.xml`     — calculation order hint (Excel rebuilds
 * it on first open if missing, but losing it forces a full recalc).
 * - `xl/connections.xml`   — external data connection metadata.
 * - `xl/persons/`          — threaded-comment author registry
 * (Excel 365). Captured under the prefix list below.
 * - `xl/metadata.xml`      — Excel 365 dynamic-array cell metadata.
 * - `xl/SheetMetadata.xml` — variant casing of the same.
 */
const PASSTHROUGH_EXACT_PATHS: ReadonlySet<string> = new Set([
  'xl/calcChain.xml',
  'xl/connections.xml',
  'xl/metadata.xml',
  'xl/SheetMetadata.xml',
  // docProps/thumbnail.jpeg — workbook preview image Excel renders in the OS
  // file browser. JPEG by default; some files use PNG.
  'docProps/thumbnail.jpeg',
  'docProps/thumbnail.jpg',
  'docProps/thumbnail.png',
  'docProps/thumbnail.wmf',
  'docProps/thumbnail.emf',
]);

const isPassthroughPath = (path: string, vml: VmlPartCache): boolean => {
  if (PASSTHROUGH_EXACT_PATHS.has(path)) return true;
  if (PASSTHROUGH_PREFIXES.some((p) => path.startsWith(p))) return true;
  // Comment VML is regenerated; control / shape / header-footer VML passes
  // through.
  if (isVmlDrawing(path)) return !vml.isCommentVml(path);
  return false;
};

/**
 * Walk the archive after the modeled parts are loaded and capture any remaining
 * content into `wb.passthrough`. The dedicated VBA project binaries land on
 * their own slots so the writer can promote the workbook content type to xlsm.
 *
 * `roots` are parts the worksheet writer will reference by their original
 * r:id (see `captureSheetRelsExtras`). Each root is captured together with
 * its own `.rels` file and, transitively, every internal target of those rels
 * — that is what keeps a header/footer VML's image (`vmlDrawingN.vml.rels` →
 * `xl/media/imageN.jpeg`) attached. Passthrough VML parts get the same
 * treatment because form-control VML references images the same way.
 */
function capturePassthrough(
  archive: ZipArchive,
  manifest: import('../packaging/manifest.js').Manifest,
  wb: Workbook,
  roots: ReadonlyArray<string>,
  vml: VmlPartCache,
): void {
  const overrides = new Map<string, string>();
  for (const o of manifest.overrides) {
    // Manifest paths are package-absolute (`/xl/...`); strip the leading slash.
    overrides.set(o.partName.replace(/^\//, ''), o.contentType);
  }
  const defaults = new Map<string, string>();
  for (const d of manifest.defaults) defaults.set(d.ext.toLowerCase(), d.contentType);

  const capture = (path: string, bytes: Uint8Array): void => {
    if (!wb.passthrough) wb.passthrough = new Map();
    wb.passthrough.set(path, bytes);
    const ct = overrides.get(path);
    if (ct !== undefined) {
      if (!wb.passthroughContentTypes) wb.passthroughContentTypes = new Map();
      wb.passthroughContentTypes.set(path, ct);
      return;
    }
    // No Override → the part is typed by its extension's Default. Remember it
    // so the writer can re-emit that Default; it only knows the extensions of
    // parts it produces itself.
    const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
    const dct = defaults.get(ext);
    if (dct !== undefined) {
      if (!wb.passthroughDefaults) wb.passthroughDefaults = new Map();
      wb.passthroughDefaults.set(ext, dct);
    }
  };

  const closureRoots: string[] = [...roots];
  for (const path of archive.list()) {
    if (path === 'xl/vbaProject.bin') {
      wb.vbaProject = archive.read(path);
      continue;
    }
    if (path === 'xl/vbaProjectSignature.bin') {
      wb.vbaSignature = archive.read(path);
      continue;
    }
    if (!isPassthroughPath(path, vml)) continue;
    if (isVmlDrawing(path)) {
      capture(path, vml.read(path));
      closureRoots.push(path);
    } else {
      capture(path, archive.read(path));
    }
  }

  // Breadth-first over rels so a chain like sheet → VML → image is captured
  // whole. `visited` (not `wb.passthrough`) gates the walk because roots
  // captured by the prefix scan above still need their rels followed.
  const visited = new Set<string>();
  // for…of re-reads the array length each step, so targets pushed mid-walk
  // are visited too.
  for (const path of closureRoots) {
    if (visited.has(path) || !archive.has(path)) continue;
    visited.add(path);
    if (!wb.passthrough?.has(path)) capture(path, archive.read(path));
    const relsPath = relsPathFor(path);
    if (!archive.has(relsPath)) continue;
    const relsBytes = archive.read(relsPath);
    capture(relsPath, relsBytes);
    for (const r of relsFromBytes(relsBytes).rels) {
      if (r.targetMode !== 'External') closureRoots.push(resolveRelTarget(path, r.target));
    }
  }
}
