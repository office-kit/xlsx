// Public `saveWorkbook` entry point.
//
// **Stage 1 minimum**: emits the bare set of parts a Workbook needs to
// round-trip through `loadWorkbook`:
//
//     [Content_Types].xml
//     _rels/.rels
//     xl/workbook.xml
//     xl/_rels/workbook.xml.rels
//     xl/worksheets/sheetN.xml ...
//     xl/styles.xml
//     xl/sharedStrings.xml             (only when sst is non-empty)
//
// docProps / theme / VBA / drawings / charts are reserved for later iterations
// — load tolerates their absence.

import { escapeXmlAttr, escapeXmlText } from '../utils/escape';
import { chartToBytes } from '../chart/chart-xml';
import { chartExToBytes } from '../chart/cx/chartex-xml';
import { userShapesToBytes } from '../chart/user-shapes-xml';
import { chartsheetToBytes } from '../chartsheet/chartsheet-xml';
import type { Drawing, DrawingItem } from '../drawing/drawing';
import { drawingToBytes } from '../drawing/drawing-xml';
import { IMAGE_FORMAT_EXTENSION, IMAGE_FORMAT_MIME, type XlsxImageFormat } from '../drawing/image';
import type { XlsxSink } from '../io/sink';
import { OpenXmlIoError, OpenXmlSchemaError } from '../utils/exceptions';
import { corePropsToBytes } from '../packaging/core';
import { customPropsToBytes } from '../packaging/custom';
import { extendedPropsToBytes } from '../packaging/extended';
import { addDefault, addOverride, makeManifest, manifestToBytes } from '../packaging/manifest';
import { makeRelationships, type Relationships, relsToBytes } from '../packaging/relationships';
import { stylesheetToBytes } from '../styles/stylesheet-writer';
import { makeSharedStrings, sharedStringsToBytes } from '../workbook/shared-strings';
import { type Workbook, validateSheetTitle } from '../workbook/workbook';
import type { LegacyComment } from '../worksheet/comments';
import { commentsToBytes, placeholderVmlDrawing } from '../worksheet/comments-xml';
import type { TableDefinition } from '../worksheet/table';
import { tableToBytes } from '../worksheet/table-xml';
import { worksheetToBytes } from '../worksheet/writer';
import { serializeXml as serializeXmlNode } from '../xml/serializer';
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
  ARC_WORKBOOK_RELS,
  CHARTEX_TYPE,
  CPROPS_TYPE,
  PKG_REL_NS,
  REL_NS,
  SHARED_STRINGS_TYPE,
  SHEET_MAIN_NS,
  STYLES_TYPE,
  THEME_TYPE,
  WORKSHEET_TYPE,
  XLSX_TYPE,
} from '../xml/namespaces';
import { createZipWriter } from '../zip/writer';

const CORE_PROPS_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';
const EXT_PROPS_TYPE = 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
const CORE_PROPS_REL = `${PKG_REL_NS}/metadata/core-properties`;
const EXT_PROPS_REL = `${REL_NS}/extended-properties`;
const CUSTOM_PROPS_REL = `${REL_NS}/custom-properties`;
const THEME_REL = `${REL_NS}/theme`;
const TABLE_REL = `${REL_NS}/table`;
const TABLE_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const COMMENTS_REL = `${REL_NS}/comments`;
const VML_DRAWING_REL = `${REL_NS}/vmlDrawing`;
const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
const VML_DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.vmlDrawing';
const DRAWING_REL = `${REL_NS}/drawing`;
const DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const CHART_REL = `${REL_NS}/chart`;
const CHARTEX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const IMAGE_REL = `${REL_NS}/image`;
const CHARTSHEET_REL = `${REL_NS}/chartsheet`;
const CHARTSHEET_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml';
const CHART_USER_SHAPES_REL = `${REL_NS}/chartUserShapes`;

export interface SaveOptions {
  /** Reserved — passes through to the underlying ZIP writer when implemented. */
  compressionLevel?: number;
}

/** Convenience: serialise a Workbook to an in-memory `Uint8Array` xlsx.
 *
 * Browser-safe: this path uses an in-memory `Uint8Array` sink and never
 * touches Node's `Buffer` global, so it works unchanged in browser bundles.
 */
export async function workbookToBytes(wb: Workbook, opts: SaveOptions = {}): Promise<Uint8Array> {
  const sink = toUint8ArraySink();
  await saveWorkbook(wb, sink, opts);
  return sink.result();
}

const toUint8ArraySink = (): XlsxSink & { result(): Uint8Array } => {
  const chunks: Uint8Array[] = [];
  let finalised: Uint8Array | undefined;
  const finalise = (): Uint8Array => {
    if (finalised !== undefined) return finalised;
    let total = 0;
    for (const c of chunks) total += c.byteLength;
    const out = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      out.set(c, off);
      off += c.byteLength;
    }
    finalised = out;
    chunks.length = 0;
    return out;
  };
  return {
    toBytes() {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) {
            throw new OpenXmlIoError('workbookToBytes sink: write after finish');
          }
          if (!(chunk instanceof Uint8Array)) {
            throw new OpenXmlIoError('workbookToBytes sink: chunk is not a Uint8Array');
          }
          chunks.push(chunk);
        },
        async finish(): Promise<Uint8Array> {
          return finalise();
        },
        abort(): void {
          if (finalised !== undefined) return;
          finalised = new Uint8Array(0);
          chunks.length = 0;
        },
      };
    },
    result(): Uint8Array {
      return finalise();
    },
  };
};

/**
 * Validate every sheet title against Excel's character + length rules and
 * confirm titles are unique within the workbook. Catches bad state introduced
 * via direct mutation (`ws.title = ...`) that bypassed the public mutators.
 *
 * Without this gate, a workbook with an invalid sheet name would round-trip to
 * an xlsx Excel rejects with a generic "file is corrupt" dialog on open —
 * making it hard to diagnose. Raising at save time keeps the bad state
 * recoverable: callers can `renameSheet` and retry.
 */
const validateSheetTitles = (wb: Workbook): void => {
  const seen = new Map<string, number>();
  for (let i = 0; i < wb.sheets.length; i++) {
    const ref = wb.sheets[i];
    if (!ref) continue;
    const title = ref.sheet.title;
    const reason = validateSheetTitle(title);
    if (reason) {
      throw new OpenXmlSchemaError(`saveWorkbook: sheet[${i}] title "${title}": ${reason}`);
    }
    const lower = title.toLowerCase();
    const prior = seen.get(lower);
    if (prior !== undefined) {
      throw new OpenXmlSchemaError(
        `saveWorkbook: sheet[${i}] title "${title}" collides with sheet[${prior}]` +
          ` (Excel treats sheet names case-insensitively for uniqueness).`,
      );
    }
    seen.set(lower, i);
  }
};

/** Save a workbook through the given sink. Returns once `finalize()` resolves. */
export async function saveWorkbook(wb: Workbook, sink: XlsxSink, _opts: SaveOptions = {}): Promise<void> {
  validateSheetTitles(wb);
  const writer = createZipWriter(sink);
  try {
    await saveWorkbookImpl(wb, writer);
  } catch (err) {
    // Release the sink so streaming destinations (`toFile` / `toWritable`)
    // don't leave a half-written file looking valid. abort() is idempotent so
    // a successful finalize() above is a no-op here.
    writer.abort(err);
    throw err;
  }
}

async function saveWorkbookImpl(wb: Workbook, writer: ReturnType<typeof createZipWriter>): Promise<void> {
  // ---- 1. assemble the per-sheet rels + serialise each worksheet ----------
  const sst = makeSharedStrings();
  interface SheetEmit {
    id: string;
    target: string;
    bytes: Uint8Array;
    /** OOXML relationship type used in workbook.xml.rels. */
    relType: string;
    /** ZIP archive path the bytes are written to. */
    archivePath: string;
    /** Override content type for [Content_Types].xml. */
    contentType: string;
    /** Per-sheet rels — populated only if the sheet has hyperlinks / drawings / etc. */
    rels?: Relationships;
  }
  const sheetEmits: SheetEmit[] = [];
  // Workbook-global table counter so xl/tables/tableN.xml ids stay unique.
  const tableEmits: Array<{ id: number; bytes: Uint8Array }> = [];
  let nextTableId = 1;
  // Same counter pattern for comments parts. The comments part and the VML
  // drawing share their N because Excel always emits them paired.
  interface CommentEmit {
    id: number;
    commentsBytes: Uint8Array;
    vmlBytes: Uint8Array;
  }
  const commentEmits: CommentEmit[] = [];
  // Passthrough carries form-control / header-footer VML under the same
  // xl/drawings/vmlDrawingN.vml naming, so start past whatever is already
  // taken — the zip writer rejects duplicate entries.
  let nextCommentsId = nextFreeIndex(wb, /^xl\/drawings\/vmlDrawing(\d+)\.vml$/);
  // Drawings: workbook-global drawingN counter for xl/drawings/drawingN.xml +
  // matching drawing-rels file (when chart items are present).
  interface DrawingEmit {
    id: number;
    bytes: Uint8Array;
    /** Drawing rels — when non-empty we also emit the rels file. */
    rels?: Relationships;
  }
  const drawingEmits: DrawingEmit[] = [];
  let nextDrawingId = 1;
  // Chart parts share a workbook-global counter so xl/charts/chartN.xml ids
  // stay unique.
  const chartEmits: Array<{
    id: number;
    bytes: Uint8Array;
    isCx: boolean;
    /** Per-chart rels file (only emitted when chart.userShapes is set). */
    rels?: Relationships;
  }> = [];
  let nextChartId = 1;
  // Workbook-global counter for chartDrawing parts. Each chart with userShapes
  // set produces one xl/drawings/chartDrawingN.xml entry.
  const userShapeEmits: Array<{ id: number; bytes: Uint8Array }> = [];
  let nextUserShapesId = 1;
  // Workbook-global counter for image media parts. Excel uses
  // xl/media/imageN.{ext} where N is shared across the package.
  interface ImageEmit {
    id: number;
    /** File extension (without the dot) — also drives the manifest Default. */
    ext: string;
    bytes: Uint8Array;
  }
  const imageEmits: ImageEmit[] = [];
  const imageExts = new Set<string>();
  // Same collision guard as for VML: images referenced by passthrough parts
  // keep their original xl/media/imageN.* names.
  let nextImageId = nextFreeIndex(wb, /^xl\/media\/image(\d+)\./);
  // Track separate counters so worksheet / chartsheet IDs do not collide (Excel
  // uses xl/worksheets/sheetN.xml and xl/chartsheets/sheetM.xml independently
  // of one another).
  let nextWorksheetId = 1;
  let nextChartsheetId = 1;
  // Pre-claim every original rId so freshly allocated ones (sheets without a
  // captured rId, plus modeled non-sheet rels) avoid collision with captured
  // workbookRelsExtras and the modeled non-sheet original ids.
  const claimedRIds = new Set<string>();
  for (const ref of wb.sheets) {
    if (ref.rId !== undefined) claimedRIds.add(ref.rId);
  }
  if (wb.workbookRelOriginalIds) {
    for (const v of Object.values(wb.workbookRelOriginalIds)) {
      if (typeof v === 'string') claimedRIds.add(v);
    }
  }
  if (wb.workbookRelsExtras) {
    for (const e of wb.workbookRelsExtras) claimedRIds.add(e.id);
  }
  let nextRIdCursor = 1;
  const allocateRId = (): string => {
    let id = `rId${nextRIdCursor}`;
    while (claimedRIds.has(id)) {
      nextRIdCursor++;
      id = `rId${nextRIdCursor}`;
    }
    claimedRIds.add(id);
    nextRIdCursor++;
    return id;
  };
  wb.sheets.forEach((ref, _i) => {
    const isChartsheet = ref.kind === 'chartsheet';
    const target = isChartsheet ? `chartsheets/sheet${nextChartsheetId}.xml` : `worksheets/sheet${nextWorksheetId}.xml`;
    const archivePath = isChartsheet
      ? `xl/chartsheets/sheet${nextChartsheetId}.xml`
      : `xl/worksheets/sheet${nextWorksheetId}.xml`;
    if (isChartsheet) nextChartsheetId++;
    else nextWorksheetId++;
    const sheetRels = makeRelationships();
    // Pre-claim every captured relsExtras rId so freshly allocated modeled rels
    // never clash. The extras themselves are appended to sheetRels at the end
    // so any rId already used inside the worksheet body (drawing / hyperlink
    // refs) keeps pointing at our newly emitted rel.
    const sheetRelsClaimed = new Set<string>();
    const sheetRelsExtras = ref.kind === 'worksheet' ? (ref.sheet.relsExtras ?? []) : [];
    for (const e of sheetRelsExtras) sheetRelsClaimed.add(e.id);
    let sheetRIdCursor = 1;
    const allocateSheetRId = (): string => {
      let id = `rId${sheetRIdCursor}`;
      while (sheetRelsClaimed.has(id) || sheetRels.rels.some((r) => r.id === id)) {
        sheetRIdCursor++;
        id = `rId${sheetRIdCursor}`;
      }
      sheetRIdCursor++;
      return id;
    };
    const registerTable = (table: TableDefinition): { rId: string } => {
      const tableId = nextTableId++;
      const rId = allocateSheetRId();
      sheetRels.rels.push({
        id: rId,
        type: TABLE_REL,
        target: `../tables/table${tableId}.xml`,
      });
      // Emit the table part with its workbook-global id baked in.
      const xmlTable: TableDefinition = { ...table, id: tableId };
      tableEmits.push({ id: tableId, bytes: tableToBytes(xmlTable) });
      return { rId };
    };
    const registerComments = (comments: ReadonlyArray<LegacyComment>): { vmlRelId: string } => {
      const id = nextCommentsId++;
      const commentsRelId = allocateSheetRId();
      sheetRels.rels.push({
        id: commentsRelId,
        type: COMMENTS_REL,
        target: `../comments${id}.xml`,
      });
      const vmlRelId = allocateSheetRId();
      sheetRels.rels.push({
        id: vmlRelId,
        type: VML_DRAWING_REL,
        target: `../drawings/vmlDrawing${id}.vml`,
      });
      commentEmits.push({
        id,
        commentsBytes: commentsToBytes(comments),
        vmlBytes: placeholderVmlDrawing(comments),
      });
      return { vmlRelId };
    };
    const registerDrawing = (drawing: Drawing): { rId: string } => {
      const id = nextDrawingId++;
      const rId = allocateSheetRId();
      sheetRels.rels.push({
        id: rId,
        type: DRAWING_REL,
        target: `../drawings/drawing${id}.xml`,
      });
      // Walk drawing items: for each chart with a payload, allocate a
      // workbook-global chartN id and a per-drawing rId. Emit the chart part +
      // collect a drawing-rels entry so drawing.xml's <c:chart r:id> resolves.
      const drawingRels = makeRelationships();
      const itemsForXml: DrawingItem[] = [];
      for (const item of drawing.items) {
        if (item.content.kind === 'chart' && (item.content.chart.space || item.content.chart.cxSpace)) {
          const chartId = nextChartId++;
          const chartRId = `rId${drawingRels.rels.length + 1}`;
          // chartex parts use a different relationship Type than the legacy
          // ECMA-376 charts. Excel rejects the workbook when the drawing rels
          // claim a `relationships/chart` target that actually contains a
          // `cx:chartSpace` root.
          const isCxChart = item.content.chart.cxSpace !== undefined;
          drawingRels.rels.push({
            id: chartRId,
            type: isCxChart ? CHARTEX_REL : CHART_REL,
            target: `../charts/chart${chartId}.xml`,
          });
          if (item.content.chart.cxSpace) {
            chartEmits.push({
              id: chartId,
              bytes: chartExToBytes(item.content.chart.cxSpace),
              isCx: true,
            });
          } else if (item.content.chart.space) {
            const space = item.content.chart.space;
            // If the chart carries user shapes, allocate the chartDrawingN part
            // + a per-chart rels file referencing it, then bake the resulting
            // r:id into the chart's <c:userShapes> element.
            let userShapesRId: string | undefined;
            let chartRelsFile: Relationships | undefined;
            if (space.userShapes && space.userShapes.shapes.length > 0) {
              const userShapesId = nextUserShapesId++;
              userShapesRId = 'rId1';
              chartRelsFile = makeRelationships();
              chartRelsFile.rels.push({
                id: userShapesRId,
                type: CHART_USER_SHAPES_REL,
                target: `../drawings/chartDrawing${userShapesId}.xml`,
              });
              userShapeEmits.push({
                id: userShapesId,
                bytes: userShapesToBytes(space.userShapes),
              });
            }
            const chartEmit: typeof chartEmits[number] = {
              id: chartId,
              bytes: chartToBytes(space, userShapesRId !== undefined ? { userShapesRId } : {}),
              isCx: false,
            };
            if (chartRelsFile) chartEmit.rels = chartRelsFile;
            chartEmits.push(chartEmit);
          }
          itemsForXml.push({
            anchor: item.anchor,
            content: {
              kind: 'chart',
              chart: {
                rId: chartRId,
                ...(item.content.chart.cxSpace ? { isCx: true } : {}),
              },
            },
          });
        } else if (item.content.kind === 'picture' && item.content.picture.image) {
          const img = item.content.picture.image;
          const ext = IMAGE_FORMAT_EXTENSION[img.format];
          const imageId = nextImageId++;
          const picRId = `rId${drawingRels.rels.length + 1}`;
          drawingRels.rels.push({
            id: picRId,
            type: IMAGE_REL,
            target: `../media/image${imageId}.${ext}`,
          });
          imageEmits.push({ id: imageId, ext, bytes: img.bytes });
          imageExts.add(ext);
          itemsForXml.push({
            anchor: item.anchor,
            content: {
              kind: 'picture',
              picture: {
                rId: picRId,
                ...(item.content.picture.name !== undefined ? { name: item.content.picture.name } : {}),
                ...(item.content.picture.descr !== undefined ? { descr: item.content.picture.descr } : {}),
                ...(item.content.picture.hidden !== undefined ? { hidden: item.content.picture.hidden } : {}),
                ...(item.content.picture.spPr ? { spPr: item.content.picture.spPr } : {}),
              },
            },
          });
        } else {
          itemsForXml.push(item);
        }
      }
      const emit: DrawingEmit = { id, bytes: drawingToBytes({ items: itemsForXml }) };
      if (drawingRels.rels.length > 0) emit.rels = drawingRels;
      drawingEmits.push(emit);
      return { rId };
    };
    let bytes: Uint8Array;
    if (ref.kind === 'worksheet') {
      bytes = worksheetToBytes(ref.sheet, {
        sharedStrings: sst,
        date1904: wb.date1904,
        rels: sheetRels,
        registerTable,
        registerComments,
        registerDrawing,
      });
    } else {
      // Chartsheet: register the drawing (if any), then emit the chartsheet
      // part with the resulting r:id baked in.
      let drawingRId: string | undefined;
      if (ref.sheet.drawing) drawingRId = registerDrawing(ref.sheet.drawing).rId;
      bytes = chartsheetToBytes(ref.sheet, drawingRId !== undefined ? { drawingRId } : {});
    }
    // Append captured per-sheet rels passthrough (pivotTable / queryTable /
    // printerSettings / oleObject / customProperty / threadedComment …)
    // verbatim. Their original rIds were pre-claimed so the modeled allocations
    // above never collide with them.
    for (const e of sheetRelsExtras) {
      sheetRels.rels.push({ id: e.id, type: e.type, target: e.target });
    }
    const sheetRId = ref.rId ?? allocateRId();
    const emit: SheetEmit = {
      id: sheetRId,
      target,
      bytes,
      relType: isChartsheet ? CHARTSHEET_REL : `${REL_NS}/worksheet`,
      archivePath,
      contentType: isChartsheet ? CHARTSHEET_TYPE : WORKSHEET_TYPE,
    };
    if (sheetRels.rels.length > 0) emit.rels = sheetRels;
    sheetEmits.push(emit);
  });

  // ---- 2. workbook rels -- sheets first, then sst (if any), then styles, then
  // theme / vbaProject, then any captured workbookRelsExtras (e.g.
  // pivotCacheDefinition rels referenced by `<pivotCaches>`). Modeled non-sheet
  // rels prefer the rId captured at load time so any captured extras XML using
  // that Id still resolves after the round-trip.
  const wbRels = makeRelationships();
  wbRels.rels = sheetEmits.map((e) => ({
    id: e.id,
    type: e.relType,
    target: e.target,
  }));
  const orig = wb.workbookRelOriginalIds;
  if (sst.entries.length > 0) {
    wbRels.rels.push({
      id: orig?.sharedStrings ?? allocateRId(),
      type: `${REL_NS}/sharedStrings`,
      target: 'sharedStrings.xml',
    });
  }
  wbRels.rels.push({
    id: orig?.styles ?? allocateRId(),
    type: `${REL_NS}/styles`,
    target: 'styles.xml',
  });
  if (wb.themeXml) {
    wbRels.rels.push({
      id: orig?.theme ?? allocateRId(),
      type: THEME_REL,
      target: 'theme/theme1.xml',
    });
  }
  if (wb.vbaProject) {
    wbRels.rels.push({
      id: orig?.vbaProject ?? allocateRId(),
      type: `${REL_NS}/vbaProject`,
      target: 'vbaProject.bin',
    });
  }
  if (wb.workbookRelsExtras) {
    for (const e of wb.workbookRelsExtras) {
      wbRels.rels.push({ id: e.id, type: e.type, target: e.target });
    }
  }

  // ---- 3. workbook.xml ----------------------------------------------------
  const workbookXml = serializeWorkbookXml(
    wb,
    sheetEmits.map((e) => e.id),
  );
  await writer.addEntry(ARC_WORKBOOK, new TextEncoder().encode(workbookXml));
  await writer.addEntry(ARC_WORKBOOK_RELS, relsToBytes(wbRels));

  // ---- 4. each worksheet / chartsheet (and its rels file when present) --
  for (const e of sheetEmits) {
    await writer.addEntry(e.archivePath, e.bytes);
    if (e.rels) {
      // The rels file sits alongside its part: `xl/<dir>/_rels/<file>.rels`.
      const slash = e.archivePath.lastIndexOf('/');
      const dir = e.archivePath.slice(0, slash);
      const file = e.archivePath.slice(slash + 1);
      await writer.addEntry(`${dir}/_rels/${file}.rels`, relsToBytes(e.rels));
    }
  }

  // ---- 4b. table parts ----------------------------------------------------
  for (const t of tableEmits) {
    await writer.addEntry(`xl/tables/table${t.id}.xml`, t.bytes);
  }

  // ---- 4c. comments parts + matching VML drawings ------------------------
  for (const c of commentEmits) {
    await writer.addEntry(`xl/comments${c.id}.xml`, c.commentsBytes);
    await writer.addEntry(`xl/drawings/vmlDrawing${c.id}.vml`, c.vmlBytes);
  }

  // ---- 4d. drawings + their rels (when charts are embedded) -------------
  for (const d of drawingEmits) {
    await writer.addEntry(`xl/drawings/drawing${d.id}.xml`, d.bytes);
    if (d.rels) {
      await writer.addEntry(`xl/drawings/_rels/drawing${d.id}.xml.rels`, relsToBytes(d.rels));
    }
  }

  // ---- 4e. chart parts (and per-chart rels when user shapes attach) -----
  for (const c of chartEmits) {
    await writer.addEntry(`xl/charts/chart${c.id}.xml`, c.bytes);
    if (c.rels) {
      await writer.addEntry(`xl/charts/_rels/chart${c.id}.xml.rels`, relsToBytes(c.rels));
    }
  }

  // ---- 4f. user-shape drawings (xl/drawings/chartDrawingN.xml) ----------
  for (const us of userShapeEmits) {
    await writer.addEntry(`xl/drawings/chartDrawing${us.id}.xml`, us.bytes);
  }

  // ---- 4g. embedded images (xl/media/imageN.{ext}) ----------------------
  for (const img of imageEmits) {
    await writer.addEntry(`xl/media/image${img.id}.${img.ext}`, img.bytes);
  }

  // ---- 5. styles.xml + sharedStrings.xml (if any) -------------------------
  await writer.addEntry(ARC_STYLE, stylesheetToBytes(wb.styles));
  if (sst.entries.length > 0) {
    await writer.addEntry(ARC_SHARED_STRINGS, sharedStringsToBytes(sst));
  }

  // ---- 5b. theme1.xml (passthrough) — only when wb carries one. The theme rel
  // was already added to wbRels above so workbook.xml.rels references it.
  if (wb.themeXml) await writer.addEntry(ARC_THEME, wb.themeXml);

  // ---- 5c. docProps/{core,app,custom}.xml (when present) ------------------
  if (wb.properties) await writer.addEntry(ARC_CORE, corePropsToBytes(wb.properties));
  if (wb.appProperties) await writer.addEntry(ARC_APP, extendedPropsToBytes(wb.appProperties));
  if (wb.customProperties) await writer.addEntry(ARC_CUSTOM, customPropsToBytes(wb.customProperties));

  // ---- 5d. VBA project + signature (when present) -------------------------
  if (wb.vbaProject) await writer.addEntry('xl/vbaProject.bin', wb.vbaProject);
  if (wb.vbaSignature) await writer.addEntry('xl/vbaProjectSignature.bin', wb.vbaSignature);

  // ---- 5e. Pass-through bytes (pivot / activeX / customXml / etc.) -------
  if (wb.passthrough) {
    for (const [path, bytes] of wb.passthrough) await writer.addEntry(path, bytes);
  }

  // ---- 6. root rels -------------------------------------------------------
  const rootRels: Relationships = { rels: [] };
  rootRels.rels.push({
    id: 'rId1',
    type: `${REL_NS}/officeDocument`,
    target: 'xl/workbook.xml',
  });
  if (wb.properties) {
    rootRels.rels.push({
      id: `rId${rootRels.rels.length + 1}`,
      type: CORE_PROPS_REL,
      target: 'docProps/core.xml',
    });
  }
  if (wb.appProperties) {
    rootRels.rels.push({
      id: `rId${rootRels.rels.length + 1}`,
      type: EXT_PROPS_REL,
      target: 'docProps/app.xml',
    });
  }
  if (wb.customProperties) {
    rootRels.rels.push({
      id: `rId${rootRels.rels.length + 1}`,
      type: CUSTOM_PROPS_REL,
      target: 'docProps/custom.xml',
    });
  }
  await writer.addEntry(ARC_ROOT_RELS, relsToBytes(rootRels));

  // ---- 7. [Content_Types].xml --------------------------------------------
  const manifest = makeManifest();
  addDefault(manifest, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');
  addDefault(manifest, 'xml', 'application/xml');
  // Defaults carried over for passthrough parts go first: addDefault is
  // last-wins, so the writer's own registrations below (vml for comments, bin
  // for VBA, image formats) still take precedence on a shared extension.
  if (wb.passthroughDefaults) {
    for (const [ext, ct] of wb.passthroughDefaults) addDefault(manifest, ext, ct);
  }
  // VBA-bearing workbooks promote the workbook content type to xlsm.
  const workbookContentType = wb.vbaProject
    ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
    : XLSX_TYPE;
  addOverride(manifest, `/${ARC_WORKBOOK}`, workbookContentType);
  for (const e of sheetEmits) {
    addOverride(manifest, `/${e.archivePath}`, e.contentType);
  }
  addOverride(manifest, `/${ARC_STYLE}`, STYLES_TYPE);
  if (sst.entries.length > 0) {
    addOverride(manifest, `/${ARC_SHARED_STRINGS}`, SHARED_STRINGS_TYPE);
  }
  if (wb.themeXml) addOverride(manifest, `/${ARC_THEME}`, THEME_TYPE);
  for (const t of tableEmits) {
    addOverride(manifest, `/xl/tables/table${t.id}.xml`, TABLE_TYPE);
  }
  if (commentEmits.length > 0) {
    addDefault(manifest, 'vml', VML_DRAWING_TYPE);
  }
  for (const c of commentEmits) {
    addOverride(manifest, `/xl/comments${c.id}.xml`, COMMENTS_TYPE);
  }
  for (const d of drawingEmits) {
    addOverride(manifest, `/xl/drawings/drawing${d.id}.xml`, DRAWING_TYPE);
  }
  for (const c of chartEmits) {
    addOverride(manifest, `/xl/charts/chart${c.id}.xml`, c.isCx ? CHARTEX_TYPE : CHART_TYPE);
  }
  for (const us of userShapeEmits) {
    addOverride(manifest, `/xl/drawings/chartDrawing${us.id}.xml`, DRAWING_TYPE);
  }
  if (wb.vbaProject) {
    addDefault(manifest, 'bin', 'application/vnd.ms-office.vbaProject');
  }
  if (wb.passthrough) {
    for (const path of wb.passthrough.keys()) {
      const ct = wb.passthroughContentTypes?.get(path);
      if (ct !== undefined) addOverride(manifest, `/${path}`, ct);
    }
  }
  // Each unique image extension gets a Default entry (`<Default Extension="png"
  // ContentType="image/png"/>`).
  for (const ext of imageExts) {
    const fmt = (Object.entries(IMAGE_FORMAT_EXTENSION).find(([, e]) => e === ext) ?? [])[0] as
      | XlsxImageFormat
      | undefined;
    if (fmt) addDefault(manifest, ext, IMAGE_FORMAT_MIME[fmt]);
  }
  if (wb.properties) addOverride(manifest, `/${ARC_CORE}`, CORE_PROPS_TYPE);
  if (wb.appProperties) addOverride(manifest, `/${ARC_APP}`, EXT_PROPS_TYPE);
  if (wb.customProperties) addOverride(manifest, `/${ARC_CUSTOM}`, CPROPS_TYPE);
  await writer.addEntry(ARC_CONTENT_TYPES, manifestToBytes(manifest));

  // ---- 8. close ----------------------------------------------------------
  await writer.finalize();
}

/**
 * 1 + the highest N among passthrough paths matching `pattern` (whose first
 * capture group is N), so freshly numbered modeled parts never collide with a
 * carried-over one.
 */
function nextFreeIndex(wb: Workbook, pattern: RegExp): number {
  let max = 0;
  if (wb.passthrough) {
    for (const path of wb.passthrough.keys()) {
      const m = pattern.exec(path);
      if (m?.[1] !== undefined) max = Math.max(max, Number.parseInt(m[1], 10));
    }
  }
  return max + 1;
}

/** Serialise the minimum `<workbook><sheets/></workbook>` Excel needs to load a sheet list. */
function serializeWorkbookXml(wb: Workbook, sheetRIds: ReadonlyArray<string>): string {
  const parts: string[] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<workbook xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}">`,
  ];
  if (wb.fileVersion) {
    const fv = serializeFileVersion(wb.fileVersion);
    if (fv) parts.push(fv);
  }
  if (wb.fileSharing) {
    const fs = serializeFileSharing(wb.fileSharing);
    if (fs) parts.push(fs);
  }
  if (wb.workbookXmlExtras?.beforeSheets) {
    for (const node of wb.workbookXmlExtras.beforeSheets) parts.push(serializeChildNode(node));
  }
  // Emit <workbookPr> from the typed model; fall back to a minimal {date1904:
  // true} synthesis so a fresh workbook (no load history) still round-trips
  // through Excel with the right epoch.
  const effectiveWp = effectiveWorkbookProperties(wb);
  if (effectiveWp) {
    const wp = serializeWorkbookProperties(effectiveWp);
    if (wp) parts.push(wp);
  }
  if (wb.workbookProtection) {
    const wp = serializeWorkbookProtection(wb.workbookProtection);
    if (wp) parts.push(wp);
  }
  if (wb.bookViews && wb.bookViews.length > 0) parts.push(serializeBookViews(wb.bookViews));
  parts.push('<sheets>');
  wb.sheets.forEach((ref, i) => {
    const stateAttr = ref.state === 'visible' ? '' : ` state="${ref.state}"`;
    const rId = sheetRIds[i] ?? `rId${i + 1}`;
    parts.push(`<sheet name="${escapeAttr(ref.sheet.title)}" sheetId="${ref.sheetId}"${stateAttr} r:id="${rId}"/>`);
  });
  parts.push('</sheets>');
  // ECMA-376 §18.2.27 makes CT_Workbook an xsd:sequence, so the order below is
  // normative: functionGroups, externalReferences, definedNames, calcPr,
  // oleSize, customWorkbookViews, pivotCaches, smartTagPr, smartTagTypes,
  // webPublishing, fileRecoveryPr, webPublishObjects, extLst. Excel rejects a
  // workbook whose children are out of sequence. `afterSheets` is the
  // unmodeled tail (webPublishing / webPublishObjects / extLst /
  // mc:AlternateContent), so it goes last.
  if (wb.functionGroups) {
    const fg = serializeFunctionGroups(wb.functionGroups);
    if (fg) parts.push(fg);
  }
  if (wb.externalReferences && wb.externalReferences.length > 0) {
    const inner: string[] = ['<externalReferences>'];
    for (const er of wb.externalReferences) {
      inner.push(`<externalReference r:id="${escapeAttr(er.rId)}"/>`);
    }
    inner.push('</externalReferences>');
    parts.push(inner.join(''));
  }
  if (wb.definedNames.length > 0) {
    parts.push('<definedNames>');
    for (const dn of wb.definedNames) {
      let attrs = ` name="${escapeAttr(dn.name)}"`;
      if (dn.scope !== undefined) attrs += ` localSheetId="${dn.scope}"`;
      if (dn.hidden) attrs += ' hidden="1"';
      if (dn.comment !== undefined) attrs += ` comment="${escapeAttr(dn.comment)}"`;
      parts.push(`<definedName${attrs}>${escapeText(dn.value)}</definedName>`);
    }
    parts.push('</definedNames>');
  }
  if (wb.calcProperties) {
    const cp = serializeCalcProperties(wb.calcProperties);
    if (cp) parts.push(cp);
  }
  // No default `<calcPr>` — Excel handles its absence by inserting one on save.
  // Hand-rolled `calcId` values (especially the modern 191029) made Excel
  // reject workbooks whose future-function formulas (LET / LAMBDA / FILTER…)
  // couldn't be evaluated against the declared engine.
  if (wb.oleSize !== undefined) {
    parts.push(`<oleSize ref="${escapeAttr(wb.oleSize)}"/>`);
  }
  if (wb.customWorkbookViews && wb.customWorkbookViews.length > 0) {
    parts.push(serializeCustomWorkbookViews(wb.customWorkbookViews));
  }
  if (wb.pivotCaches && wb.pivotCaches.length > 0) {
    const inner: string[] = ['<pivotCaches>'];
    for (const pc of wb.pivotCaches) {
      inner.push(`<pivotCache cacheId="${pc.cacheId}" r:id="${escapeAttr(pc.rId)}"/>`);
    }
    inner.push('</pivotCaches>');
    parts.push(inner.join(''));
  }
  if (wb.smartTagPr) {
    const stp = serializeSmartTagPr(wb.smartTagPr);
    if (stp) parts.push(stp);
  }
  if (wb.smartTagTypes && wb.smartTagTypes.length > 0) {
    parts.push(serializeSmartTagTypes(wb.smartTagTypes));
  }
  if (wb.fileRecoveryPr) {
    const fp = serializeFileRecoveryPr(wb.fileRecoveryPr);
    if (fp) parts.push(fp);
  }
  if (wb.workbookXmlExtras?.afterSheets) {
    for (const node of wb.workbookXmlExtras.afterSheets) parts.push(serializeChildNode(node));
  }
  parts.push('</workbook>');
  return parts.join('');
}

function serializeCalcProperties(
  cp: import('../workbook/calc-properties').CalcProperties,
): string | undefined {
  let attrs = '';
  if (cp.calcId !== undefined) attrs += ` calcId="${cp.calcId}"`;
  if (cp.calcMode !== undefined) attrs += ` calcMode="${cp.calcMode}"`;
  if (cp.fullCalcOnLoad !== undefined) attrs += ` fullCalcOnLoad="${cp.fullCalcOnLoad ? '1' : '0'}"`;
  if (cp.refMode !== undefined) attrs += ` refMode="${cp.refMode}"`;
  if (cp.iterate !== undefined) attrs += ` iterate="${cp.iterate ? '1' : '0'}"`;
  if (cp.iterateCount !== undefined) attrs += ` iterateCount="${cp.iterateCount}"`;
  if (cp.iterateDelta !== undefined) attrs += ` iterateDelta="${cp.iterateDelta}"`;
  if (cp.fullPrecision !== undefined) attrs += ` fullPrecision="${cp.fullPrecision ? '1' : '0'}"`;
  if (cp.calcCompleted !== undefined) attrs += ` calcCompleted="${cp.calcCompleted ? '1' : '0'}"`;
  if (cp.calcOnSave !== undefined) attrs += ` calcOnSave="${cp.calcOnSave ? '1' : '0'}"`;
  if (cp.concurrentCalc !== undefined) attrs += ` concurrentCalc="${cp.concurrentCalc ? '1' : '0'}"`;
  if (cp.concurrentManualCount !== undefined) attrs += ` concurrentManualCount="${cp.concurrentManualCount}"`;
  if (cp.forceFullCalc !== undefined) attrs += ` forceFullCalc="${cp.forceFullCalc ? '1' : '0'}"`;
  if (attrs.length === 0) return undefined;
  return `<calcPr${attrs}/>`;
}

function serializeFunctionGroups(
  fg: import('../workbook/function-groups').FunctionGroups,
): string | undefined {
  let attrs = '';
  if (fg.builtInGroupCount !== undefined) attrs += ` builtInGroupCount="${fg.builtInGroupCount}"`;
  if (fg.groups.length === 0 && attrs.length === 0) return undefined;
  if (fg.groups.length === 0) return `<functionGroups${attrs}/>`;
  const inner: string[] = [`<functionGroups${attrs}>`];
  for (const g of fg.groups) inner.push(`<functionGroup name="${escapeAttr(g.name)}"/>`);
  inner.push('</functionGroups>');
  return inner.join('');
}

function serializeSmartTagPr(
  stp: import('../workbook/smart-tags').SmartTagProperties,
): string | undefined {
  let attrs = '';
  if (stp.embed !== undefined) attrs += ` embed="${stp.embed ? '1' : '0'}"`;
  if (stp.show !== undefined) attrs += ` show="${stp.show}"`;
  if (attrs.length === 0) return undefined;
  return `<smartTagPr${attrs}/>`;
}

function serializeSmartTagTypes(
  tags: ReadonlyArray<import('../workbook/smart-tags').SmartTagType>,
): string {
  const inner: string[] = ['<smartTagTypes>'];
  for (const t of tags) {
    let attrs = '';
    if (t.namespaceUri !== undefined) attrs += ` namespaceUri="${escapeAttr(t.namespaceUri)}"`;
    if (t.name !== undefined) attrs += ` name="${escapeAttr(t.name)}"`;
    if (t.url !== undefined) attrs += ` url="${escapeAttr(t.url)}"`;
    inner.push(`<smartTagType${attrs}/>`);
  }
  inner.push('</smartTagTypes>');
  return inner.join('');
}

function serializeFileRecoveryPr(
  fp: import('../workbook/file-recovery').FileRecoveryProperties,
): string | undefined {
  let attrs = '';
  if (fp.autoRecover !== undefined) attrs += ` autoRecover="${fp.autoRecover ? '1' : '0'}"`;
  if (fp.crashSave !== undefined) attrs += ` crashSave="${fp.crashSave ? '1' : '0'}"`;
  if (fp.dataExtractLoad !== undefined) attrs += ` dataExtractLoad="${fp.dataExtractLoad ? '1' : '0'}"`;
  if (fp.repairLoad !== undefined) attrs += ` repairLoad="${fp.repairLoad ? '1' : '0'}"`;
  if (attrs.length === 0) return undefined;
  return `<fileRecoveryPr${attrs}/>`;
}

function serializeFileSharing(
  fs: import('../workbook/file-sharing').FileSharing,
): string | undefined {
  let attrs = '';
  if (fs.readOnlyRecommended !== undefined)
    attrs += ` readOnlyRecommended="${fs.readOnlyRecommended ? '1' : '0'}"`;
  if (fs.userName !== undefined) attrs += ` userName="${escapeAttr(fs.userName)}"`;
  if (fs.reservationPassword !== undefined)
    attrs += ` reservationPassword="${escapeAttr(fs.reservationPassword)}"`;
  if (fs.algorithmName !== undefined) attrs += ` algorithmName="${escapeAttr(fs.algorithmName)}"`;
  if (fs.hashValue !== undefined) attrs += ` hashValue="${escapeAttr(fs.hashValue)}"`;
  if (fs.saltValue !== undefined) attrs += ` saltValue="${escapeAttr(fs.saltValue)}"`;
  if (fs.spinCount !== undefined) attrs += ` spinCount="${fs.spinCount}"`;
  if (attrs.length === 0) return undefined;
  return `<fileSharing${attrs}/>`;
}

function serializeFileVersion(
  fv: import('../workbook/file-version').FileVersion,
): string | undefined {
  let attrs = '';
  if (fv.appName !== undefined) attrs += ` appName="${escapeAttr(fv.appName)}"`;
  if (fv.lastEdited !== undefined) attrs += ` lastEdited="${escapeAttr(fv.lastEdited)}"`;
  if (fv.lowestEdited !== undefined) attrs += ` lowestEdited="${escapeAttr(fv.lowestEdited)}"`;
  if (fv.rupBuild !== undefined) attrs += ` rupBuild="${escapeAttr(fv.rupBuild)}"`;
  if (fv.codeName !== undefined) attrs += ` codeName="${escapeAttr(fv.codeName)}"`;
  if (attrs.length === 0) return undefined;
  return `<fileVersion${attrs}/>`;
}

function effectiveWorkbookProperties(
  wb: Workbook,
): import('../workbook/workbook-properties').WorkbookProperties | undefined {
  const explicit = wb.workbookProperties;
  if (explicit) {
    // Mirror the canonical date1904 flag if the typed model omits it.
    if (wb.date1904 && explicit.date1904 === undefined) {
      return { ...explicit, date1904: true };
    }
    return explicit;
  }
  if (wb.date1904) return { date1904: true };
  return undefined;
}

function serializeWorkbookProperties(
  wp: import('../workbook/workbook-properties').WorkbookProperties,
): string | undefined {
  let attrs = '';
  const boolKeys: ReadonlyArray<keyof import('../workbook/workbook-properties').WorkbookProperties> = [
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
  ];
  for (const k of boolKeys) {
    const v = wp[k];
    if (v !== undefined) attrs += ` ${k}="${v ? '1' : '0'}"`;
  }
  if (wp.showObjects !== undefined) attrs += ` showObjects="${wp.showObjects}"`;
  if (wp.updateLinks !== undefined) attrs += ` updateLinks="${wp.updateLinks}"`;
  if (wp.codeName !== undefined) attrs += ` codeName="${escapeAttr(wp.codeName)}"`;
  if (wp.defaultThemeVersion !== undefined) attrs += ` defaultThemeVersion="${wp.defaultThemeVersion}"`;
  if (attrs.length === 0) return undefined;
  return `<workbookPr${attrs}/>`;
}

function serializeCustomWorkbookViews(
  views: ReadonlyArray<import('../workbook/views').CustomWorkbookView>,
): string {
  const parts: string[] = ['<customWorkbookViews>'];
  for (const v of views) {
    let attrs = ` name="${escapeAttr(v.name)}" guid="${escapeAttr(v.guid)}"`;
    if (v.autoUpdate !== undefined) attrs += ` autoUpdate="${v.autoUpdate ? '1' : '0'}"`;
    if (v.mergeInterval !== undefined) attrs += ` mergeInterval="${v.mergeInterval}"`;
    if (v.changesSavedWin !== undefined) attrs += ` changesSavedWin="${v.changesSavedWin ? '1' : '0'}"`;
    if (v.onlySync !== undefined) attrs += ` onlySync="${v.onlySync ? '1' : '0'}"`;
    if (v.personalView !== undefined) attrs += ` personalView="${v.personalView ? '1' : '0'}"`;
    if (v.includePrintSettings !== undefined)
      attrs += ` includePrintSettings="${v.includePrintSettings ? '1' : '0'}"`;
    if (v.includeHiddenRowCol !== undefined)
      attrs += ` includeHiddenRowCol="${v.includeHiddenRowCol ? '1' : '0'}"`;
    if (v.maximized !== undefined) attrs += ` maximized="${v.maximized ? '1' : '0'}"`;
    if (v.minimized !== undefined) attrs += ` minimized="${v.minimized ? '1' : '0'}"`;
    if (v.showHorizontalScroll !== undefined)
      attrs += ` showHorizontalScroll="${v.showHorizontalScroll ? '1' : '0'}"`;
    if (v.showVerticalScroll !== undefined)
      attrs += ` showVerticalScroll="${v.showVerticalScroll ? '1' : '0'}"`;
    if (v.showSheetTabs !== undefined) attrs += ` showSheetTabs="${v.showSheetTabs ? '1' : '0'}"`;
    if (v.xWindow !== undefined) attrs += ` xWindow="${v.xWindow}"`;
    if (v.yWindow !== undefined) attrs += ` yWindow="${v.yWindow}"`;
    attrs += ` windowWidth="${v.windowWidth}" windowHeight="${v.windowHeight}"`;
    if (v.tabRatio !== undefined) attrs += ` tabRatio="${v.tabRatio}"`;
    attrs += ` activeSheetId="${v.activeSheetId}"`;
    if (v.showFormulaBar !== undefined) attrs += ` showFormulaBar="${v.showFormulaBar ? '1' : '0'}"`;
    if (v.showStatusbar !== undefined) attrs += ` showStatusbar="${v.showStatusbar ? '1' : '0'}"`;
    if (v.showComments !== undefined) attrs += ` showComments="${v.showComments}"`;
    if (v.showObjects !== undefined) attrs += ` showObjects="${v.showObjects}"`;
    parts.push(`<customWorkbookView${attrs}/>`);
  }
  parts.push('</customWorkbookViews>');
  return parts.join('');
}

function serializeBookViews(views: ReadonlyArray<import('../workbook/views').WorkbookView>): string {
  const parts: string[] = ['<bookViews>'];
  for (const v of views) {
    let attrs = '';
    if (v.visibility !== undefined) attrs += ` visibility="${v.visibility}"`;
    if (v.minimized !== undefined) attrs += ` minimized="${v.minimized ? '1' : '0'}"`;
    if (v.showHorizontalScroll !== undefined)
      attrs += ` showHorizontalScroll="${v.showHorizontalScroll ? '1' : '0'}"`;
    if (v.showVerticalScroll !== undefined)
      attrs += ` showVerticalScroll="${v.showVerticalScroll ? '1' : '0'}"`;
    if (v.showSheetTabs !== undefined) attrs += ` showSheetTabs="${v.showSheetTabs ? '1' : '0'}"`;
    if (v.xWindow !== undefined) attrs += ` xWindow="${v.xWindow}"`;
    if (v.yWindow !== undefined) attrs += ` yWindow="${v.yWindow}"`;
    if (v.windowWidth !== undefined) attrs += ` windowWidth="${v.windowWidth}"`;
    if (v.windowHeight !== undefined) attrs += ` windowHeight="${v.windowHeight}"`;
    if (v.tabRatio !== undefined) attrs += ` tabRatio="${v.tabRatio}"`;
    if (v.firstSheet !== undefined) attrs += ` firstSheet="${v.firstSheet}"`;
    if (v.activeTab !== undefined) attrs += ` activeTab="${v.activeTab}"`;
    if (v.autoFilterDateGrouping !== undefined)
      attrs += ` autoFilterDateGrouping="${v.autoFilterDateGrouping ? '1' : '0'}"`;
    parts.push(`<workbookView${attrs}/>`);
  }
  parts.push('</bookViews>');
  return parts.join('');
}

function serializeWorkbookProtection(
  wp: import('../workbook/protection').WorkbookProtection,
): string | undefined {
  let attrs = '';
  const strAttrs = [
    'workbookPassword',
    'workbookPasswordCharacterSet',
    'workbookAlgorithmName',
    'workbookHashValue',
    'workbookSaltValue',
    'revisionsPassword',
    'revisionsPasswordCharacterSet',
    'revisionsAlgorithmName',
    'revisionsHashValue',
    'revisionsSaltValue',
  ] as const;
  for (const k of strAttrs) {
    const v = wp[k];
    if (v !== undefined) attrs += ` ${k}="${escapeAttr(v)}"`;
  }
  if (wp.workbookSpinCount !== undefined) attrs += ` workbookSpinCount="${wp.workbookSpinCount}"`;
  if (wp.revisionsSpinCount !== undefined) attrs += ` revisionsSpinCount="${wp.revisionsSpinCount}"`;
  if (wp.lockStructure !== undefined) attrs += ` lockStructure="${wp.lockStructure ? '1' : '0'}"`;
  if (wp.lockWindows !== undefined) attrs += ` lockWindows="${wp.lockWindows ? '1' : '0'}"`;
  if (wp.lockRevision !== undefined) attrs += ` lockRevision="${wp.lockRevision ? '1' : '0'}"`;
  if (attrs.length === 0) return undefined;
  return `<workbookProtection${attrs}/>`;
}

// Local aliases keep the dense call sites below readable. Canonical
// implementations live in utils/escape.ts so all three writers (this file,
// xml/serializer, xml/stream-writer) agree on how `>` / whitespace are
// handled.
const escapeText = escapeXmlText;
const escapeAttr = escapeXmlAttr;

/**
 * Serialise an XmlNode child of `<workbook>` for inline injection back into the
 * workbook XML stream. Reuses serializeXml then strips the declaration.
 * Captured nodes carry Clark-notation names so namespace prefixes get
 * reallocated by serializeXml — Excel tolerates the extra `xmlns="…"`
 * declarations on each captured root.
 */
function serializeChildNode(node: import('../xml/tree').XmlNode): string {
  const bytes = serializeXmlNode(node, { xmlDeclaration: false });
  return new TextDecoder().decode(bytes);
}
