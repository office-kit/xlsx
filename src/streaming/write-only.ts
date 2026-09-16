// Streaming write-only workbook.
//
// `createWriteOnlyWorkbook` lets callers append rows one at a time without
// holding a full Workbook in memory. Each `appendRow` serialises the row
// directly into a string buffer (no `Cell` objects, no Worksheet rows Map) —
// when `close()` runs we glue the rows under a `<sheetData>` envelope and hand
// the bytes to the streaming-deflate ZIP writer. The buffer keeps ~30 bytes per
// cell of XML text instead of the ~200-byte V8 footprint a Cell + Map entry
// costs, so the heap budget at 3M cells drops roughly an order of magnitude vs.
// the previous setCell-based path.

import type { Cell, CellValue } from '../cell/cell.js';
import type { XlsxSink } from '../io/sink.js';
import { addDefault, addOverride, makeManifest, manifestToBytes } from '../packaging/manifest.js';
import { makeRelationships, relsToBytes } from '../packaging/relationships.js';
import {
  addBorder,
  addCellXf,
  addFill,
  addFont,
  addNumFmt,
  type CellXf,
  defaultCellXf,
  makeStylesheet,
  type Stylesheet,
} from '../styles/stylesheet.js';
import { stylesheetToBytes } from '../styles/stylesheet-writer.js';
import type { Alignment } from '../styles/alignment.js';
import type { Border } from '../styles/borders.js';
import type { Fill } from '../styles/fills.js';
import type { Font } from '../styles/fonts.js';
import type { Protection } from '../styles/protection.js';
import { escapeXmlAttr } from '../utils/escape.js';
import { OpenXmlIoError } from '../utils/exceptions.js';
import { utf8ByteLength } from '../utils/utf8.js';
import { makeSharedStrings, sharedStringsToBytes } from '../workbook/shared-strings.js';
import { validateSheetTitle } from '../workbook/workbook.js';
import { serializeCell } from '../worksheet/writer.js';
import {
  ARC_CONTENT_TYPES,
  ARC_ROOT_RELS,
  ARC_SHARED_STRINGS,
  ARC_STYLE,
  ARC_WORKBOOK,
  ARC_WORKBOOK_RELS,
  PKG_REL_NS,
  REL_NS,
  SHARED_STRINGS_TYPE,
  SHEET_MAIN_NS,
  STYLES_TYPE,
  WORKSHEET_TYPE,
  XLSX_TYPE,
} from '../xml/namespaces.js';
import { type CompressionLevel, createZipWriter, type ZipWriterOptions } from '../zip/writer.js';

const escapeAttr = escapeXmlAttr;

export interface WriteOnlyOptions {
  /**
   * Last-modified timestamp for every ZIP entry. Same reproducibility story as
   * `SaveOptions.mtime`: unset, fflate stamps the wall clock per entry and two
   * runs over identical rows differ in bytes.
   */
  mtime?: Date;
  /** Deflate level, 0 (store) to 9 (smallest). Defaults to fflate's own 6. */
  compressionLevel?: CompressionLevel;
}

export interface WriteOnlyStyle {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  numberFormat?: string;
  protection?: Protection;
}

export type WriteOnlyRowItem = CellValue | { value: CellValue; style?: WriteOnlyStyle };

export interface WriteOnlyWorksheet {
  readonly title: string;
  appendRow(row: WriteOnlyRowItem[]): Promise<void>;
  setColumnWidth(col: number, width: number): void;
  close(): Promise<void>;
}

export interface WriteOnlyWorkbook {
  /** Add a new worksheet. The previous worksheet must be `close()`d first. */
  addWorksheet(title: string): Promise<WriteOnlyWorksheet>;
  /** Finalise the archive: emits styles / sharedStrings / workbook / manifest / rels. */
  finalize(): Promise<void>;
  /**
   * Release the sink without finalising the archive. Call this from a
   * surrounding catch block when the producer pipeline throws before
   * `finalize()` — without it, streaming destinations (`toFile` /
   * `toWritable`) keep the file descriptor / Writable open and the partial
   * xlsx looks valid on disk.
   *
   * Idempotent. Subsequent `addWorksheet` / `finalize` calls throw.
   */
  abort(cause?: unknown): void;
}

const validateTitle = (title: string, taken: Set<string>): void => {
  const reason = validateSheetTitle(title);
  if (reason) {
    throw new OpenXmlIoError(`Worksheet title "${title}": ${reason}`);
  }
  if (taken.has(title.toLowerCase())) {
    throw new OpenXmlIoError(`Worksheet title "${title}" is already in use`);
  }
};

/** Allocate a CellXf id for a style spec. Mirrors cell-style.ts but works directly on the pool. */
const allocateXfId = (ss: Stylesheet, style: WriteOnlyStyle): number => {
  // `xfId` is intentionally omitted: leaving it undefined skips the
  // cellStyleXfs bounds check and matches what Excel emits when there is no
  // parent style.
  let xf: CellXf = { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 };
  if (style.font !== undefined) {
    xf = { ...xf, fontId: addFont(ss, style.font), applyFont: true };
  }
  if (style.fill !== undefined) {
    xf = { ...xf, fillId: addFill(ss, style.fill), applyFill: true };
  }
  if (style.border !== undefined) {
    xf = { ...xf, borderId: addBorder(ss, style.border), applyBorder: true };
  }
  if (style.numberFormat !== undefined) {
    xf = { ...xf, numFmtId: addNumFmt(ss, style.numberFormat), applyNumberFormat: true };
  }
  if (style.alignment !== undefined) {
    xf = { ...xf, alignment: style.alignment, applyAlignment: true };
  }
  if (style.protection !== undefined) {
    xf = { ...xf, protection: style.protection, applyProtection: true };
  }
  return addCellXf(ss, xf);
};

interface WorkbookState {
  styles: Stylesheet;
  sst: ReturnType<typeof makeSharedStrings>;
  /** Sheet emit metadata, in addWorksheet order. */
  sheets: Array<{
    title: string;
    sheetId: number;
  }>;
  /** True once finalize() has been called (further mutations throw). */
  finalised: boolean;
  /** True while a worksheet is open (the next addWorksheet must wait). */
  hasOpenWorksheet: boolean;
  /** ZIP writer the workbook + each open worksheet stream chunks through. */
  writer: import('../zip/writer.js').ZipWriter;
}

/**
 * Flush threshold for the worksheet's pending-row text buffer. Smaller values
 * minimise heap; larger values amortise the TextEncoder + push overhead. 64 KB
 * is a sweet spot — heap stays low and per-row JS work is dominated by the
 * actual XML construction, not flushing.
 */
const FLUSH_THRESHOLD_BYTES = 64 * 1024;

/**
 * Factory: build a {@link WriteOnlyWorksheet} that closes over the shared
 * {@link WorkbookState}. Per the project-wide "no classes" rule (CLAUDE.md) the
 * worksheet is a plain object holding the row buffer + column-width map in
 * closure state.
 *
 * The worksheet streams its `<sheetData>` body chunk-by-chunk through the ZIP
 * writer's `addStreamingEntry` API, so the heap footprint stays at one ~64 KB
 * pending text buffer plus deflate scratch — no Cell objects, no all-rows
 * accumulation. The XML envelope (decl / worksheet open / cols / sheetData
 * open) flushes on the first `appendRow` (or `close()` if the sheet is empty);
 * column widths staged via `setColumnWidth` *must* land before the first row.
 */
const makeWriteOnlyWorksheet = (state: WorkbookState, title: string, sheetId: number): WriteOnlyWorksheet => {
  let nextRow = 1;
  let closed = false;
  let headerFlushed = false;
  const columnWidths = new Map<number, number>();
  const dummyCtx = { sharedStrings: state.sst, rels: makeRelationships() };
  const encoder = new TextEncoder();
  const stream = state.writer.addStreamingEntry(`xl/worksheets/sheet${sheetId}.xml`);
  let pendingText = '';
  let pendingBytes = 0;

  const writeText = (text: string): void => {
    pendingText += text;
    // UTF-8 byte length, computed without a full encode: scan once and add the
    // exact codepoint cost. `text.length` would undercount CJK by ~3× and let
    // the buffer balloon past FLUSH_THRESHOLD_BYTES on Japanese / Chinese
    // workloads.
    pendingBytes += utf8ByteLength(text);
    if (pendingBytes >= FLUSH_THRESHOLD_BYTES) {
      stream.write(encoder.encode(pendingText));
      pendingText = '';
      pendingBytes = 0;
    }
  };

  const flushHeader = (): void => {
    if (headerFlushed) return;
    headerFlushed = true;
    let header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    header += `<worksheet xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}">`;
    if (columnWidths.size > 0) {
      header += '<cols>';
      const sorted = [...columnWidths.entries()].sort((a, b) => a[0] - b[0]);
      for (const [col, width] of sorted) {
        header += `<col min="${col}" max="${col}" width="${width}" customWidth="1"/>`;
      }
      header += '</cols>';
    }
    header += '<sheetData>';
    writeText(header);
  };

  const appendRow = async (row: WriteOnlyRowItem[]): Promise<void> => {
    if (closed) throw new OpenXmlIoError('appendRow: worksheet already closed');
    flushHeader();
    const r = nextRow++;
    let xml = `<row r="${r}">`;
    for (let i = 0; i < row.length; i++) {
      const item = row[i];
      if (item === undefined || item === null) continue;
      const col = i + 1;
      let value: CellValue;
      let style: WriteOnlyStyle | undefined;
      if (item !== null && typeof item === 'object' && 'value' in (item as object)) {
        const wrapped = item as { value: CellValue; style?: WriteOnlyStyle };
        value = wrapped.value;
        style = wrapped.style;
      } else {
        value = item as CellValue;
      }
      const styleId = style ? allocateXfId(state.styles, style) : 0;
      // Ephemeral cell-shaped object — discarded as soon as serializeCell
      // returns its `<c .../>` string. Keeps the heap footprint at the size of
      // the pending text buffer instead of a full Worksheet model.
      const cell: Cell = { row: r, col, value, styleId };
      xml += serializeCell(cell, dummyCtx);
    }
    xml += '</row>';
    writeText(xml);
  };

  const setColumnWidth = (col: number, width: number): void => {
    if (closed) throw new OpenXmlIoError('setColumnWidth: worksheet already closed');
    if (headerFlushed) {
      throw new OpenXmlIoError(
        'setColumnWidth: must be called before the first appendRow — column widths are emitted as part of the worksheet header',
      );
    }
    columnWidths.set(col, width);
  };

  const close = async (): Promise<void> => {
    if (closed) return;
    closed = true;
    flushHeader();
    writeText('</sheetData></worksheet>');
    if (pendingText.length > 0) {
      stream.write(encoder.encode(pendingText));
      pendingText = '';
      pendingBytes = 0;
    }
    await stream.end();
    state.sheets.push({ title, sheetId });
    state.hasOpenWorksheet = false;
  };

  return { title, appendRow, setColumnWidth, close };
};

/**
 * Factory: build a {@link WriteOnlyWorkbook} from a sink. State lives in a
 * closure rather than on a class instance per the project-wide "no classes"
 * rule (CLAUDE.md).
 */
const makeWriteOnlyWorkbook = (sink: XlsxSink, zipOpts: ZipWriterOptions): WriteOnlyWorkbook => {
  const styles = makeStylesheet();
  // Reserve cellXfs[0] for the default (no apply* flags). Unstyled cells point
  // at this slot via styleId=0; user-styled cells start at index 1 so the
  // writer emits an `s="N"` attribute for them.
  addCellXf(styles, defaultCellXf());
  // The ZIP writer is created up front: each addWorksheet opens a streaming
  // entry on it and flushes row chunks through the deflate stream as they
  // arrive. Sheets emit before styles / sst / workbook / rels / content-types
  // so the writer can serialise them in order.
  const state: WorkbookState = {
    styles,
    sst: makeSharedStrings(),
    sheets: [],
    finalised: false,
    hasOpenWorksheet: false,
    writer: createZipWriter(sink, zipOpts),
  };

  const addWorksheet = async (title: string): Promise<WriteOnlyWorksheet> => {
    if (state.finalised) {
      throw new OpenXmlIoError('addWorksheet: workbook already finalised');
    }
    if (state.hasOpenWorksheet) {
      throw new OpenXmlIoError(
        'addWorksheet: previous worksheet still open — call close() before opening the next one',
      );
    }
    const taken = new Set(state.sheets.map((s) => s.title.toLowerCase()));
    validateTitle(title, taken);
    state.hasOpenWorksheet = true;
    const sheetId = state.sheets.length + 1;
    return makeWriteOnlyWorksheet(state, title, sheetId);
  };

  const finalize = async (): Promise<void> => {
    if (state.finalised) {
      throw new OpenXmlIoError('finalize: already finalised');
    }
    if (state.hasOpenWorksheet) {
      throw new OpenXmlIoError('finalize: a worksheet is still open — call close() before finalising');
    }
    state.finalised = true;
    const writer = state.writer;
    try {
      await finalizeImpl(state, writer);
    } catch (err) {
      // Release the sink so the half-emitted archive doesn't linger as a
      // valid-looking file on disk. abort() is idempotent.
      writer.abort(err);
      throw err;
    }
  };

  const abort = (cause?: unknown): void => {
    if (state.finalised) return;
    state.finalised = true;
    state.writer.abort(cause);
  };

  return { addWorksheet, finalize, abort };
};

async function finalizeImpl(state: WorkbookState, writer: WorkbookState['writer']): Promise<void> {
  // 1. Worksheets — already streamed through writer.addStreamingEntry
  // during each WriteOnlyWorksheet's appendRow / close cycle.

  // 2. Stylesheet.
  await writer.addEntry(ARC_STYLE, stylesheetToBytes(state.styles));

  // 3. SharedStrings (only when non-empty).
  if (state.sst.entries.length > 0) {
    await writer.addEntry(ARC_SHARED_STRINGS, sharedStringsToBytes(state.sst));
  }

  // 4. workbook.xml.
  const workbookXml = serializeWorkbookXml(state.sheets);
  await writer.addEntry(ARC_WORKBOOK, new TextEncoder().encode(workbookXml));

  // 5. workbook.xml.rels.
  const wbRels = makeRelationships();
  state.sheets.forEach((s, i) => {
    wbRels.rels.push({
      id: `rId${i + 1}`,
      type: `${REL_NS}/worksheet`,
      target: `worksheets/sheet${s.sheetId}.xml`,
    });
  });
  if (state.sst.entries.length > 0) {
    wbRels.rels.push({
      id: `rId${wbRels.rels.length + 1}`,
      type: `${REL_NS}/sharedStrings`,
      target: 'sharedStrings.xml',
    });
  }
  wbRels.rels.push({
    id: `rId${wbRels.rels.length + 1}`,
    type: `${REL_NS}/styles`,
    target: 'styles.xml',
  });
  await writer.addEntry(ARC_WORKBOOK_RELS, relsToBytes(wbRels));

  // 6. root rels.
  const rootRels = makeRelationships();
  rootRels.rels.push({
    id: 'rId1',
    type: `${REL_NS}/officeDocument`,
    target: 'xl/workbook.xml',
  });
  await writer.addEntry(ARC_ROOT_RELS, relsToBytes(rootRels));
  void PKG_REL_NS; // imported for future docProps support

  // 7. [Content_Types].xml.
  const manifest = makeManifest();
  // Excel rejects packages whose [Content_Types].xml is missing the Default
  // entries for `rels` / `xml` — without them the package relationships file
  // can't be classified and Excel refuses to open.
  addDefault(manifest, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');
  addDefault(manifest, 'xml', 'application/xml');
  addOverride(manifest, `/${ARC_WORKBOOK}`, XLSX_TYPE);
  for (const s of state.sheets) {
    addOverride(manifest, `/xl/worksheets/sheet${s.sheetId}.xml`, WORKSHEET_TYPE);
  }
  addOverride(manifest, `/${ARC_STYLE}`, STYLES_TYPE);
  if (state.sst.entries.length > 0) {
    addOverride(manifest, `/${ARC_SHARED_STRINGS}`, SHARED_STRINGS_TYPE);
  }
  await writer.addEntry(ARC_CONTENT_TYPES, manifestToBytes(manifest));

  await writer.finalize();
}

const serializeWorkbookXml = (
  sheets: ReadonlyArray<{ title: string; sheetId: number }>,
): string => {
  const parts: string[] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<workbook xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}">`,
    '<sheets>',
  ];
  sheets.forEach((s, i) => {
    parts.push(
      `<sheet name="${escapeAttr(s.title)}" sheetId="${s.sheetId}" r:id="rId${i + 1}"/>`,
    );
  });
  parts.push('</sheets></workbook>');
  return parts.join('');
};

/** Open a workbook for streaming write-only output. */
export async function createWriteOnlyWorkbook(
  sink: XlsxSink,
  opts: WriteOnlyOptions = {},
): Promise<WriteOnlyWorkbook> {
  // The streaming-deflate ZIP writer is constructed eagerly here: each
  // addWorksheet opens an entry on it and flushes row chunks through fflate's
  // `Zip` + `ZipDeflate` immediately, so peak memory stays at one pending row
  // buffer plus deflate scratch (no all-sheets accumulation).
  return makeWriteOnlyWorkbook(sink, opts);
}

