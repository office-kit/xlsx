// Streaming read-only workbook.
//
// `loadWorkbookStream` opens the zip + parses workbook.xml / sharedStrings.xml
// / styles.xml metadata up front (small even on million-row archives), then
// exposes a lazy `openWorksheet(name)` that SAX-iterates the sheet body via
// `iterParse`. The iterator streams rows without materialising the full sheet
// in memory.

import { makeSharedStrings, parseSharedStringsXml, type SharedStringsTable } from '../workbook/shared-strings.js';
import { ARC_CONTENT_TYPES, ARC_ROOT_RELS, ARC_SHARED_STRINGS, ARC_STYLE, REL_NS, SHEET_MAIN_NS } from '../xml/namespaces.js';
import { findById, relsFromBytes } from '../packaging/relationships.js';
import { manifestFromBytes } from '../packaging/manifest.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import type { DecompressionLimits } from '../zip/decompression-guard.js';
import { type ZipArchive, openZip } from '../zip/reader.js';
import type { CellValue, ExcelErrorCode } from '../cell/cell.js';
import { unescapeCellString } from '../utils/escape.js';
import { ERROR_CODES } from '../utils/inference.js';
import { iterParse, type SaxEvent, type SaxInput } from '../xml/iterparse.js';
import { parseXml } from '../xml/parser.js';
import { findChild, findChildren, type XmlNode } from '../xml/tree.js';
import type { XlsxSource } from '../io/source.js';
import { coordinateToTuple } from '../utils/coordinate.js';
import { type Stylesheet, makeStylesheet } from '../styles/stylesheet.js';
import { parseStylesheetXml } from '../styles/stylesheet-reader.js';
import { parseDate1904, resolveRelTarget } from '../io/load.js';

const SHEET_TAG = `{${SHEET_MAIN_NS}}sheet`;
const SHEETS_TAG = `{${SHEET_MAIN_NS}}sheets`;

export interface IterRowsOptions {
  minRow?: number;
  maxRow?: number;
  minCol?: number;
  maxCol?: number;
}

export interface ReadOnlyCell {
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
  readonly styleId: number;
}

export interface ReadOnlyWorksheet {
  title: string;
  iterRows(opts?: IterRowsOptions): AsyncIterableIterator<ReadOnlyCell[]>;
  iterValues(opts?: IterRowsOptions): AsyncIterableIterator<CellValue[]>;
}

export interface ReadOnlyWorkbook {
  sheetNames: string[];
  styles: Stylesheet;
  /** Date1904 mode toggles between Excel's two epoch systems. */
  date1904: boolean;
  openWorksheet(name: string): ReadOnlyWorksheet;
  close(): Promise<void>;
}

interface SheetEntry {
  name: string;
  rId: string;
  partPath: string;
}

const parseSheetList = (root: XmlNode, workbookPath: string, archive: ZipArchive): SheetEntry[] => {
  const sheetsEl = findChild(root, SHEETS_TAG);
  if (!sheetsEl) return [];
  const wbRelsPath = relsPathFor(workbookPath);
  const wbRels = archive.has(wbRelsPath) ? relsFromBytes(archive.read(wbRelsPath)) : { rels: [] };
  const out: SheetEntry[] = [];
  for (const sheet of findChildren(sheetsEl, SHEET_TAG)) {
    const name = sheet.attrs['name'];
    const rId = sheet.attrs[`{${REL_NS}}id`];
    if (!name || !rId) continue;
    const rel = findById(wbRels, rId);
    if (!rel) continue;
    const partPath = resolveRelTarget(workbookPath, rel.target);
    out.push({ name, rId, partPath });
  }
  return out;
};

const relsPathFor = (partPath: string): string => {
  const i = partPath.lastIndexOf('/');
  if (i < 0) return `_rels/${partPath}.rels`;
  return `${partPath.slice(0, i)}/_rels/${partPath.slice(i + 1)}.rels`;
};

const localName = (qname: string): string => {
  const i = qname.lastIndexOf('}');
  return i < 0 ? qname : qname.slice(i + 1);
};

const decodeCellValue = (
  t: string,
  vText: string | undefined,
  inlineText: string | undefined,
  sst: ReadonlyArray<string>,
): CellValue => {
  switch (t) {
    case 'n':
      return vText !== undefined && vText !== '' ? Number.parseFloat(vText) : null;
    case 's': {
      if (vText === undefined) return null;
      const idx = Number.parseInt(vText, 10);
      if (!Number.isInteger(idx) || idx < 0 || idx >= sst.length) return null;
      return sst[idx] ?? null;
    }
    case 'b':
      return vText === '1';
    case 'e': {
      if (!vText || !ERROR_CODES.has(vText)) return null;
      return { kind: 'error', code: vText as ExcelErrorCode };
    }
    case 'str':
      return vText ?? '';
    case 'inlineStr':
      return inlineText ?? '';
    default:
      return vText !== undefined && vText !== '' ? Number.parseFloat(vText) : null;
  }
};

/**
 * SAX-iterate `<sheetData>/<row>/<c>` events out of the worksheet bytes (or a
 * stream that yields them), yielding one `ReadOnlyCell[]` per row that matches
 * `opts`.
 */
async function* iterSheetRows(
  sheetInput: SaxInput,
  sst: ReadonlyArray<string>,
  opts: IterRowsOptions,
): AsyncIterableIterator<ReadOnlyCell[]> {
  const minRow = opts.minRow ?? 1;
  const maxRow = opts.maxRow ?? Number.POSITIVE_INFINITY;
  const minCol = opts.minCol ?? 1;
  const maxCol = opts.maxCol ?? Number.POSITIVE_INFINITY;

  let inSheetData = false;
  let currentRow = -1;
  let currentRowAttrs: Record<string, string> | null = null;
  let currentCells: ReadOnlyCell[] = [];

  // Per-cell state. Reset when each <c> starts.
  let cellOpen = false;
  let cellRow = 0;
  let cellCol = 0;
  let cellType = 'n';
  let cellStyleId = 0;
  let inV = false;
  let vText = '';
  let inIs = false;
  let inIsT = false;
  let isText = '';
  let isRunText = '';

  for await (const ev of iterParse(sheetInput)) {
    const e = ev as SaxEvent;
    if (e.kind === 'start') {
      const local = localName(e.name);
      if (!inSheetData) {
        if (local === 'sheetData') inSheetData = true;
        continue;
      }
      switch (local) {
        case 'row': {
          currentRowAttrs = e.attrs;
          const rRaw = e.attrs['r'];
          currentRow = rRaw ? Number.parseInt(rRaw, 10) : currentRow + 1;
          currentCells = [];
          break;
        }
        case 'c': {
          if (currentRow < 0) break;
          // Skip cell-attr parsing entirely when the row is outside the
          // requested band — saves the parseInt + coordinateToTuple hit on
          // every cell of every excluded row.
          if (currentRow < minRow || currentRow > maxRow) break;
          cellOpen = true;
          cellType = e.attrs['t'] ?? 'n';
          const sRaw = e.attrs['s'];
          cellStyleId = sRaw ? Number.parseInt(sRaw, 10) || 0 : 0;
          const ref = e.attrs['r'];
          if (ref) {
            const tup = coordinateToTuple(ref);
            cellRow = tup.row;
            cellCol = tup.col;
          } else {
            cellRow = currentRow;
            cellCol = (currentCells[currentCells.length - 1]?.col ?? 0) + 1;
          }
          vText = '';
          isText = '';
          break;
        }
        case 'v':
          if (cellOpen) inV = true;
          break;
        case 'is':
          if (cellOpen) inIs = true;
          break;
        case 't':
          if (inIs) {
            inIsT = true;
            isRunText = '';
          }
          break;
        default:
          break;
      }
      continue;
    }
    if (e.kind === 'text') {
      if (inV) vText += e.text;
      else if (inIsT) isRunText += e.text;
      continue;
    }
    // end
    const local = localName(e.name);
    if (!inSheetData) continue;
    switch (local) {
      case 'sheetData':
        inSheetData = false;
        return;
      case 'row': {
        if (currentRow >= minRow && currentRow <= maxRow && currentCells.length > 0) {
          yield currentCells;
        }
        // Once we've crossed maxRow there are no more rows to yield — every
        // subsequent <row> would just be parsed and dropped. Stop iterating
        // early. ECMA-376 emits rows in ascending order.
        if (currentRow > maxRow) {
          inSheetData = false;
          return;
        }
        currentRow = -1;
        currentRowAttrs = null;
        currentCells = [];
        break;
      }
      case 'c': {
        if (cellOpen && cellCol >= minCol && cellCol <= maxCol && cellRow >= minRow && cellRow <= maxRow) {
          const value = decodeCellValue(cellType, vText, isText, sst);
          currentCells.push({ row: cellRow, col: cellCol, value, styleId: cellStyleId });
        }
        cellOpen = false;
        break;
      }
      case 'v':
        inV = false;
        break;
      case 'is':
        inIs = false;
        break;
      case 't':
        if (inIsT) {
          // Decode each complete <t>, so SAX chunks can split an escape but
          // adjacent rich-text runs cannot accidentally create one.
          isText += unescapeCellString(isRunText);
          inIsT = false;
        }
        break;
      default:
        break;
    }
  }
  // Avoid unused-var lint when row attrs never touched.
  void currentRowAttrs;
}

const isXmlSpace = (b: number | undefined): boolean =>
  b === 0x20 /* sp */ || b === 0x09 /* tab */ || b === 0x0a /* lf */ || b === 0x0d /* cr */;

/**
 * Read the `r="N"` row number out of a `<row …>` attribute region, scanning
 * bytes directly. Returns -1 when the attribute is absent or malformed.
 *
 * Decoding the region to a string and running a regex instead costs one
 * TextDecoder, one string and one match array per row, which on a million-row
 * sheet dominates the scan this index exists to keep cheap.
 */
const readRowAttr = (bytes: Uint8Array, from: number, to: number): number => {
  for (let p = from; p + 3 < to; p++) {
    if (!isXmlSpace(bytes[p])) continue;
    if (bytes[p + 1] !== 0x72 /* 'r' */ || bytes[p + 2] !== 0x3d /* '=' */ || bytes[p + 3] !== 0x22 /* '"' */) {
      continue;
    }
    let value = 0;
    let digits = 0;
    let q = p + 4;
    while (q < to) {
      const d = bytes[q];
      if (d === undefined || d < 0x30 || d > 0x39) break;
      value = value * 10 + (d - 0x30);
      digits++;
      q++;
    }
    // A non-numeric or unterminated `r` is not the row ref; keep scanning the
    // rest of the region the way the equivalent regex would have.
    if (digits > 0 && q < to && bytes[q] === 0x22 /* '"' */) return value;
  }
  return -1;
};

const isQuote = (b: number | undefined): boolean => b === 0x22 /* dquote */ || b === 0x27 /* squote */;

/**
 * Byte offset just past the worksheet's opening `sheetData` tag, or
 * -1 when there is none. Skips the XML declaration, comments and processing
 * instructions ahead of it, and ignores `>` inside attribute values.
 */
const findSheetDataTagEnd = (bytes: Uint8Array): number => {
  const td = new TextDecoder();
  let i = 0;
  while (i < bytes.length) {
    if (bytes[i] !== 0x3c /* '<' */) {
      i++;
      continue;
    }
    const kind = bytes[i + 1];
    if (kind === 0x3f /* '?' */ || kind === 0x21 /* '!' */) {
      // Comments and processing instructions may contain `>` in their body.
      const isComment = bytes[i + 2] === 0x2d /* '-' */ && bytes[i + 3] === 0x2d;
      i += 2;
      while (i < bytes.length) {
        if (bytes[i] === 0x3e /* '>' */ &&
          (isComment ? bytes[i - 1] === 0x2d && bytes[i - 2] === 0x2d
            : kind !== 0x3f || bytes[i - 1] === 0x3f)) break;
        i++;
      }
      i++;
      continue;
    }
    let nameEnd = i + 1;
    while (nameEnd < bytes.length && !isXmlSpace(bytes[nameEnd]) && bytes[nameEnd] !== 0x3e && bytes[nameEnd] !== 0x2f) nameEnd++;
    const isSheetData = td.decode(bytes.subarray(i + 1, nameEnd)) === 'sheetData';
    let quote = 0;
    let tagEnd = -1;
    for (let j = i + 1; j < bytes.length; j++) {
      const b = bytes[j];
      if (b === undefined) break;
      if (quote !== 0) {
        if (b === quote) quote = 0;
        continue;
      }
      if (isQuote(b)) {
        quote = b;
        continue;
      }
      if (b === 0x3e /* '>' */) {
        tagEnd = j + 1;
        break;
      }
    }
    if (tagEnd < 0 || isSheetData) return tagEnd;
    i = tagEnd;
  }
  return -1;
};

/**
 * Build a sorted `[rowNum, byteOffset]` index for every `<row r="N">`
 * occurrence in a worksheet's bytes. Pure byte-level scan (no SAX), cheap
 * relative to the per-cell SAX walk.
 *
 * `sheetDataEnd` is the byte offset of `</sheetData>` so callers can clip the
 * region that gets handed to saxes. `sheetDataTagEnd` includes the original
 * namespace-bearing ancestor tags, reused by {@link replayFromRow}.
 */
const buildRowOffsetIndex = (
  bytes: Uint8Array,
): { index: ReadonlyArray<{ row: number; offset: number }>; sheetDataEnd: number; sheetDataTagEnd: number } => {
  const out: Array<{ row: number; offset: number }> = [];
  let sheetDataEnd = -1;
  let i = 0;
  while (i < bytes.length) {
    if (bytes[i] !== 0x3c /* '<' */) {
      i++;
      continue;
    }
    // Detect `</sheetData>` once — used to clip the slice fed to saxes.
    if (
      sheetDataEnd < 0 &&
      bytes[i + 1] === 0x2f /* '/' */ &&
      bytes[i + 2] === 0x73 /* 's' */ &&
      bytes[i + 3] === 0x68 /* 'h' */ &&
      bytes[i + 4] === 0x65 /* 'e' */ &&
      bytes[i + 5] === 0x65 /* 'e' */ &&
      bytes[i + 6] === 0x74 /* 't' */ &&
      bytes[i + 7] === 0x44 /* 'D' */ &&
      bytes[i + 8] === 0x61 /* 'a' */ &&
      bytes[i + 9] === 0x74 /* 't' */ &&
      bytes[i + 10] === 0x61 /* 'a' */ &&
      bytes[i + 11] === 0x3e /* '>' */
    ) {
      sheetDataEnd = i;
      break;
    }
    // Match `<row` followed by ASCII whitespace or '>'.
    if (
      bytes[i + 1] !== 0x72 /* 'r' */ ||
      bytes[i + 2] !== 0x6f /* 'o' */ ||
      bytes[i + 3] !== 0x77 /* 'w' */
    ) {
      i++;
      continue;
    }
    const next = bytes[i + 4];
    if (
      next !== 0x20 /* sp */ &&
      next !== 0x09 /* tab */ &&
      next !== 0x0a /* lf */ &&
      next !== 0x0d /* cr */ &&
      next !== 0x3e /* > */ &&
      next !== 0x2f /* / */
    ) {
      i++;
      continue;
    }
    // Walk to the closing '>'; the attrs region carries `r="N"`.
    const start = i;
    let j = i + 4;
    while (j < bytes.length && bytes[j] !== 0x3e) j++;
    if (j >= bytes.length) break;
    const row = readRowAttr(bytes, start + 4, j);
    // Rows are 1-based (ECMA-376 §18.3.1.73), so `r="0"` is not an offset this
    // index can seek to; -1 means there was no usable `r` at all.
    if (row > 0) out.push({ row, offset: start });
    i = j + 1;
  }
  if (sheetDataEnd < 0) sheetDataEnd = bytes.length;
  return { index: out, sheetDataEnd, sheetDataTagEnd: findSheetDataTagEnd(bytes) };
};

/**
 * Binary-search the row index for the first entry with `row >= target`. Returns
 * -1 when every recorded row is below the target.
 */
const firstRowAtOrAfter = (
  index: ReadonlyArray<{ row: number; offset: number }>,
  target: number,
): number => {
  let lo = 0;
  let hi = index.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    const entry = index[mid];
    if (!entry || entry.row < target) lo = mid + 1;
    else hi = mid;
  }
  return lo < index.length ? lo : -1;
};

/**
 * Replay a row band with the original prefix through the opening sheetData
 * tag. Namespace declarations can occur on either worksheet or sheetData;
 * synthesising either tag loses bindings required by the retained rows.
 * Views avoid copying the inflated sheet just to prepend its ancestor tags.
 */
const replayFromRow = (
  bytes: Uint8Array,
  sheetDataTagEnd: number,
  fromOffset: number,
  sheetDataEnd: number,
): ReadableStream<Uint8Array> => {
  const parts = [
    bytes.subarray(0, sheetDataTagEnd),
    bytes.subarray(fromOffset, sheetDataEnd),
    bytes.subarray(sheetDataEnd),
  ];
  let next = 0;
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      const part = parts[next];
      if (part === undefined) {
        controller.close();
        return;
      }
      next++;
      controller.enqueue(part);
    },
  });
};

interface RowIndexCache {
  byWorksheet: WeakMap<object, ReturnType<typeof buildRowOffsetIndex> & { bytes: Uint8Array }>;
}

/**
 * Factory: build a {@link ReadOnlyWorksheet} bound to a single worksheet part
 * inside an opened archive. SAX iteration runs lazily: a whole-sheet `iterRows`
 * streams the part again on every call, so nothing but the inflate window and
 * the SAX state is ever resident.
 *
 * `iterRows({ minRow > 1 })` needs random access instead, so the first band
 * query inflates the part and indexes its row offsets, and holds both for the
 * life of the worksheet handle or until the workbook closes. Subsequent band
 * queries jump straight to the first matching row without inflating or scanning again.
 */
const makeStreamingReadOnlyWorksheet = (
  title: string,
  archive: ZipArchive,
  partPath: string,
  sst: ReadonlyArray<string>,
  indexes: RowIndexCache,
): ReadOnlyWorksheet => {
  // Lazy + cached, bytes included: the archive only keeps small entries, so a
  // second band query that went back to `read` would inflate the whole part
  // again. The index adds one small object per row; the part is held from the first
  // band query until the worksheet handle goes away or the workbook closes.
  // This avoids parsing or inflating the whole part on every query.
  const cacheKey = {};
  const ensureIndexed = () => {
    let cached = indexes.byWorksheet.get(cacheKey);
    if (!cached) {
      const bytes = archive.read(partPath);
      cached = { bytes, ...buildRowOffsetIndex(bytes) };
      indexes.byWorksheet.set(cacheKey, cached);
    }
    return cached;
  };

  const iterRows = (opts: IterRowsOptions = {}): AsyncIterableIterator<ReadOnlyCell[]> => {
    const minRow = opts.minRow ?? 1;
    if (minRow <= 1) {
      // Whole-sheet (or no-min) iter — feed the SAX parser directly off the
      // archive's streaming inflate path so the worksheet's inflated payload
      // is never fully resident. Peak memory for the walk drops to the
      // inflate window + SAX state instead of the entire `<sheetData>` body.
      return iterSheetRows(archive.readStream(partPath), sst, opts);
    }
    // Band query (minRow > 1): the row-offset index needs the full inflated
    // bytes so we can binary-search to the byte offset of the first matching
    // row. Materialise once and reuse via `ensureIndexed`.
    const { bytes, index, sheetDataEnd, sheetDataTagEnd } = ensureIndexed();
    if (index.length === 0 || sheetDataTagEnd < 0) return iterSheetRows(bytes, sst, opts);
    const pos = firstRowAtOrAfter(index, minRow);
    if (pos < 0) {
      // Every row is below minRow, so there is nothing to yield.
      return (async function* () {})();
    }
    const target = index[pos];
    if (!target) return iterSheetRows(bytes, sst, opts);
    return iterSheetRows(replayFromRow(bytes, sheetDataTagEnd, target.offset, sheetDataEnd), sst, opts);
  };
  const iterValues = async function* (opts: IterRowsOptions = {}): AsyncIterableIterator<CellValue[]> {
    for await (const row of iterRows(opts)) {
      yield row.map((c) => c.value);
    }
  };
  return { title, iterRows, iterValues };
};

/**
 * Factory: build a {@link ReadOnlyWorkbook} from an opened archive + pre-parsed
 * sheet list / styles / shared strings. Per the project-wide "no classes" rule
 * (CLAUDE.md), the workbook is a plain object closing over the archive handle.
 */
const makeStreamingReadOnlyWorkbook = (
  sheetNames: string[],
  styles: Stylesheet,
  date1904: boolean,
  archive: ZipArchive,
  entries: ReadonlyMap<string, SheetEntry>,
  sst: ReadonlyArray<string>,
): ReadOnlyWorkbook => {
  // Weak keys let unused worksheet handles release their indexed bytes. The
  // indirection also lets close() release every index while handles remain live.
  const indexes: RowIndexCache = { byWorksheet: new WeakMap() };
  return {
    sheetNames,
    styles,
    date1904,
    openWorksheet(name) {
      const entry = entries.get(name);
      if (!entry) {
        throw new OpenXmlSchemaError(`loadWorkbookStream: no worksheet named "${name}"`);
      }
      return makeStreamingReadOnlyWorksheet(name, archive, entry.partPath, sst, indexes);
    },
    async close() {
      archive.close();
      indexes.byWorksheet = new WeakMap();
    },
  };
};

/** Options for {@link loadWorkbookStream}. */
export interface LoadWorkbookStreamOptions {
  /**
   * Decompression-bomb safeguards applied while inflating zip entries.
   * Defaults to limits that fit any legitimate xlsx; pass `false` to disable
   * (only safe for fully trusted sources).
   */
  decompressionLimits?: DecompressionLimits | false;
}

/** Open an xlsx for read-only streaming access. */
export async function loadWorkbookStream(
  source: XlsxSource,
  opts: LoadWorkbookStreamOptions = {},
): Promise<ReadOnlyWorkbook> {
  const archive = await openZip(
    source,
    opts.decompressionLimits === undefined ? {} : { decompressionLimits: opts.decompressionLimits },
  );
  if (!archive.has(ARC_CONTENT_TYPES)) {
    throw new OpenXmlSchemaError(`loadWorkbookStream: missing "${ARC_CONTENT_TYPES}"`);
  }
  // Manifest parse is intentionally cheap and discarded — we resolve sheets by
  // walking workbook.xml.rels directly.
  manifestFromBytes(archive.read(ARC_CONTENT_TYPES));

  if (!archive.has(ARC_ROOT_RELS)) {
    throw new OpenXmlSchemaError(`loadWorkbookStream: missing "${ARC_ROOT_RELS}"`);
  }
  const rootRels = relsFromBytes(archive.read(ARC_ROOT_RELS));
  const officeDocRel = rootRels.rels.find((r) => r.type === `${REL_NS}/officeDocument`);
  if (!officeDocRel) {
    throw new OpenXmlSchemaError(`loadWorkbookStream: no officeDocument relationship in root rels`);
  }
  const workbookPath = resolveRelTarget('', officeDocRel.target);
  if (!archive.has(workbookPath)) {
    throw new OpenXmlSchemaError(`loadWorkbookStream: workbook part "${workbookPath}" missing`);
  }
  const workbookRoot = parseXml(archive.read(workbookPath));
  const sheetEntries = parseSheetList(workbookRoot, workbookPath, archive);
  const entryMap = new Map<string, SheetEntry>();
  for (const e of sheetEntries) entryMap.set(e.name, e);

  let sst: SharedStringsTable = makeSharedStrings();
  if (archive.has(ARC_SHARED_STRINGS)) {
    sst = parseSharedStringsXml(archive.read(ARC_SHARED_STRINGS));
  }
  let styles: Stylesheet = makeStylesheet();
  if (archive.has(ARC_STYLE)) {
    styles = parseStylesheetXml(archive.read(ARC_STYLE));
  }

  return makeStreamingReadOnlyWorkbook(
    sheetEntries.map((e) => e.name),
    styles,
    parseDate1904(workbookRoot),
    archive,
    entryMap,
    sst.entries.map((e) => (typeof e === 'string' ? e : e.runs.map((r) => r.text).join(''))),
  );
}
