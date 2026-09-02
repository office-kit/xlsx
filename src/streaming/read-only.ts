// Streaming read-only workbook.
//
// `loadWorkbookStream` opens the zip + parses workbook.xml / sharedStrings.xml
// / styles.xml metadata up front (small even on million-row archives), then
// exposes a lazy `openWorksheet(name)` that SAX-iterates the sheet body via
// `iterParse`. The iterator streams rows without materialising the full sheet
// in memory.

import { makeSharedStrings, parseSharedStringsXml, type SharedStringsTable } from '../workbook/shared-strings';
import { ARC_CONTENT_TYPES, ARC_ROOT_RELS, ARC_SHARED_STRINGS, ARC_STYLE, REL_NS, SHEET_MAIN_NS } from '../xml/namespaces';
import { findById, relsFromBytes } from '../packaging/relationships';
import { manifestFromBytes } from '../packaging/manifest';
import { OpenXmlSchemaError } from '../utils/exceptions';
import type { DecompressionLimits } from '../zip/decompression-guard';
import { type ZipArchive, openZip } from '../zip/reader';
import type { CellValue, ExcelErrorCode } from '../cell/cell';
import { ERROR_CODES } from '../utils/inference';
import { iterParse, type SaxEvent, type SaxInput } from '../xml/iterparse';
import { parseXml } from '../xml/parser';
import { findChild, findChildren } from '../xml/tree';
import type { XlsxSource } from '../io/source';
import { coordinateToTuple } from '../utils/coordinate';
import { type Stylesheet, makeStylesheet } from '../styles/stylesheet';
import { parseStylesheetXml } from '../styles/stylesheet-reader';
import { resolveRelTarget } from '../io/load';

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
  openWorksheet(name: string): ReadOnlyWorksheet;
  close(): Promise<void>;
}

interface SheetEntry {
  name: string;
  rId: string;
  partPath: string;
}

const parseSheetList = (workbookXml: Uint8Array, workbookPath: string, archive: ZipArchive): SheetEntry[] => {
  const root = parseXml(workbookXml);
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
          if (inIs) inIsT = true;
          break;
        default:
          break;
      }
      continue;
    }
    if (e.kind === 'text') {
      if (inV) vText += e.text;
      else if (inIsT) isText += e.text;
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
        if (inIs) inIsT = false;
        break;
      default:
        break;
    }
  }
  // Avoid unused-var lint when row attrs never touched.
  void currentRowAttrs;
}

/**
 * Build a sorted `[rowNum, byteOffset]` index for every `<row r="N">`
 * occurrence in a worksheet's bytes. Pure byte-level scan (no SAX), cheap
 * relative to the per-cell SAX walk: ~50 ns per row on M-series Node 22.
 *
 * `sheetDataEnd` is the byte offset of `</sheetData>` so callers can clip the
 * slice that gets handed to saxes.
 */
const buildRowOffsetIndex = (
  bytes: Uint8Array,
): { index: ReadonlyArray<{ row: number; offset: number }>; sheetDataEnd: number } => {
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
    const attrsBuf = bytes.subarray(start + 4, j);
    const attrs = new TextDecoder('ascii', { fatal: false }).decode(attrsBuf);
    const m = /\sr="(\d+)"/.exec(attrs);
    if (m?.[1]) {
      const row = Number.parseInt(m[1], 10);
      if (Number.isInteger(row)) out.push({ row, offset: start });
    }
    i = j + 1;
  }
  if (sheetDataEnd < 0) sheetDataEnd = bytes.length;
  return { index: out, sheetDataEnd };
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
 * Slice a worksheet's bytes to start at the row at index `idxPos` of the
 * row-offset index, wrapping the result with a synthetic `<sheetData>` envelope
 * so saxes parses it in the right namespace.
 */
const SHEET_DATA_OPEN = `<?xml version="1.0" encoding="UTF-8"?><sheetData xmlns="${SHEET_MAIN_NS}">`;
const SHEET_DATA_CLOSE = `</sheetData>`;
const sliceFromRow = (
  bytes: Uint8Array,
  fromOffset: number,
  sheetDataEnd: number,
): Uint8Array => {
  const prefix = new TextEncoder().encode(SHEET_DATA_OPEN);
  const suffix = new TextEncoder().encode(SHEET_DATA_CLOSE);
  const middle = bytes.subarray(fromOffset, sheetDataEnd);
  const out = new Uint8Array(prefix.length + middle.length + suffix.length);
  out.set(prefix, 0);
  out.set(middle, prefix.length);
  out.set(suffix, prefix.length + middle.length);
  return out;
};

/**
 * Factory: build a {@link ReadOnlyWorksheet} bound to a single worksheet part
 * inside an opened archive. SAX iteration runs lazily — `iterRows` re-reads the
 * part bytes each time so the caller can iterate the same sheet repeatedly
 * without keeping a buffered decoder around.
 *
 * For `iterRows({ minRow > 1 })`, a row-offset index is built lazily on first
 * use and cached; subsequent band queries jump straight to the byte offset of
 * the first matching row instead of SAX-walking the entire `<sheetData>`.
 */
const makeStreamingReadOnlyWorksheet = (
  title: string,
  archive: ZipArchive,
  partPath: string,
  sst: ReadonlyArray<string>,
): ReadOnlyWorksheet => {
  // Lazy + cached. The index is small (~16 B per row); for 1M rows that's 16 MB
  // of working set, vs. the alternative of walking the sheet bytes through
  // saxes on every band query.
  let cached: ReturnType<typeof buildRowOffsetIndex> | undefined;
  const ensureIndex = (bytes: Uint8Array) => {
    if (!cached) cached = buildRowOffsetIndex(bytes);
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
    // row. Materialise once and reuse via `ensureIndex`.
    const bytes = archive.read(partPath);
    const { index, sheetDataEnd } = ensureIndex(bytes);
    if (index.length === 0) return iterSheetRows(bytes, sst, opts);
    const pos = firstRowAtOrAfter(index, minRow);
    if (pos < 0) {
      // Every row is below minRow — nothing to yield.
      return (async function* () {})();
    }
    const target = index[pos];
    if (!target) return iterSheetRows(bytes, sst, opts);
    const sliced = sliceFromRow(bytes, target.offset, sheetDataEnd);
    return iterSheetRows(sliced, sst, opts);
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
  archive: ZipArchive,
  entries: ReadonlyMap<string, SheetEntry>,
  sst: ReadonlyArray<string>,
): ReadOnlyWorkbook => ({
  sheetNames,
  styles,
  openWorksheet(name) {
    const entry = entries.get(name);
    if (!entry) {
      throw new OpenXmlSchemaError(`loadWorkbookStream: no worksheet named "${name}"`);
    }
    return makeStreamingReadOnlyWorksheet(name, archive, entry.partPath, sst);
  },
  async close() {
    archive.close();
  },
});

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
  const sheetEntries = parseSheetList(archive.read(workbookPath), workbookPath, archive);
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
    archive,
    entryMap,
    sst.entries.map((e) => (typeof e === 'string' ? e : e.runs.map((r) => r.text).join(''))),
  );
}
