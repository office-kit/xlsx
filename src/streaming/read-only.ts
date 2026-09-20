// Streaming read-only workbook.
//
// `loadWorkbookStream` opens the zip + parses workbook.xml / sharedStrings.xml
// / styles.xml metadata up front (small even on million-row archives), then
// exposes a lazy `openWorksheet(name)` that SAX-iterates the sheet body via
// `iterParse`. The iterator streams rows without materialising the full sheet
// in memory.

import { normalizeStrictArchive } from '../io/strict.js';
import { makeSharedStrings, parseSharedStringsXml, type SharedStringsTable } from '../workbook/shared-strings.js';
import { ARC_CONTENT_TYPES, ARC_ROOT_RELS, localNameOf } from '../xml/namespaces.js';
import { findById, findByType, makeRelationships, relsFromBytes } from '../packaging/relationships.js';
import { manifestFromBytes } from '../packaging/manifest.js';
import { parseCellNumber } from '../utils/cell-number.js';
import {
  chargeCell,
  chargeRow,
  type ContentBudget,
  type ContentLimits,
  makeContentBudget,
  resolveContentLimits,
  type ResolvedContentLimits,
} from '../worksheet/content-budget.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import type { DecompressionLimits } from '../zip/decompression-guard.js';
import { type ZipArchive, openZip } from '../zip/reader.js';
import type { CellValue, ExcelErrorCode } from '../cell/cell.js';
import { unescapeCellString } from '../utils/escape.js';
import { ERROR_CODES } from '../utils/inference.js';
import { parseXsdBoolean } from '../utils/xsd-boolean.js';
import { iterParse, type SaxEvent, type SaxInput } from '../xml/iterparse.js';
import { parseXml } from '../xml/parser.js';
import { assertNotStrictRelTypes, assertNotStrictRoot } from '../xml/strict-package.js';
import type { XlsxSource } from '../io/source.js';
import { coordinateToTuple, derivedRowNumber, MAX_ROW, rowNumberFromAttr } from '../utils/coordinate.js';
import { type Stylesheet, makeStylesheet } from '../styles/stylesheet.js';
import { parseStylesheetXml } from '../styles/stylesheet-reader.js';
import {
  OFFICE_DOC_REL_TYPE,
  parseDate1904,
  parseSheetEntries,
  readOptionalWorkbookPart,
  resolveRelTarget,
  SHARED_STRINGS_PART,
  STYLES_PART,
  WORKBOOK_TAG,
} from '../io/load.js';

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

const relsPathFor = (partPath: string): string => {
  const i = partPath.lastIndexOf('/');
  if (i < 0) return `_rels/${partPath}.rels`;
  return `${partPath.slice(0, i)}/_rels/${partPath.slice(i + 1)}.rels`;
};

const decodeCellValue = (
  t: string,
  vText: string | undefined,
  inlineText: string | undefined,
  sst: ReadonlyArray<string>,
  sheet: string,
  col: number,
  row: number,
): CellValue => {
  switch (t) {
    case 'n':
      // Throws on text or an exponent past the double range, as loadWorkbook
      // does: the two entry points have to answer the same bytes the same way,
      // and a null here would be indistinguishable from an empty cell.
      return parseCellNumber(vText, sheet, col, row);
    case 's': {
      if (vText === undefined) return null;
      const idx = Number.parseInt(vText, 10);
      if (!Number.isInteger(idx) || idx < 0 || idx >= sst.length) return null;
      return sst[idx] ?? null;
    }
    case 'b':
      // Invalid boolean values stay empty in this reader, like out-of-range
      // shared-string indexes. Numeric values have stricter validation.
      return parseXsdBoolean(vText) ?? null;
    case 'e': {
      if (!vText || !ERROR_CODES.has(vText)) return null;
      return { kind: 'error', code: vText as ExcelErrorCode };
    }
    case 'str':
      return vText ?? '';
    case 'inlineStr':
      return inlineText ?? '';
    default:
      // An unhandled `t` (`"d"`, or something not in ST_CellType at all). The
      // type says nothing about what the text holds, so there is no finiteness
      // rule to apply; loadWorkbook rejects the cell type outright instead.
      return vText !== undefined && vText !== '' ? Number.parseFloat(vText) : null;
  }
};

/**
 * SAX-iterate `<sheetData>/<row>/<c>` events out of the worksheet bytes (or a
 * stream that yields them), yielding one `ReadOnlyCell[]` per row that matches
 * `opts`.
 *
 * `readSheetData` in `../worksheet/reader.ts` walks the same element shapes for
 * `loadWorkbook`. The two stay separate because this one is an async generator
 * over a stream with a row band and an early exit, and that per-event cost is
 * what the other one exists to avoid. Where a row or cell with no `@r` lands
 * has to come out the same in both, and is asserted to: see "agrees with
 * loadWorkbook on where every cell lands" in
 * `tests/worksheet/row-without-r-attribute.test.ts`.
 */
async function* iterSheetRows(
  title: string,
  sheetInput: SaxInput,
  sst: ReadonlyArray<string>,
  opts: IterRowsOptions,
  contentLimits: ResolvedContentLimits,
): AsyncIterableIterator<ReadOnlyCell[]> {
  // One budget per traversal, not per workbook: this reader holds a row at a
  // time and the same sheet can be iterated again, so a cumulative total would
  // refuse the second pass over a workbook it already allowed.
  const budget = makeContentBudget(contentLimits);
  const minRow = opts.minRow ?? 1;
  const maxRow = opts.maxRow ?? Number.POSITIVE_INFINITY;
  const minCol = opts.minCol ?? 1;
  const maxCol = opts.maxCol ?? Number.POSITIVE_INFINITY;

  let inSheetData = false;
  // -1: outside a <row>. 0: inside one that carried no `@r`, whose number the
  // first located cell settles. `nextRow` is the high-water mark a row with
  // neither an `@r` nor a located cell falls back to.
  let currentRow = -1;
  let nextRow = 1;
  let currentCells: ReadOnlyCell[] = [];
  let nextCol = 1;
  let pendingCells: Array<{ col: number; type: string; text: string; inline: string; styleId: number }> = [];
  const settleRow = (row: number): void => {
    currentRow = row;
    nextRow = Math.max(nextRow, row + 1);
    if (row >= minRow && row <= maxRow) {
      for (const cell of pendingCells) {
        currentCells.push({ row, col: cell.col, value: decodeCellValue(cell.type, cell.text, cell.inline, sst, title, cell.col, row), styleId: cell.styleId });
      }
    }
    pendingCells = [];
  };

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

  let checkedRoot = false;
  for await (const ev of iterParse(sheetInput)) {
    const e = ev as SaxEvent;
    if (e.kind === 'start') {
      if (!checkedRoot) {
        assertNotStrictRoot(e.name);
        checkedRoot = true;
      }
      const local = localNameOf(e.name);
      if (!inSheetData) {
        if (local === 'sheetData') inSheetData = true;
        continue;
      }
      switch (local) {
        case 'row': {
          // `@r` is optional on CT_Row (ECMA-376 §18.3.1.73); such a row waits
          // for its first cell to name it.
          const rRaw = e.attrs['r'];
          if (rRaw === undefined) {
            chargeRow(budget, title, undefined);
            currentRow = 0;
          } else {
            currentRow = rowNumberFromAttr(rRaw, 'loadWorkbookStream');
            nextRow = Math.max(nextRow, currentRow + 1);
            // Charged here rather than at `</row>` so a row past the cap is
            // refused before its cells are decoded. Unnumbered rows are
            // charged at their opening tag too, before buffering any cells.
            chargeRow(budget, title, currentRow);
          }
          currentCells = [];
          pendingCells = [];
          nextCol = 1;
          break;
        }
        case 'c': {
          if (currentRow < 0) break;
          const ref = e.attrs['r'];
          if (currentRow === 0 && ref) settleRow(coordinateToTuple(ref).row);
          // Skip cell-attr parsing entirely when the row is outside the
          // requested band — saves the parseInt + coordinateToTuple hit on
          // every cell of every excluded row.
          if (currentRow !== 0 && (currentRow < minRow || currentRow > maxRow)) break;
          cellType = e.attrs['t'] ?? 'n';
          const sRaw = e.attrs['s'];
          cellStyleId = sRaw ? Number.parseInt(sRaw, 10) || 0 : 0;
          if (ref) {
            const tup = coordinateToTuple(ref);
            cellRow = tup.row;
            cellCol = tup.col;
          } else {
            cellRow = currentRow;
            cellCol = nextCol;
          }
          nextCol = cellCol + 1;
          cellOpen = cellCol >= minCol && cellCol <= maxCol;
          if (cellOpen) {
            // An unresolved row might be selected later, so its buffered cells
            // consume budget even if its eventual number is outside the band.
            chargeCell(budget, title, cellCol, cellRow === 0 ? undefined : cellRow);
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
    const local = localNameOf(e.name);
    if (!inSheetData) continue;
    switch (local) {
      case 'sheetData':
        inSheetData = false;
        return;
      case 'row': {
        // A row that held no located cell still consumes a slot, so settle it
        // before moving the high-water mark past it.
        if (currentRow === 0) {
          settleRow(derivedRowNumber(nextRow, 'loadWorkbookStream'));
        }
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
        currentCells = [];
        break;
      }
      case 'c': {
        if (cellOpen && cellRow === 0 && cellCol >= minCol && cellCol <= maxCol) {
          pendingCells.push({ col: cellCol, type: cellType, text: vText, inline: isText, styleId: cellStyleId });
        } else if (cellOpen && cellCol >= minCol && cellCol <= maxCol && cellRow >= minRow && cellRow <= maxRow) {
          const value = decodeCellValue(cellType, vText, isText, sst, title, cellCol, cellRow);
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
}

const isXmlSpace = (b: number | undefined): boolean =>
  b === 0x20 /* sp */ || b === 0x09 /* tab */ || b === 0x0a /* lf */ || b === 0x0d /* cr */;

/**
 * Read the `r="N"` row number out of a `<row …>` attribute region, scanning
 * bytes directly. Returns -1 when the attribute is absent, or holds anything
 * `parseRowNumberAttr` rejects: the two accept the same shapes, so the index
 * can never seek by a number the SAX walk reads differently.
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
    let q = p + 4;
    while (q < to && isXmlSpace(bytes[q])) q++;
    if (bytes[q] === 0x2b /* + */) q++;
    let value = 0;
    let digits = 0;
    while (q < to) {
      const d = bytes[q];
      if (d === undefined || d < 0x30 || d > 0x39) break;
      value = value * 10 + (d - 0x30);
      digits++;
      q++;
    }
    while (q < to && isXmlSpace(bytes[q])) q++;
    // A non-numeric, out-of-range or unterminated `r` is not the row ref; keep
    // scanning the rest of the region the way the equivalent regex would have.
    if (digits > 0 && value >= 1 && value <= MAX_ROW && q < to && bytes[q] === 0x22 /* '"' */) return value;
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
 *
 * `hasUnnumberedRow` reports a `<row>` this scan cannot number, which makes the
 * whole sheet unseekable: a row without `@r` takes the row its first cell
 * names, and reading cell refs is the SAX walk's job.
 *
 * Every recorded row is charged to `budget`. The index holds one object per
 * row of the whole part, however narrow the band that asked for it, so a cap
 * that only counted the rows a band yields would leave the largest allocation
 * on this path uncapped.
 */
const buildRowOffsetIndex = (
  bytes: Uint8Array,
  budget: ContentBudget,
  title: string,
): {
  index: ReadonlyArray<{ row: number; offset: number }>;
  sheetDataEnd: number;
  sheetDataTagEnd: number;
  hasUnnumberedRow: boolean;
} => {
  const out: Array<{ row: number; offset: number }> = [];
  let hasUnnumberedRow = false;
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
    if (row > 0) {
      chargeRow(budget, title, row);
      out.push({ row, offset: start });
    } else hasUnnumberedRow = true;
    i = j + 1;
  }
  if (sheetDataEnd < 0) sheetDataEnd = bytes.length;
  return { index: out, sheetDataEnd, sheetDataTagEnd: findSheetDataTagEnd(bytes), hasUnnumberedRow };
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

/**
 * A worksheet is seekable when every one of its rows carries an `@r` the byte
 * scan can read. Otherwise the scan's own numbering would disagree with the SAX
 * walk's, so band queries stream the part instead, and nothing is retained.
 */
type SheetSeek =
  | { seekable: false }
  | ({ seekable: true; bytes: Uint8Array } & Omit<ReturnType<typeof buildRowOffsetIndex>, 'hasUnnumberedRow'>);

interface RowIndexCache {
  byWorksheet: WeakMap<object, SheetSeek>;
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
 * queries jump straight to the first matching row without inflating or scanning
 * again. A sheet whose rows omit `@r` cannot be indexed by row number, so its
 * band queries keep streaming and hold nothing.
 */
const makeStreamingReadOnlyWorksheet = (
  title: string,
  archive: ZipArchive,
  partPath: string,
  sst: ReadonlyArray<string>,
  indexes: RowIndexCache,
  contentLimits: ResolvedContentLimits,
): ReadOnlyWorksheet => {
  // Lazy + cached, bytes included: the archive only keeps small entries, so a
  // second band query that went back to `read` would inflate the whole part
  // again. The index adds one small object per row; the part is held from the first
  // band query until the worksheet handle goes away or the workbook closes.
  // This avoids parsing or inflating the whole part on every query.
  const cacheKey = {};
  const ensureIndexed = (): SheetSeek => {
    let cached = indexes.byWorksheet.get(cacheKey);
    if (!cached) {
      const bytes = archive.read(partPath);
      // The index is built once and reused, so it gets a budget of its own
      // rather than spending a traversal's.
      const { hasUnnumberedRow, index, sheetDataEnd, sheetDataTagEnd } = buildRowOffsetIndex(
        bytes,
        makeContentBudget(contentLimits),
        title,
      );
      cached = hasUnnumberedRow
        ? { seekable: false }
        : { seekable: true, bytes, index, sheetDataEnd, sheetDataTagEnd };
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
      return iterSheetRows(title, archive.readStream(partPath), sst, opts, contentLimits);
    }
    // Band query (minRow > 1): the row-offset index needs the full inflated
    // bytes so we can binary-search to the byte offset of the first matching
    // row. Materialise once and reuse via `ensureIndexed`.
    const seek = ensureIndexed();
    // Rows the index cannot number: walk the part instead of seeking into it,
    // which is the only way the derived numbers stay the SAX walk's.
    if (!seek.seekable) return iterSheetRows(title, archive.readStream(partPath), sst, opts, contentLimits);
    const { bytes, index, sheetDataEnd, sheetDataTagEnd } = seek;
    if (index.length === 0 || sheetDataTagEnd < 0) return iterSheetRows(title, bytes, sst, opts, contentLimits);
    const pos = firstRowAtOrAfter(index, minRow);
    if (pos < 0) {
      // Every row is below minRow, so there is nothing to yield.
      return (async function* () {})();
    }
    const target = index[pos];
    if (!target) return iterSheetRows(title, bytes, sst, opts, contentLimits);
    return iterSheetRows(
      title,
      replayFromRow(bytes, sheetDataTagEnd, target.offset, sheetDataEnd),
      sst,
      opts,
      contentLimits,
    );
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
  partPathByName: ReadonlyMap<string, string>,
  sst: ReadonlyArray<string>,
  contentLimits: ResolvedContentLimits,
): ReadOnlyWorkbook => {
  // Weak keys let unused worksheet handles release their indexed bytes. The
  // indirection also lets close() release every index while handles remain live.
  const indexes: RowIndexCache = { byWorksheet: new WeakMap() };
  return {
    sheetNames,
    styles,
    date1904,
    openWorksheet(name) {
      const partPath = partPathByName.get(name);
      if (partPath === undefined) {
        throw new OpenXmlSchemaError(`loadWorkbookStream: no worksheet named "${name}"`);
      }
      return makeStreamingReadOnlyWorksheet(name, archive, partPath, sst, indexes, contentLimits);
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
  /**
   * Caps on how much content one row-iteration will read, the same option
   * `loadWorkbook` takes. Unlimited by default, and counted per traversal here
   * rather than per workbook: this reader holds a row at a time, so the cap
   * bounds how long a pass can run rather than how much it retains, and a
   * sheet can be iterated again without the cap having been spent. Exceeding
   * either count raises an `OpenXmlContentLimitError`: from the iterator for a
   * row the pass reached, and from the `iterRows` call itself when the index a
   * band query has to build is what passes the cap.
   *
   * A pass is charged for the rows it walks, not the ones it yields. Reaching
   * `minRow` costs a walk or an index over everything before it, and the index
   * a band query builds holds one object per row of the whole part, so that is
   * where a cap has to bite for `iterRows({ minRow })` to be bounded at all.
   * The index is built once per worksheet handle and charged once. Cells are
   * charged before buffering or decoding them, including cells of unnumbered
   * rows whose eventual row number might fall outside the requested band.
   * Cells in known excluded rows or columns are not charged.
   *
   * See `LoadOptions.contentLimits` for a starting ingestion profile and
   * `SECURITY.md` for the threat model.
   */
  contentLimits?: ContentLimits;
}

/** Open an xlsx for read-only streaming access. */
export async function loadWorkbookStream(
  source: XlsxSource,
  opts: LoadWorkbookStreamOptions = {},
): Promise<ReadOnlyWorkbook> {
  // Settled here rather than inside the row iterator: a cap that cannot mean
  // anything has to be reported from this call, not from an iterator the
  // caller reaches later, and a band query that yields nothing never builds a
  // budget at all.
  const contentLimits = resolveContentLimits(opts.contentLimits);
  const rawArchive = await openZip(
    source,
    opts.decompressionLimits === undefined ? {} : { decompressionLimits: opts.decompressionLimits },
  );
  try {
    const archive = normalizeStrictArchive(rawArchive);
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
    const officeDocRel = findByType(rootRels, OFFICE_DOC_REL_TYPE);
    if (!officeDocRel) {
      assertNotStrictRelTypes(rootRels.rels.map((r) => r.type));
      throw new OpenXmlSchemaError(`loadWorkbookStream: no officeDocument relationship in root rels`);
    }
    const workbookPath = resolveRelTarget('', officeDocRel.target);
    if (!archive.has(workbookPath)) {
      throw new OpenXmlSchemaError(`loadWorkbookStream: workbook part "${workbookPath}" missing`);
    }
    const workbookRoot = parseXml(archive.read(workbookPath));
    if (workbookRoot.name !== WORKBOOK_TAG) {
      assertNotStrictRoot(workbookRoot.name);
      throw new OpenXmlSchemaError(
        `loadWorkbookStream: ${workbookPath} root is "${workbookRoot.name}", expected workbook`,
      );
    }
    // loadWorkbook's `<sheets>` parser, and its rejections below: a declaration
    // that loader refuses must not read here as a workbook without that sheet.
    const declaredSheets = parseSheetEntries(workbookRoot);
    const wbRelsPath = relsPathFor(workbookPath);
    if (declaredSheets.length > 0 && !archive.has(wbRelsPath)) {
      throw new OpenXmlSchemaError(
        `loadWorkbookStream: workbook has sheets but rels part "${wbRelsPath}" is missing`,
      );
    }
    const wbRels = archive.has(wbRelsPath) ? relsFromBytes(archive.read(wbRelsPath)) : makeRelationships();
    const partPathByName = new Map<string, string>();
    for (const declared of declaredSheets) {
      if (partPathByName.has(declared.name)) {
        throw new OpenXmlSchemaError(`loadWorkbookStream: duplicate sheet name "${declared.name}"`);
      }
      const rel = findById(wbRels, declared.rId);
      if (!rel) {
        throw new OpenXmlSchemaError(
          `loadWorkbookStream: sheet "${declared.name}" rId "${declared.rId}" has no matching rels entry`,
        );
      }
      partPathByName.set(declared.name, resolveRelTarget(workbookPath, rel.target));
    }

    const sstBytes = readOptionalWorkbookPart(archive, workbookPath, wbRels, SHARED_STRINGS_PART);
    const sst: SharedStringsTable = sstBytes === undefined ? makeSharedStrings() : parseSharedStringsXml(sstBytes);
    const stylesBytes = readOptionalWorkbookPart(archive, workbookPath, wbRels, STYLES_PART);
    const styles: Stylesheet = stylesBytes === undefined ? makeStylesheet() : parseStylesheetXml(stylesBytes);

    return makeStreamingReadOnlyWorkbook(
      declaredSheets.map((e) => e.name),
      styles,
      parseDate1904(workbookRoot),
      archive,
      partPathByName,
      sst.entries.map((e) => (typeof e === 'string' ? e : e.runs.map((r) => r.text).join(''))),
      contentLimits,
    );
  } catch (cause) {
    rawArchive.close();
    throw cause;
  }
}
