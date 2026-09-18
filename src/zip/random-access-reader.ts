// Random-access ZIP reader.
//
// Handing every entry to `fflate.unzipSync` materialises every uncompressed
// payload at once, so a 100 MB xlsx with 500 MB of decompressed sheet data
// spikes the resident set accordingly. This reader keeps only the compressed
// archive bytes resident, parses the central directory once (cheap, about 46 B
// per entry plus the filename), and inflates each entry lazily on `read(path)`.
//
// Small inflated payloads are kept for re-reads (see ./inflate-cache.ts).
// Keeping every one of them would put the whole uncompressed package back in
// memory, which is the cost this reader exists to avoid, so that cache is
// bounded in both entry size and total and anything outside those bounds
// inflates again on the next read.
//
// Limitations:
// - ZIP64 reads only when the standard ZIP32 fields fit. EOCD with
// sentinel values (0xFFFF / 0xFFFFFFFF) falls back to fflate's `unzipSync` so
// external ZIP64 archives still load (the writer side has its own ZIP32 cap
// guard).
// - Compression methods: STORE (0) and DEFLATE (8). Anything else
// throws OpenXmlIoError.

import { Inflate, unzipSync } from 'fflate';
import { OpenXmlDecompressionBombError, OpenXmlIoError } from '../utils/exceptions.js';
import {
  beginEntryInflate,
  checkDeclaredTotals,
  createBudget,
  type DecompressionBudget,
  type DecompressionLimitsInput,
  entryInflateCap,
  entryOverflowError,
  resolveDecompressionLimits,
} from './decompression-guard.js';
import { createInflateCache } from './inflate-cache.js';
import type { ZipArchive } from './reader.js';

/**
 * Chunk size used when feeding compressed bytes into fflate's `Inflate` for
 * streaming reads. 64 KB matches the saxes/SAX consumer's typical batch and
 * keeps peak transient memory bounded.
 */
const INFLATE_CHUNK_BYTES = 64 * 1024;

/**
 * `decode()` without `{ stream: true }` resets the decoder's state on every
 * call, so one instance is safe to share across entries and across archives.
 */
const CD_NAME_DECODER = new TextDecoder('utf-8');

const singleChunkStream = (bytes: Uint8Array): ReadableStream<Uint8Array> =>
  new ReadableStream<Uint8Array>({
    start(controller) {
      if (bytes.byteLength > 0) controller.enqueue(bytes);
      controller.close();
    },
  });

const SIG_EOCD = 0x06054b50;
const SIG_CD = 0x02014b50;
const SIG_LFH = 0x04034b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_EOCD_LOCATOR = 0x07064b50;
/** Size of a zip End-of-Central-Directory record with no trailing comment. */
const EOCD_MIN_BYTES = 22;
const ZIP32_MAX_U16 = 0xffff;
const ZIP32_MAX_U32 = 0xffffffff;
const ZIP64_LOCATOR_SIZE = 20;
const ZIP64_EXTRA_HEADER_ID = 0x0001;
const COMP_STORE = 0;
const COMP_DEFLATE = 8;

interface CdEntry {
  path: string;
  lfhOffset: number;
  compMethod: number;
  compSize: number;
  uncompSize: number;
}

const u16 = (b: Uint8Array, off: number): number => (b[off] ?? 0) | ((b[off + 1] ?? 0) << 8);
const u32 = (b: Uint8Array, off: number): number => {
  const v0 = b[off] ?? 0;
  const v1 = b[off + 1] ?? 0;
  const v2 = b[off + 2] ?? 0;
  const v3 = b[off + 3] ?? 0;
  return (v0 | (v1 << 8) | (v2 << 16) | (v3 << 24)) >>> 0;
};

// Read a little-endian 64-bit value as a JS Number. Safe for values up to
// 2^53-1 (Number.MAX_SAFE_INTEGER ~ 9 PiB), which is well past any realistic
// xlsx archive — we throw if a parsed value exceeds the safe-integer range.
const u64 = (b: Uint8Array, off: number): number => {
  const lo = u32(b, off);
  const hi = u32(b, off + 4);
  if (hi > 0x1fffff) {
    throw new OpenXmlIoError(
      `openZip: ZIP64 field at byte ${off} exceeds the safe-integer range (${hi}*2^32 + ${lo})`,
    );
  }
  return hi * 0x100000000 + lo;
};

/** Find the End-of-Central-Directory record by scanning backwards from EOF. */
function findEocd(b: Uint8Array): number {
  const minStart = Math.max(0, b.length - EOCD_MIN_BYTES - 0xffff);
  for (let i = b.length - EOCD_MIN_BYTES; i >= minStart; i--) {
    if (u32(b, i) === SIG_EOCD) return i;
  }
  throw new OpenXmlIoError('openZip: no End-of-Central-Directory signature found');
}

interface CdSummary {
  totalEntries: number;
  cdSize: number;
  cdOffset: number;
}

/**
 * Resolve the central-directory totals. When the EOCD carries ZIP64 sentinel
 * values (0xFFFF / 0xFFFFFFFF) the real totals live in a ZIP64 EOCD record
 * located via the ZIP64 EOCD locator just before the regular EOCD. ECMA-376
 * xlsx archives stay in ZIP32 territory; the ZIP64 path exists so the
 * decompression-bomb guards still apply to large external archives instead of
 * falling all the way back to `unzipSync` (which produces every entry's bytes
 * before any cap can fire).
 */
function readCdSummary(b: Uint8Array, eocdOff: number): CdSummary {
  let totalEntries = u16(b, eocdOff + 10);
  let cdSize = u32(b, eocdOff + 12);
  let cdOffset = u32(b, eocdOff + 16);

  const usesZip64Eocd =
    totalEntries === ZIP32_MAX_U16 || cdSize === ZIP32_MAX_U32 || cdOffset === ZIP32_MAX_U32;
  if (!usesZip64Eocd) {
    return { totalEntries, cdSize, cdOffset };
  }

  // Locate the ZIP64 EOCD locator. It sits immediately before the regular
  // EOCD when one is present.
  const locatorOff = eocdOff - ZIP64_LOCATOR_SIZE;
  if (locatorOff < 0 || u32(b, locatorOff) !== SIG_ZIP64_EOCD_LOCATOR) {
    throw new OpenXmlIoError('openZip: ZIP64 EOCD locator missing despite EOCD sentinel values');
  }
  const zip64EocdOff = u64(b, locatorOff + 8);
  if (zip64EocdOff < 0 || zip64EocdOff + 56 > b.length) {
    throw new OpenXmlIoError(`openZip: ZIP64 EOCD offset ${zip64EocdOff} out of bounds`);
  }
  if (u32(b, zip64EocdOff) !== SIG_ZIP64_EOCD) {
    throw new OpenXmlIoError(`openZip: ZIP64 EOCD signature missing at byte ${zip64EocdOff}`);
  }
  totalEntries = u64(b, zip64EocdOff + 32);
  cdSize = u64(b, zip64EocdOff + 40);
  cdOffset = u64(b, zip64EocdOff + 48);
  return { totalEntries, cdSize, cdOffset };
}

/** Walk the ZIP64 Extended Information extra field for sentinel-valued sizes/offset. */
function readZip64Extra(
  b: Uint8Array,
  extraStart: number,
  extraLen: number,
  wantsUncompSize: boolean,
  wantsCompSize: boolean,
  wantsLfhOffset: boolean,
  entryPath: string,
): { uncompSize?: number; compSize?: number; lfhOffset?: number } {
  let p = extraStart;
  const end = extraStart + extraLen;
  while (p + 4 <= end) {
    const id = u16(b, p);
    const size = u16(b, p + 2);
    const dataStart = p + 4;
    const next = dataStart + size;
    if (next > end) break;
    if (id === ZIP64_EXTRA_HEADER_ID) {
      // Sanity-check that the declared extra-field size actually covers every
      // u64 the CD asked us to read. A truncated / mis-sized ZIP64 extra
      // field would otherwise let u64() interpret bytes belonging to the
      // next extra record (or the next CD entry) as a size/offset, producing
      // bogus values that downstream bomb checks then trust. Fail closed.
      const needed = (wantsUncompSize ? 8 : 0) + (wantsCompSize ? 8 : 0) + (wantsLfhOffset ? 8 : 0);
      if (size < needed) {
        throw new OpenXmlIoError(
          `openZip: ZIP64 extended-info field for "${entryPath}" declares ${size} bytes` +
            ` but the central directory sentinels require ${needed}`,
        );
      }
      let q = dataStart;
      const result: { uncompSize?: number; compSize?: number; lfhOffset?: number } = {};
      if (wantsUncompSize) {
        result.uncompSize = u64(b, q);
        q += 8;
      }
      if (wantsCompSize) {
        result.compSize = u64(b, q);
        q += 8;
      }
      if (wantsLfhOffset) {
        result.lfhOffset = u64(b, q);
        q += 8;
      }
      return result;
    }
    p = next;
  }
  return {};
}

/** Parse the central directory into an array of entry descriptors. */
function parseCentralDirectory(b: Uint8Array, cdOffset: number, expectedCount: number): CdEntry[] {
  const entries: CdEntry[] = [];
  let p = cdOffset;
  for (let i = 0; i < expectedCount; i++) {
    if (u32(b, p) !== SIG_CD) {
      throw new OpenXmlIoError(`openZip: malformed central directory at byte ${p}`);
    }
    const compMethod = u16(b, p + 10);
    let compSize = u32(b, p + 20);
    let uncompSize = u32(b, p + 24);
    const nameLen = u16(b, p + 28);
    const extraLen = u16(b, p + 30);
    const commentLen = u16(b, p + 32);
    let lfhOffset = u32(b, p + 42);
    const nameBytes = b.subarray(p + 46, p + 46 + nameLen);
    // Bit 11 of the general-purpose flag signals a UTF-8 name, but every name
    // decodes as UTF-8 without consulting it: CP437 and UTF-8 agree on ASCII,
    // and xlsx part names are ASCII.
    const path = CD_NAME_DECODER.decode(nameBytes);

    // ZIP64 Extended Information rewrites whichever of {uncompSize, compSize,
    // lfhOffset} are 0xFFFFFFFF sentinels in the canonical fields. The extra
    // field stores them in the order listed in the spec (and omits any that
    // weren't sentinels), so we have to track which slots to read.
    const wantsUncomp = uncompSize === ZIP32_MAX_U32;
    const wantsComp = compSize === ZIP32_MAX_U32;
    const wantsOffset = lfhOffset === ZIP32_MAX_U32;
    if (wantsUncomp || wantsComp || wantsOffset) {
      const extra = readZip64Extra(b, p + 46 + nameLen, extraLen, wantsUncomp, wantsComp, wantsOffset, path);
      if (extra.uncompSize !== undefined) uncompSize = extra.uncompSize;
      if (extra.compSize !== undefined) compSize = extra.compSize;
      if (extra.lfhOffset !== undefined) lfhOffset = extra.lfhOffset;
    }
    entries.push({ path, lfhOffset, compMethod, compSize, uncompSize });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/**
 * RFC 1951 gives every deflate stream at least one block, and the last block
 * carries the BFINAL bit, so a method-8 entry declaring zero compressed bytes
 * has no stream to decode. Reject it where both read paths meet the bytes,
 * rather than letting each one invent an answer for an entry that isn't there.
 */
const requireDeflatePayload = (path: string, compressed: Uint8Array): void => {
  if (compressed.byteLength === 0) {
    throw new OpenXmlIoError(`openZip: entry "${path}" is declared DEFLATE but carries no compressed bytes`);
  }
};

/** Message for a deflate stream that runs out before signalling its last block. */
const incompleteDeflateError = (path: string): OpenXmlIoError =>
  new OpenXmlIoError(`openZip: deflate stream for "${path}" ended before its final block`);

/** Read the compressed bytes for a CD entry by walking its local file header. */
function readCompressedBytes(b: Uint8Array, entry: CdEntry): Uint8Array {
  if (u32(b, entry.lfhOffset) !== SIG_LFH) {
    throw new OpenXmlIoError(`openZip: malformed local file header for "${entry.path}"`);
  }
  const nameLen = u16(b, entry.lfhOffset + 26);
  const extraLen = u16(b, entry.lfhOffset + 28);
  const dataStart = entry.lfhOffset + 30 + nameLen + extraLen;
  // `subarray` clamps to the end of the buffer instead of throwing, so without
  // this an entry whose declared span runs past EOF would silently inflate (or
  // pass STORE through) from however many bytes happened to be left.
  if (dataStart + entry.compSize > b.length) {
    throw new OpenXmlIoError(
      `openZip: entry "${entry.path}" declares ${entry.compSize} compressed bytes at offset ${dataStart},` +
        ` past the end of the ${b.length}-byte archive`,
    );
  }
  return b.subarray(dataStart, dataStart + entry.compSize);
}

const PDF_MAGIC = [0x25, 0x50, 0x44, 0x46]; // "%PDF"
const UTF8_BOM = [0xef, 0xbb, 0xbf];
const UTF16_LE_BOM = [0xff, 0xfe];
const UTF16_BE_BOM = [0xfe, 0xff];

/**
 * How far to read when deciding that bytes with no recognised magic number are
 * text. Enough for a CSV header row, cheap enough to run on every failure.
 */
const TEXT_SNIFF_BYTES = 64;

const startsWith = (b: Uint8Array, magic: ReadonlyArray<number>): boolean => {
  if (b.length < magic.length) return false;
  for (let i = 0; i < magic.length; i++) {
    if (b[i] !== magic[i]) return false;
  }
  return true;
};

const looksLikeText = (b: Uint8Array): boolean => {
  const end = Math.min(b.length, TEXT_SNIFF_BYTES);
  if (end === 0) return false;
  for (let i = 0; i < end; i++) {
    const c = b[i] ?? 0;
    const printable = c >= 0x20 && c < 0x7f;
    const whitespace = c === 0x09 || c === 0x0a || c === 0x0d;
    if (!printable && !whitespace) return false;
  }
  return true;
};

/**
 * Name what the leading bytes are, for the error a caller sees when something
 * that passed an extension check turns out not to be an xlsx. States only what
 * a magic number actually says and returns `undefined` rather than guessing,
 * so a wrong guess never displaces the real message.
 *
 * The OLE compound-document signature is absent on purpose: `openZip` rejects
 * that container before any of this runs, with a message that names it.
 */
const describeLeadingBytes = (b: Uint8Array): string | undefined => {
  if (b.length >= 4 && u32(b, 0) === SIG_LFH) {
    return 'a zip with no readable central directory, the shape of a truncated or partially uploaded file';
  }
  if (startsWith(b, PDF_MAGIC)) return 'a PDF';
  if (startsWith(b, UTF8_BOM)) return 'text with a UTF-8 byte-order mark, such as a CSV saved under an .xlsx name';
  if (startsWith(b, UTF16_LE_BOM) || startsWith(b, UTF16_BE_BOM)) return 'text with a UTF-16 byte-order mark';
  if (looksLikeText(b)) return 'plain text, such as a CSV saved under an .xlsx name';
  return undefined;
};

const leadingBytesSuffix = (bytes: Uint8Array): string => {
  const looksLike = describeLeadingBytes(bytes);
  return looksLike === undefined ? '' : `; the leading bytes look like ${looksLike}`;
};

const notAZipError = (bytes: Uint8Array, cause: unknown): OpenXmlIoError =>
  new OpenXmlIoError(`openZip: archive is not a valid zip${leadingBytesSuffix(bytes)}`, { cause });

/**
 * Open a buffered xlsx archive in random-access mode. The archive bytes stay
 * resident; entries inflate on demand inside `read(path)`.
 *
 * Falls back to `fflate.unzipSync` when the central directory uses ZIP64
 * sentinel values (entry count == 0xFFFF or any size field == 0xFFFFFFFF) so
 * external ZIP64 archives still load. xlsx files in the wild fit comfortably in
 * ZIP32; the fallback exists for safety.
 *
 * `decompressionLimits` opts the archive into the zip-bomb safeguards
 * documented on {@link DecompressionLimits}; pass `false` to disable. Defaults
 * fit any legitimate xlsx.
 *
 * Every failure here is an {@link OpenXmlIoError} and every one of them is
 * permanent for these bytes: the same input fails the same way, so a caller
 * validating an upload should reject it rather than retry.
 */
export function openRandomAccessArchive(
  bytes: Uint8Array,
  decompressionLimits?: DecompressionLimitsInput,
): ZipArchive {
  // Quick sanity on min archive size.
  if (bytes.length < EOCD_MIN_BYTES) {
    throw new OpenXmlIoError(
      `openZip: archive is ${bytes.length} bytes, shorter than the ${EOCD_MIN_BYTES}-byte minimum` +
        ` for a zip End-of-Central-Directory (EOCD) record${leadingBytesSuffix(bytes)}`,
    );
  }

  let eocdOff: number;
  try {
    eocdOff = findEocd(bytes);
  } catch (cause) {
    throw notAZipError(bytes, cause);
  }

  const resolvedLimits = resolveDecompressionLimits(decompressionLimits);

  // Detect ZIP64 up front so a malformed ZIP64 record fails closed instead of
  // falling back to `unzipSync`. The fallback inflates every entry before any
  // bomb cap runs; the whole point of the ZIP64 path is to keep the streaming
  // caps in play, so we must not silently downgrade when ZIP64 is in use.
  const eocdTotalEntries = u16(bytes, eocdOff + 10);
  const eocdCdSize = u32(bytes, eocdOff + 12);
  const eocdCdOffset = u32(bytes, eocdOff + 16);
  const claimsZip64 =
    eocdTotalEntries === ZIP32_MAX_U16 ||
    eocdCdSize === ZIP32_MAX_U32 ||
    eocdCdOffset === ZIP32_MAX_U32;

  let summary: CdSummary;
  try {
    summary = readCdSummary(bytes, eocdOff);
  } catch (cause) {
    if (claimsZip64) {
      // ZIP64 sentinels are set but the matching record / locator is missing
      // or invalid. Fail closed — `unzipSync` would inflate every entry.
      throw cause instanceof OpenXmlIoError
        ? cause
        : new OpenXmlIoError('openZip: ZIP64 central directory is malformed', { cause });
    }
    // Plain ZIP32 with an unparseable EOCD — fflate's `unzipSync` is more
    // tolerant of quirky archives; the post-hoc bomb check still applies.
    return openViaUnzipSync(bytes, resolvedLimits);
  }

  let entries: CdEntry[];
  try {
    entries = parseCentralDirectory(bytes, summary.cdOffset, summary.totalEntries);
  } catch (cause) {
    if (claimsZip64) {
      throw cause instanceof OpenXmlIoError
        ? cause
        : new OpenXmlIoError('openZip: ZIP64 central directory is malformed', { cause });
    }
    // Malformed ZIP32 CD — fall back to fflate which is more tolerant.
    return openViaUnzipSync(bytes, resolvedLimits);
  }

  const byPath = new Map<string, CdEntry>();
  for (const e of entries) byPath.set(e.path, e);

  const budget: DecompressionBudget | null = resolvedLimits ? createBudget(resolvedLimits) : null;
  if (budget) {
    // Declared CD totals are cheap to inspect — reject obvious bombs before
    // wiring up the inflate state machine.
    checkDeclaredTotals(budget, entries);
  }

  // Re-read cache for small entries, so the `.rels` parts the passthrough walk
  // revisits do not inflate twice.
  const inflateCache = createInflateCache();
  let live = true;
  let archiveBytes: Uint8Array | undefined = bytes;

  const ensureLive = (): Uint8Array => {
    if (!live || !archiveBytes) {
      throw new OpenXmlIoError('openZip: archive is closed');
    }
    return archiveBytes;
  };

  const readEntry = (path: string): Uint8Array => {
    const buf = ensureLive();
    const cached = inflateCache.get(path);
    if (cached) return cached;
    const entry = byPath.get(path);
    if (!entry) {
      throw new OpenXmlIoError(`openZip: no entry at "${path}"`);
    }
    const compressed = readCompressedBytes(buf, entry);
    let out: Uint8Array;
    if (entry.compMethod === COMP_STORE) {
      // STORE has ratio 1, so the only thing left to check is the absolute
      // size cap (and the running archive total).
      if (budget) {
        if (compressed.byteLength > budget.limits.maxEntryUncompressedBytes) {
          throw entryOverflowError(path, budget.limits.maxEntryUncompressedBytes);
        }
        const record = beginEntryInflate(budget, path);
        record(compressed.byteLength);
      }
      // Copy so callers can safely mutate the returned bytes without perturbing
      // the underlying archive view.
      out = compressed.slice();
    } else if (entry.compMethod === COMP_DEFLATE) {
      out = inflateBounded(path, compressed, budget);
    } else {
      throw new OpenXmlIoError(`openZip: unsupported compression method ${entry.compMethod} for "${path}"`);
    }
    inflateCache.set(path, out);
    return out;
  };

  const readEntryStream = (path: string): ReadableStream<Uint8Array> => {
    const buf = ensureLive();
    const cached = inflateCache.get(path);
    if (cached) return singleChunkStream(cached);
    const entry = byPath.get(path);
    if (!entry) {
      throw new OpenXmlIoError(`openZip: no entry at "${path}"`);
    }
    const compressed = readCompressedBytes(buf, entry);
    if (entry.compMethod === COMP_STORE) {
      // STORE means the bytes on disk are the bytes the caller wants; no need
      // to involve the inflate state machine. Run the same caps the sync
      // `read()` STORE branch applies — otherwise an uncompressed bomb entry
      // reached via `readStream()` would bypass the per-entry / archive-total
      // accounting. Ratio is fixed at 1, so only the absolute bounds apply.
      if (budget) {
        if (compressed.byteLength > budget.limits.maxEntryUncompressedBytes) {
          throw entryOverflowError(path, budget.limits.maxEntryUncompressedBytes);
        }
        const record = beginEntryInflate(budget, path);
        record(compressed.byteLength);
      }
      // Copy because callers may mutate the returned bytes.
      return singleChunkStream(compressed.slice());
    }
    if (entry.compMethod !== COMP_DEFLATE) {
      throw new OpenXmlIoError(`openZip: unsupported compression method ${entry.compMethod} for "${path}"`);
    }
    requireDeflatePayload(path, compressed);
    // DEFLATE: drive fflate's `Inflate` from `pull()` so the consumer's
    // demand controls how much we inflate. Each `pull()` either emits one
    // already-buffered inflated chunk or pushes one more block of compressed
    // input — never both — so the ReadableStream internal queue stays at
    // depth 1 and the producer can't race ahead of the consumer.
    const entryCap = budget ? entryInflateCap(budget, compressed.byteLength) : Number.POSITIVE_INFINITY;
    const record = budget ? beginEntryInflate(budget, path) : null;
    let entryEmitted = 0;
    const pending: Uint8Array[] = [];
    let pushedOffset = 0;
    let inflaterFinal = false;
    let inflateError: Error | undefined;
    const inflater = new Inflate((chunk, final) => {
      if (inflateError) return;
      if (chunk.byteLength > 0) {
        entryEmitted += chunk.byteLength;
        if (entryEmitted > entryCap) {
          inflateError = entryOverflowError(path, entryCap);
          return;
        }
        if (record) {
          try {
            record(chunk.byteLength);
          } catch (err) {
            inflateError = err as Error;
            return;
          }
        }
        pending.push(chunk);
      }
      if (final) inflaterFinal = true;
    });
    // Every `pull` must enqueue, close or error: a return that does none of
    // the three leaves the stream machinery to call it again with identical
    // state, which is an uncancellable busy loop. This is the close/error
    // half. An empty `pending` with every compressed byte pushed means there
    // is nothing left to emit, and getting there without a final block from
    // the inflater means the deflate stream ended early.
    const settleWhenDrained = (controller: ReadableStreamDefaultController<Uint8Array>): void => {
      if (pending.length > 0 || pushedOffset < compressed.byteLength) return;
      if (inflaterFinal) {
        controller.close();
        return;
      }
      controller.error(incompleteDeflateError(path));
    };
    return new ReadableStream<Uint8Array>({
      pull(controller) {
        if (inflateError) {
          controller.error(inflateError);
          return;
        }
        // Emit at most one already-buffered inflated chunk per pull; subsequent
        // pulls drain the rest. This caps the stream's internal queue at one
        // chunk regardless of how many ondata callbacks fflate fired off the
        // most recent push.
        const buffered = pending.shift();
        if (buffered) {
          controller.enqueue(buffered);
          settleWhenDrained(controller);
          return;
        }
        // No buffered output: push one block of compressed input and let
        // inflate's ondata fill `pending`. We stop pushing the moment we have
        // something to emit so the next pull can return it without racing
        // further inflation.
        while (pending.length === 0 && pushedOffset < compressed.byteLength) {
          const end = Math.min(pushedOffset + INFLATE_CHUNK_BYTES, compressed.byteLength);
          const slice = compressed.subarray(pushedOffset, end);
          const isLast = end >= compressed.byteLength;
          try {
            inflater.push(slice, isLast);
          } catch (cause) {
            if (!inflateError) {
              inflateError = new OpenXmlIoError(`openZip: failed to inflate "${path}"`, { cause });
            }
            controller.error(inflateError);
            return;
          }
          // The ondata callback may have set inflateError (decompression-bomb
          // guard tripping mid-inflate). Surface it to the consumer before
          // continuing — otherwise we'd loop forever on a CD-lying bomb whose
          // chunks all land beyond the cap and never make it into `pending`.
          if (inflateError) {
            controller.error(inflateError);
            return;
          }
          pushedOffset = end;
        }
        const next = pending.shift();
        if (next) {
          controller.enqueue(next);
        }
        settleWhenDrained(controller);
      },
      cancel() {
        // Consumer abandoned the stream early — drop buffered chunks and
        // advance the cursor past the end so the inflater and the compressed
        // slice are eligible for GC. fflate's `Inflate` has no terminate API
        // but losing the only reference is enough.
        pending.length = 0;
        pushedOffset = compressed.byteLength;
        inflaterFinal = true;
      },
    });
  };

  return {
    list(): string[] {
      ensureLive();
      return [...byPath.keys()].sort();
    },
    has(path: string): boolean {
      if (!live) return false;
      return byPath.has(path);
    },
    read(path: string): Uint8Array {
      return readEntry(path);
    },
    async readAsync(path: string): Promise<Uint8Array> {
      return readEntry(path);
    },
    readStream(path: string): ReadableStream<Uint8Array> {
      return readEntryStream(path);
    },
    close(): void {
      live = false;
      archiveBytes = undefined;
      inflateCache.clear();
      budget?.chargedByPath.clear();
      byPath.clear();
    },
  };
}

/**
 * Inflate `compressed` into a single `Uint8Array`, abort if it crosses the
 * configured per-entry cap or the archive-wide budget. Uses fflate's streaming
 * `Inflate` so the abort can fire on any internal block boundary rather than
 * after fflate has materialised the entire payload.
 */
function inflateBounded(
  path: string,
  compressed: Uint8Array,
  budget: DecompressionBudget | null,
): Uint8Array {
  requireDeflatePayload(path, compressed);
  const cap = budget ? entryInflateCap(budget, compressed.byteLength) : Number.POSITIVE_INFINITY;
  const record = budget ? beginEntryInflate(budget, path) : null;
  const acc: Uint8Array[] = [];
  let emitted = 0;
  let aborted: Error | undefined;
  let sawFinal = false;
  const inflater = new Inflate((chunk, final) => {
    // Record the final block before the early returns below: a zero-byte
    // entry's stream reports `final` on a chunk that carries no bytes.
    if (final) sawFinal = true;
    if (aborted) return;
    if (chunk.byteLength === 0) return;
    emitted += chunk.byteLength;
    if (emitted > cap) {
      aborted = entryOverflowError(path, cap);
      return;
    }
    if (record) {
      try {
        record(chunk.byteLength);
      } catch (err) {
        aborted = err as Error;
        return;
      }
    }
    acc.push(chunk);
  });
  let off = 0;
  while (off < compressed.byteLength) {
    const end = Math.min(off + INFLATE_CHUNK_BYTES, compressed.byteLength);
    const isLast = end >= compressed.byteLength;
    try {
      inflater.push(compressed.subarray(off, end), isLast);
    } catch (cause) {
      if (aborted) throw aborted;
      throw new OpenXmlIoError(`openZip: failed to inflate "${path}"`, { cause });
    }
    if (aborted) throw aborted;
    off = end;
  }
  if (!sawFinal) {
    throw incompleteDeflateError(path);
  }
  const out = new Uint8Array(emitted);
  let cursor = 0;
  for (const chunk of acc) {
    out.set(chunk, cursor);
    cursor += chunk.byteLength;
  }
  return out;
}

/**
 * Fallback for archives we can't parse ourselves. The random-access reader
 * now understands ZIP64 EOCD + the Zip64 Extended Information extra field, so
 * the only way to reach this path is a malformed central directory. fflate's
 * `unzipSync` is more forgiving but inflates every entry up front; for
 * adversarial archives we still rely on a post-hoc cap check (the only one
 * available without our own streaming decoder).
 */
function openViaUnzipSync(
  bytes: Uint8Array,
  limits: ReturnType<typeof resolveDecompressionLimits>,
): ZipArchive {
  let entries: Record<string, Uint8Array> | undefined;
  try {
    entries = unzipSync(bytes);
  } catch (cause) {
    throw notAZipError(bytes, cause);
  }
  // fflate's `unzipSync` returns already-inflated bytes — we can't abort the
  // inflate mid-flight here, but a post-hoc check still rejects a malicious
  // archive *before* any caller-level code touches the bytes. The peak memory
  // spike is bounded by what fflate just produced; the random-access path
  // catches the common cases (ZIP64 + malformed-but-parseable CDs) up front.
  if (limits) {
    const budget = createBudget(limits);
    for (const [path, payload] of Object.entries(entries)) {
      if (payload.byteLength > limits.maxEntryUncompressedBytes) {
        throw new OpenXmlDecompressionBombError(
          `openZip: entry "${path}" inflated to ${payload.byteLength} bytes,` +
            ` exceeding the ${limits.maxEntryUncompressedBytes}-byte per-entry limit` +
            ` (decompression-bomb guard).`,
        );
      }
      const record = beginEntryInflate(budget, path);
      record(payload.byteLength);
    }
  }
  let live = true;
  return {
    list(): string[] {
      if (!live || !entries) throw new OpenXmlIoError('openZip: archive is closed');
      return Object.keys(entries).sort();
    },
    has(path: string): boolean {
      if (!live || !entries) return false;
      return Object.hasOwn(entries, path);
    },
    read(path: string): Uint8Array {
      if (!live || !entries) throw new OpenXmlIoError('openZip: archive is closed');
      const e = entries[path];
      if (!e) throw new OpenXmlIoError(`openZip: no entry at "${path}"`);
      // Copy, as the random-access path does: `read` hands the caller an array
      // of its own, and this one holds every entry for the archive's lifetime.
      return e.slice();
    },
    async readAsync(path: string): Promise<Uint8Array> {
      return this.read(path);
    },
    readStream(path: string): ReadableStream<Uint8Array> {
      // The unzipSync fallback path already has the entry fully inflated
      // (that's the price of dropping back from random-access). Hand it out as
      // a single-chunk stream so callers using the streaming reader don't have
      // to branch on the implementation.
      const inflated = this.read(path);
      return singleChunkStream(inflated);
    },
    close(): void {
      live = false;
      entries = undefined;
    },
  };
}
