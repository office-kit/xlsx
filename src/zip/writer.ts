// ZIP write layer. Per-entry compression via fflate keeps the whole archive
// out of memory. Buffered entries use deflateSync; streaming entries use
// ZipDeflate. Each addEntry compresses its supplied bytes and the
// resulting ZIP chunks land on the sink one at a time — the buffered
// `toBytes()` sink concatenates them on finish, while a streaming sink can
// flush them as they arrive.
//
// ZIP64 (entry count > 65535): fflate's `Zip` emits a plain ZIP32 EOCD in all
// cases, so on finalize we splice in a ZIP64 EOCD record + locator when needed
// via `applyZip64EntryCountPatch`. That keeps the per-entry LFH/CDH layout
// fflate produces and only rewrites the trailing records.
//
// Scope: this covers the entry-count-overflow case (the limit xlsx archives
// realistically hit — `tens of millions of cells` → tens of thousands of
// worksheet entries via the streaming writer). Per-entry compressed/uncompressed
// sizes and the central-directory offset must still fit in 32 bits (≤ 4 GiB
// each); a single >4 GiB entry would need full ZIP64 size support and
// `applyZip64EntryCountPatch` throws `OpenXmlNotImplementedError` if we ever
// detect that. xlsx workbooks don't approach that limit in practice, but the
// constraint is real — surface it in your own size estimates.

import { deflateSync, Zip, ZipDeflate, ZipPassThrough, type DeflateOptions } from 'fflate';
import type { XlsxSink } from '../io/sink.js';
import { OpenXmlIoError } from '../utils/exceptions.js';
import { applyZip64EntryCountPatch } from './zip64-patch.js';

// fflate 0.8.3's streaming compressor can reference its zero-filled lookback
// before the first input byte for some binary payloads. A complete entry uses
// deflateSync, which has no synthetic lookback and preserves these bytes.
class BufferedZipDeflate extends ZipPassThrough {
  constructor(path: string, private readonly options: DeflateOptions | undefined) {
    super(path);
    this.compression = 8;
  }

  override process(chunk: Uint8Array, final: boolean): void {
    this.ondata(null, deflateSync(chunk, this.options), final);
  }
}

const ZIP32_MAX_ENTRIES = 0xffff;
const LOCAL_TIMESTAMP_OFFSET = 10;
const CENTRAL_TIMESTAMP_OFFSET = 12;
const CENTRAL_HEADER_SIZE = 46;
const CENTRAL_NAME_LENGTH_OFFSET = 28;
const CENTRAL_EXTRA_LENGTH_OFFSET = 30;
const CENTRAL_COMMENT_LENGTH_OFFSET = 32;

/** Deflate effort: 0 skips compression, 9 is the slowest and smallest. */
export type CompressionLevel = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

/** Supported year range of the fflate ZIP backend. */
const MIN_ZIP_YEAR = 1980;
const MAX_ZIP_YEAR = 2099;

export interface ZipWriterOptions {
  /**
   * Last-modified timestamp stamped into every entry's local header and
   * central-directory record. ZIP has no "no timestamp" encoding, so fflate
   * defaults each entry to the wall clock and two archives built from
   * identical input differ in bytes. Pin this to get reproducible output for
   * golden-file tests or content-addressed caching.
   *
   * Recorded as the date's UTC wall time, to a two-second resolution, with the
   * year required to fall in 1980-2099. The DOS field carries no timezone, so
   * writing local components would leave the bytes depending on the writer's
   * `TZ`, which is the opposite of what pinning a stamp is for.
   */
  mtime?: Date;
  /** Deflate level handed to fflate. Defaults to fflate's own 6. */
  compressionLevel?: CompressionLevel;
}

export interface ZipWriter {
  /**
   * Stage an entry. Bytes are pushed through fflate's `ZipDeflate` /
   * `ZipPassThrough` stream synchronously, so the deflated chunks land on the
   * sink as the call runs (no per-entry buffering — see the streaming-behaviour
   * test in `tests/phase-1/zip/writer.test.ts`). Streams (`ReadableStream`)
   * are not accepted today; pass an already-materialised entry, or use
   * {@link addStreamingEntry} for chunked writes.
   *
   * `compress` defaults to `true`. Pass `false` for already-compressed payloads
   * (PNG/JPEG/zip-as-binary content like vbaProject.bin) so we don't pay
   * deflate costs for no gain.
   */
  addEntry(path: string, bytes: Uint8Array | ReadableStream<Uint8Array>, opts?: { compress?: boolean }): Promise<void>;

  /**
   * Open a streaming entry. Returns a writer the caller can `write()` chunks to
   * and `end()` to seal the entry. Each chunk pushes through the same fflate
   * `ZipDeflate` / `ZipPassThrough` machinery as `addEntry`, so peak memory
   * stays at one chunk + deflate scratch even for multi-GB worksheets.
   *
   * Sequencing: only one streaming entry may be open at a time — `addEntry` and
   * a second `addStreamingEntry` both throw until the current entry's `end()`
   * resolves.
   */
  addStreamingEntry(path: string, opts?: { compress?: boolean }): StreamingEntryWriter;

  /**
   * Build the central directory and flush all bytes through the sink.
   * Idempotent; subsequent calls resolve to the same payload.
   */
  finalize(): Promise<Uint8Array>;

  /**
   * Release the sink and underlying writer without producing a valid archive.
   * Use this from a surrounding catch block when serialization fails part-way
   * through. Without it, streaming sinks (`toFile` / `toWritable`) keep their
   * file descriptors / writables open and the half-written xlsx looks valid on
   * disk. Idempotent; safe to call after `finalize()`.
   *
   * Await the result before reporting the failure: `toFile` removes its partial
   * file asynchronously, so an unawaited abort can still have the file on disk
   * when the caller sees the error.
   */
  abort(cause?: unknown): void | Promise<void>;
}

/** Writer handle for a single streaming entry. */
export interface StreamingEntryWriter {
  /** Push a chunk of bytes (already-encoded). Throws after `end()`. */
  write(chunk: Uint8Array): void;
  /** Seal the entry. Subsequent `write()` throws. Idempotent. */
  end(): Promise<void>;
}

/** Encode UTC components directly: local dates cannot represent a DST gap. */
const toZipStamp = (mtime: Date): number => {
  const ms = mtime.getTime();
  if (Number.isNaN(ms)) {
    throw new OpenXmlIoError('createZipWriter: mtime is an invalid Date');
  }
  const year = mtime.getUTCFullYear();
  if (year < MIN_ZIP_YEAR || year > MAX_ZIP_YEAR) {
    throw new OpenXmlIoError(
      `createZipWriter: mtime ${mtime.toISOString()} is outside the supported ZIP timestamp range (${MIN_ZIP_YEAR}-${MAX_ZIP_YEAR})`,
    );
  }
  return ((year - MIN_ZIP_YEAR) << 25)
    | ((mtime.getUTCMonth() + 1) << 21)
    | (mtime.getUTCDate() << 16)
    | (mtime.getUTCHours() << 11)
    | (mtime.getUTCMinutes() << 5)
    | (mtime.getUTCSeconds() >> 1);
};

/**
 * fflate looks the level up in a table and falls back to 6 for anything off
 * the end, so an out-of-range level would quietly produce default output. The
 * union type catches that for TypeScript callers; this catches it for the rest.
 */
const validateCompressionLevel = (level: number): void => {
  if (!Number.isInteger(level) || level < 0 || level > 9) {
    throw new OpenXmlIoError(`createZipWriter: compressionLevel must be an integer in [0, 9]; got ${level}`);
  }
};

/**
 * ZIP writer backed by fflate's streaming `Zip` class. Entries are pushed
 * through `ZipDeflate` / `ZipPassThrough` streams as they arrive, so peak
 * memory stays at the size of the in-flight entry plus the output buffer rather
 * than the full archive.
 *
 * The sink contract is `toBytes()`, but that name is historical: the sink is
 * driven by a chunked `write(chunk)` API that fans bytes out as they arrive.
 * The buffered Node/browser sinks (`toBuffer`, `toBlob`, `toArrayBuffer`)
 * concatenate the chunks for a single-shot result; streaming sinks
 * (`toFile`, `toWritable`) forward each chunk to disk / the wrapped writable
 * without ever holding the full archive resident. Either kind plugs in here.
 */
export function createZipWriter(sink: XlsxSink, opts: ZipWriterOptions = {}): ZipWriter {
  // Both options are checked before the sink is opened: a bad one otherwise
  // surfaces from fflate half-way through the first entry, by which time a file
  // sink holds a partial archive.
  if (opts.compressionLevel !== undefined) validateCompressionLevel(opts.compressionLevel);
  const stamp = opts.mtime === undefined ? undefined : toZipStamp(opts.mtime);
  const writer = sink.toBytes();
  const deflateOpts = opts.compressionLevel === undefined ? undefined : { level: opts.compressionLevel };
  let pendingLocalHeader = false;
  // fflate clones mtime and reads local getters, so even a Date subclass cannot
  // represent UTC times in a local DST gap. Give it a safe placeholder and patch
  // only its header chunks below; payload chunks must never be signature-scanned.
  const newEntry = (path: string, compress: boolean, buffered = false): ZipPassThrough | ZipDeflate => {
    const file = !compress ? new ZipPassThrough(path)
      : buffered ? new BufferedZipDeflate(path, deflateOpts) : new ZipDeflate(path, deflateOpts);
    if (stamp !== undefined) {
      file.mtime = new Date(2000, 0, 1);
      pendingLocalHeader = true;
    }
    return file;
  };
  let finalised: Promise<Uint8Array> | undefined;
  let endCalled = false;
  const seen = new Set<string>();
  const errors: Error[] = [];
  // fflate emits the [CD | EOCD] block in a single ondata call with
  // `final=true`. We capture only that chunk so we can apply the ZIP64 patch on
  // finalize; all preceding entry-data chunks stream straight to the sink to
  // preserve the writer's incremental flushing contract.
  let finalChunk: Uint8Array | undefined;
  let zipFinishResolve: (() => void) | undefined;
  const zipFinishPromise = new Promise<void>((resolve) => {
    zipFinishResolve = resolve;
  });

  const zip = new Zip((err, chunk, final) => {
    if (err) {
      errors.push(err instanceof Error ? err : new Error(String(err)));
      return;
    }
    // ZipDeflate emits an empty trailer chunk on the final callback even when
    // there are no bytes; guard against pushing an undefined chunk.
    if (chunk && chunk.byteLength > 0) {
      if (stamp !== undefined && (final || pendingLocalHeader)) {
        const view = new DataView(chunk.buffer, chunk.byteOffset, chunk.byteLength);
        if (final) {
          // fflate emits the complete central directory plus EOCD as one chunk.
          let offset = 0;
          for (let i = 0; i < seen.size; i++) {
            view.setUint32(offset + CENTRAL_TIMESTAMP_OFFSET, stamp, true);
            offset += CENTRAL_HEADER_SIZE + view.getUint16(offset + CENTRAL_NAME_LENGTH_OFFSET, true)
              + view.getUint16(offset + CENTRAL_EXTRA_LENGTH_OFFSET, true)
              + view.getUint16(offset + CENTRAL_COMMENT_LENGTH_OFFSET, true);
          }
        } else if (pendingLocalHeader) {
          // With one entry open at a time, its first chunk is the local header.
          view.setUint32(LOCAL_TIMESTAMP_OFFSET, stamp, true);
          pendingLocalHeader = false;
        }
      }
      if (final) {
        // Buffer the trailing CD + EOCD block; written after possible patch.
        finalChunk = chunk;
      } else {
        writer.write(chunk);
      }
    }
    if (final && zipFinishResolve) {
      zipFinishResolve();
      zipFinishResolve = undefined;
    }
  });

  let streamingOpen = false;

  const guardAdd = (path: string): void => {
    if (finalised !== undefined) {
      throw new OpenXmlIoError('createZipWriter: addEntry after finalize');
    }
    if (streamingOpen) {
      throw new OpenXmlIoError('createZipWriter: a streaming entry is still open — call end() first');
    }
    if (seen.has(path)) {
      throw new OpenXmlIoError(`createZipWriter: duplicate entry "${path}"`);
    }
  };

  return {
    async addEntry(path, bytes, entryOpts) {
      if (!(bytes instanceof Uint8Array)) {
        throw new OpenXmlIoError(
          'createZipWriter: ReadableStream entries are not yet supported (deferred to streaming writer)',
        );
      }
      guardAdd(path);
      seen.add(path);
      const file = newEntry(path, entryOpts?.compress ?? true, true);
      try {
        zip.add(file);
        file.push(bytes, /* final */ true);
      } catch (cause) {
        throw new OpenXmlIoError(`createZipWriter: failed to add entry "${path}"`, { cause });
      }
      if (errors.length > 0) {
        throw new OpenXmlIoError('createZipWriter: stream error during addEntry', { cause: errors[0] });
      }
    },

    addStreamingEntry(path, entryOpts) {
      guardAdd(path);
      seen.add(path);
      streamingOpen = true;
      const file = newEntry(path, entryOpts?.compress ?? true);
      try {
        zip.add(file);
      } catch (cause) {
        streamingOpen = false;
        throw new OpenXmlIoError(`createZipWriter: failed to open streaming entry "${path}"`, { cause });
      }
      let ended = false;
      return {
        write(chunk: Uint8Array): void {
          if (ended) throw new OpenXmlIoError(`createZipWriter: write after end on "${path}"`);
          if (!(chunk instanceof Uint8Array)) {
            throw new OpenXmlIoError(`createZipWriter: streaming entry "${path}" chunk is not a Uint8Array`);
          }
          if (chunk.byteLength === 0) return;
          try {
            file.push(chunk, /* final */ false);
          } catch (cause) {
            throw new OpenXmlIoError(`createZipWriter: failed to push chunk on "${path}"`, { cause });
          }
          if (errors.length > 0) {
            throw new OpenXmlIoError('createZipWriter: stream error during write', { cause: errors[0] });
          }
        },
        async end(): Promise<void> {
          if (ended) return;
          ended = true;
          try {
            file.push(new Uint8Array(0), /* final */ true);
          } catch (cause) {
            throw new OpenXmlIoError(`createZipWriter: failed to end streaming entry "${path}"`, { cause });
          }
          streamingOpen = false;
          if (errors.length > 0) {
            throw new OpenXmlIoError('createZipWriter: stream error during end', { cause: errors[0] });
          }
        },
      };
    },

    async finalize() {
      if (finalised !== undefined) return finalised;
      if (streamingOpen) {
        throw new OpenXmlIoError('createZipWriter: cannot finalize while a streaming entry is open');
      }
      finalised = (async () => {
        try {
          if (!endCalled) {
            zip.end();
            endCalled = true;
          }
        } catch (cause) {
          throw new OpenXmlIoError('createZipWriter: failed to finalize zip archive', { cause });
        }
        await zipFinishPromise;
        if (errors.length > 0) {
          throw new OpenXmlIoError('createZipWriter: stream error during finalize', { cause: errors[0] });
        }

        // Apply the ZIP64 patch to fflate's [CD | EOCD] tail when the entry
        // count exceeds ZIP32's 16-bit cap, then flush the (possibly patched)
        // tail to the sink.
        if (finalChunk) {
          const patched =
            seen.size > ZIP32_MAX_ENTRIES
              ? applyZip64EntryCountPatch(finalChunk, seen.size)
              : finalChunk;
          writer.write(patched);
        }
        return writer.finish();
      })();
      return finalised;
    },

    abort(cause?: unknown): void | Promise<void> {
      if (finalised !== undefined) return;
      // Mark finalised so any subsequent addEntry / finalize short-circuits.
      finalised = Promise.resolve(new Uint8Array(0));
      // Drop fflate's listener: further `ondata` callbacks no longer matter.
      if (zipFinishResolve) {
        zipFinishResolve();
        zipFinishResolve = undefined;
      }
      return writer.abort?.(cause);
    },
  };
}
