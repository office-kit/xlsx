// ZIP read layer.
//
// `openZip(source)` walks the central directory once and inflates each entry on
// demand inside `read(path)` (see `./random-access-reader.ts`). That keeps peak
// memory at compressed-archive size + per-entry inflate scratch + the bounded
// cache of small re-read entries (`./inflate-cache.ts`), instead of holding
// every uncompressed entry resident at once the way the old `unzipSync`
// shortcut did. The fallback path through fflate's `unzipSync` is preserved for
// ZIP64 / non-standard archives.

import type { XlsxSource } from '../io/source.js';
import { OpenXmlIoError, OpenXmlNotImplementedError } from '../utils/exceptions.js';
import type { DecompressionLimits } from './decompression-guard.js';
import { openRandomAccessArchive } from './random-access-reader.js';

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const isCfbCompoundDocument = (bytes: Uint8Array): boolean => {
  if (bytes.length < CFB_MAGIC.length) return false;
  for (let i = 0; i < CFB_MAGIC.length; i++) {
    if (bytes[i] !== CFB_MAGIC[i]) return false;
  }
  return true;
};

export interface ZipArchive {
  /** Sorted list of all entry paths in the archive. */
  list(): string[];
  /**
   * Synchronous read; throws OpenXmlIoError when the path is unknown. Each
   * call returns an array the caller owns: mutating it changes neither the
   * archive nor what a later read of the same path returns.
   */
  read(path: string): Uint8Array;
  /** Promise variant for symmetry with the future streaming reader. */
  readAsync(path: string): Promise<Uint8Array>;
  /**
   * Streaming read: returns the entry's inflated bytes as a Web
   * `ReadableStream<Uint8Array>` chunk-by-chunk. Lets callers (the streaming
   * worksheet iterator, in particular) push the inflated payload through a SAX
   * parser without first materialising it in full — peak memory for a sheet
   * walk drops to the inflate window + SAX state instead of the entire
   * uncompressed worksheet body. Throws OpenXmlIoError when the path is
   * unknown.
   */
  readStream(path: string): ReadableStream<Uint8Array>;
  /** Whether the archive holds an entry at the given path. */
  has(path: string): boolean;
  /** Release the in-memory entry table. Subsequent reads throw. */
  close(): void;
}

/** Options for {@link openZip}. */
export interface OpenZipOptions {
  /**
   * Decompression-bomb safeguards applied while inflating archive entries. The
   * default limits admit any legitimate xlsx and reject pathological archives
   * (extreme compression ratios, gigabyte-scale entries). Pass `false` to
   * disable the guard entirely — only safe when the source is fully trusted.
   * See {@link DecompressionLimits} for the individual knobs.
   */
  decompressionLimits?: DecompressionLimits | false;
}

/**
 * Open a zip archive from any {@link XlsxSource}. The source is fully
 * materialised in memory, the central directory is parsed once, and each
 * entry is inflated on demand by {@link openRandomAccessArchive}: peak memory
 * stays at compressed-archive size, plus per-entry inflate scratch, plus a few
 * MB at most of small entries kept for re-reads, rather than holding every
 * uncompressed entry resident. The fflate `unzipSync` fallback is preserved
 * internally for ZIP64 / non-standard archives the random-access reader
 * rejects, and it does hold every entry inflated.
 *
 * Throws {@link OpenXmlIoError} when the bytes are not a readable zip, and
 * when the source itself fails to produce them. The two are the only reason a
 * caller has to look past the class: see the error-handling contract on
 * {@link OpenXmlIoError}.
 */
export async function openZip(source: XlsxSource, opts: OpenZipOptions = {}): Promise<ZipArchive> {
  let bytes: Uint8Array;
  try {
    bytes = await source.toBytes();
  } catch (cause) {
    throw new OpenXmlIoError('openZip: failed to read source bytes', { cause });
  }

  // `D0 CF 11 E0 A1 B1 1A E1` is an OLE Compound File Binary container, not a
  // zip. Both an encrypted xlsx (Excel 2007+ password protection) and a legacy
  // .xls workbook arrive in one, and telling them apart means parsing the CFB
  // directory, so the message names both instead of asserting one. Detected
  // here so neither fails later with a generic invalid-zip message.
  if (isCfbCompoundDocument(bytes)) {
    throw new OpenXmlNotImplementedError(
      'openZip: the input is an OLE compound-document container, not a zip. Encrypted xlsx is' +
        ' not supported (decrypt with msoffcrypto-tool first); a legacy .xls workbook is not' +
        ' either (re-save it as .xlsx, or read it with SheetJS).',
    );
  }

  return openRandomAccessArchive(bytes, opts.decompressionLimits);
}
