// ZIP layer public surface. Reader and writer are memory-mode for now;
// streaming variants live alongside.

export type {
  DecompressionLimits,
  ResolvedDecompressionLimits,
} from './decompression-guard.js';
export { DEFAULT_DECOMPRESSION_LIMITS } from './decompression-guard.js';
export type { OpenZipOptions, ZipArchive } from './reader.js';
export { openZip } from './reader.js';
export type { CompressionLevel, StreamingEntryWriter, ZipWriter, ZipWriterOptions } from './writer.js';
export { createZipWriter } from './writer.js';
