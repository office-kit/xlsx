// Node filesystem + Readable / Writable I/O helpers.
//
// Kept separate from `./node.ts` so the buffer-only entry stays free of
// `node:fs` / `node:stream` imports — important for the `@office-kit/xlsx/streaming`
// browser-targeted bundle, which can re-export `fromBuffer` / `toBuffer`
// without dragging Node-only modules into the browser surface. Users who want
// filesystem I/O reach this module directly (or through `@office-kit/xlsx/node` once
// that subpath lands).

import { createReadStream, createWriteStream, readFileSync } from 'node:fs';
import { readFile, unlink } from 'node:fs/promises';
import { once } from 'node:events';
import { Readable, Writable } from 'node:stream';
import { OpenXmlIoError } from '../utils/exceptions';
import type { BufferedSinkWriter, XlsxSink } from './sink';
import type { XlsxSource } from './source';

const EMPTY_BYTES = new Uint8Array(0);

/**
 * Wrap a filesystem path as an XlsxSource. `toBytes` reads the whole file into
 * memory; `toStream` opens a `fs.createReadStream` and bridges it to a Web
 * {@link ReadableStream} via `Readable.toWeb` so the ZIP reader can iterate
 * without loading the entire xlsx up front.
 */
export function fromFile(path: string): XlsxSource {
  if (typeof path !== 'string' || path.length === 0) {
    throw new OpenXmlIoError('fromFile expects a non-empty path string');
  }
  return {
    async toBytes() {
      try {
        return new Uint8Array(await readFile(path));
      } catch (cause) {
        throw new OpenXmlIoError(`fromFile: failed to read "${path}"`, { cause });
      }
    },
    toStream() {
      const nodeStream = createReadStream(path);
      return Readable.toWeb(nodeStream) as unknown as ReadableStream<Uint8Array>;
    },
  };
}

/**
 * Synchronous variant of {@link fromFile}. Convenience for tooling / scripts
 * where the cost of `await fs.readFile` outweighs the ergonomic gain. The
 * returned source's `toBytes` resolves immediately with the bytes already in
 * memory.
 */
export function fromFileSync(path: string): XlsxSource {
  if (typeof path !== 'string' || path.length === 0) {
    throw new OpenXmlIoError('fromFileSync expects a non-empty path string');
  }
  let bytes: Uint8Array;
  try {
    bytes = new Uint8Array(readFileSync(path));
  } catch (cause) {
    throw new OpenXmlIoError(`fromFileSync: failed to read "${path}"`, { cause });
  }
  return {
    async toBytes() {
      return bytes;
    },
    toStream() {
      return new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(bytes);
          controller.close();
        },
      });
    },
  };
}

/**
 * Filesystem sink. Each `write(chunk)` call streams the bytes to disk via
 * `fs.createWriteStream`, honouring backpressure: the actual `writable.write`
 * for each chunk is queued behind any pending `drain`, so the writable's
 * internal buffer never grows past its `highWaterMark` (default 16 KB) no
 * matter how fast the producer hands chunks over.
 *
 * Note on the producer-side memory budget: the sink contract is
 * intentionally synchronous (`write(chunk): void`), so a producer that races
 * ahead without yielding will let chunk references pile up in the queue.
 * That keeps `writable`'s buffer bounded but does not bound the queue
 * itself. Producers that need a hard ceiling should yield between writes
 * (`await new Promise(setImmediate)` is enough) or use a sink with an async
 * write contract.
 *
 * `result()` returns the destination path; `finish()` resolves with an empty
 * `Uint8Array` once the stream has flushed. Callers that need the on-disk
 * bytes should `readFile()` the returned path themselves — re-reading inside
 * `finish()` would defeat the "streamed to disk, never resident" guarantee.
 */
export function toFile(path: string): XlsxSink & { toBytes(): BufferedSinkWriter; result(): string } {
  if (typeof path !== 'string' || path.length === 0) {
    throw new OpenXmlIoError('toFile expects a non-empty path string');
  }
  let stream: ReturnType<typeof createWriteStream> | undefined;
  let streamCreated = false;
  let finalised: Promise<Uint8Array> | undefined;
  let pendingError: Error | undefined;
  // Backpressure queue: every chunk's actual `writable.write` call is staged
  // behind the previous chunk's completion. When a write returns `false` the
  // queue parks on `drain` before the next chunk goes out, so the writable's
  // internal buffer stays within its highWaterMark.
  let writeQueue: Promise<void> = Promise.resolve();

  const ensureStream = (): NonNullable<typeof stream> => {
    if (!stream) {
      stream = createWriteStream(path);
      streamCreated = true;
      stream.on('error', (err) => {
        pendingError = err instanceof Error ? err : new Error(String(err));
      });
    }
    return stream;
  };

  // Best-effort: remove a half-written file when finish() fails so callers
  // don't mistake a corrupt artefact for a successful save. Swallows unlink
  // errors because we're already in a failure path and the original cause
  // is what the caller needs to see.
  const cleanupOnFailure = async (): Promise<void> => {
    if (!streamCreated) return;
    try {
      await unlink(path);
    } catch {
      // ignore — file may be gone, path may be on a read-only fs, etc.
    }
  };

  return {
    toBytes(): BufferedSinkWriter {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) throw new OpenXmlIoError(`toFile sink: write after finish ("${path}")`);
          if (!(chunk instanceof Uint8Array)) {
            throw new OpenXmlIoError(`toFile sink: chunk is not a Uint8Array ("${path}")`);
          }
          if (pendingError) throw new OpenXmlIoError(`toFile sink: write error on "${path}"`, { cause: pendingError });
          const s = ensureStream();
          writeQueue = writeQueue.then(async () => {
            // Skip remaining work once the stream has errored — the error
            // surfaces from `finish()` so callers see one consistent failure.
            if (pendingError) return;
            const ok = s.write(chunk);
            if (!ok) {
              // Writable's internal buffer is over highWaterMark; wait for it
              // to flush before the next queued chunk runs.
              await once(s, 'drain');
            }
          });
        },
        async finish(): Promise<Uint8Array> {
          if (finalised) return finalised;
          finalised = (async () => {
            const s = ensureStream();
            try {
              await writeQueue;
              await new Promise<void>((resolve, reject) => {
                s.end((err?: Error | null) => (err ? reject(err) : resolve()));
              });
              if (pendingError) {
                throw new OpenXmlIoError(`toFile sink: write error on "${path}"`, { cause: pendingError });
              }
            } catch (err) {
              await cleanupOnFailure();
              throw err;
            }
            return EMPTY_BYTES;
          })();
          return finalised;
        },
        abort(): void {
          // Idempotent: subsequent finish() / abort() calls become no-ops.
          if (finalised) return;
          finalised = Promise.resolve(EMPTY_BYTES);
          if (stream) {
            // destroy() releases the fd synchronously without flushing the
            // pending buffer — exactly what we want for an aborted save.
            stream.destroy();
          }
          // Fire-and-forget unlink — abort() is sync (void), and the caller
          // is already on a failure path so any unlink error is noise.
          void cleanupOnFailure();
        },
      };
    },
    result(): string {
      return path;
    },
  };
}

/**
 * Wrap a Node.js {@link Readable} as an XlsxSource. `toBytes` consumes the
 * entire stream synchronously (collecting chunks); `toStream` bridges to a Web
 * ReadableStream via `Readable.toWeb` so the ZIP reader can pull chunks lazily.
 */
export function fromReadable(readable: Readable): XlsxSource {
  if (!(readable instanceof Readable)) {
    throw new OpenXmlIoError('fromReadable expects a Node Readable');
  }
  let bytes: Promise<Uint8Array> | undefined;
  return {
    async toBytes() {
      if (bytes) return bytes;
      bytes = (async () => {
        const chunks: Uint8Array[] = [];
        for await (const c of readable) {
          chunks.push(c instanceof Uint8Array ? c : new Uint8Array(c));
        }
        let total = 0;
        for (const c of chunks) total += c.byteLength;
        const out = new Uint8Array(total);
        let off = 0;
        for (const c of chunks) {
          out.set(c, off);
          off += c.byteLength;
        }
        return out;
      })();
      return bytes;
    },
    toStream() {
      return Readable.toWeb(readable) as unknown as ReadableStream<Uint8Array>;
    },
  };
}

/**
 * Wrap a Node.js {@link Writable} as an XlsxSink. The actual
 * `writable.write` for each chunk is queued behind any pending `drain`, so
 * the writable's internal buffer never exceeds its `highWaterMark` regardless
 * of how fast the producer is. See {@link toFile} for the same caveat about
 * producer-side memory: the synchronous `write(chunk)` API does not let
 * backpressure flow back to the caller, so a tight non-yielding producer can
 * still let chunk references accumulate in the queue.
 *
 * `result()` returns the writable itself for downstream chaining.
 */
export function toWritable(writable: Writable): XlsxSink & { toBytes(): BufferedSinkWriter; result(): Writable } {
  if (!(writable instanceof Writable)) {
    throw new OpenXmlIoError('toWritable expects a Node Writable');
  }
  let finalised: Promise<Uint8Array> | undefined;
  let pendingError: Error | undefined;
  let writeQueue: Promise<void> = Promise.resolve();
  writable.on('error', (err) => {
    pendingError = err instanceof Error ? err : new Error(String(err));
  });

  return {
    toBytes(): BufferedSinkWriter {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) throw new OpenXmlIoError('toWritable sink: write after finish');
          if (!(chunk instanceof Uint8Array)) throw new OpenXmlIoError('toWritable sink: chunk is not a Uint8Array');
          if (pendingError) throw new OpenXmlIoError('toWritable sink: write error', { cause: pendingError });
          writeQueue = writeQueue.then(async () => {
            if (pendingError) return;
            const ok = writable.write(chunk);
            if (!ok) {
              await once(writable, 'drain');
            }
          });
        },
        async finish(): Promise<Uint8Array> {
          if (finalised) return finalised;
          finalised = (async () => {
            await writeQueue;
            await new Promise<void>((resolve, reject) => {
              writable.end((err?: Error | null) => (err ? reject(err) : resolve()));
            });
            if (pendingError) throw new OpenXmlIoError('toWritable sink: write error', { cause: pendingError });
            return EMPTY_BYTES;
          })();
          return finalised;
        },
        abort(cause?: unknown): void {
          if (finalised) return;
          finalised = Promise.resolve(EMPTY_BYTES);
          // Pass the cause to destroy() so downstream `error` listeners can
          // distinguish a deliberate abort from spontaneous fs/network errors.
          writable.destroy(cause instanceof Error ? cause : undefined);
        },
      };
    },
    result(): Writable {
      return writable;
    },
  };
}
