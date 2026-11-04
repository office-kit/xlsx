// Browser-target I/O helpers. Phase-1 §1: synchronous in-memory paths.
// Response / streaming-WritableStream bridges land alongside the
// streaming ZIP work in §2.
//
// Node 18+ provides Blob / File / FormData / fetch / Web Streams natively,
// so this module is also exercised by the Node-hosted vitest runner.

import { OpenXmlIoError } from '../utils/exceptions.js';
import type { BufferedSinkWriter, XlsxSink } from './sink.js';
import type { XlsxSource } from './source.js';
import { isArrayBuffer, isUint8Array } from '../utils/bytes.js';

/**
 * Wrap a Blob (or File, since File extends Blob) as an XlsxSource. The
 * source materialises bytes lazily on the first {@link XlsxSource.toBytes}
 * call.
 */
export function fromBlob(blob: Blob): XlsxSource {
  if (!(blob instanceof Blob)) {
    throw new OpenXmlIoError('fromBlob expects a Blob (or File)');
  }
  return {
    async toBytes() {
      try {
        const ab = await blob.arrayBuffer();
        return new Uint8Array(ab);
      } catch (cause) {
        throw new OpenXmlIoError('fromBlob: failed to read blob contents', { cause });
      }
    },
    toStream() {
      return blob.stream();
    },
  };
}

/**
 * Wrap a fetch {@link Response} as an XlsxSource. `toBytes` collects
 * the entire response body via `Response.arrayBuffer()`; `toStream`
 * returns `response.body` directly so the ZIP reader can pull chunks
 * lazily from the network. The Response can have been fetched with
 * any method and content-type; this helper does no validation.
 */
export function fromResponse(response: Response): XlsxSource {
  if (!(response instanceof Response)) {
    throw new OpenXmlIoError('fromResponse expects a fetch Response');
  }
  let bytes: Promise<Uint8Array> | undefined;
  return {
    async toBytes() {
      if (bytes) return bytes;
      bytes = (async () => {
        try {
          return new Uint8Array(await response.arrayBuffer());
        } catch (cause) {
          throw new OpenXmlIoError('fromResponse: failed to read response body', { cause });
        }
      })();
      return bytes;
    },
    toStream() {
      const body = response.body;
      if (!body) {
        // Bodyless responses (HEAD / 204 / 304) — return an empty stream.
        return new ReadableStream<Uint8Array>({
          start(controller) {
            controller.close();
          },
        });
      }
      return body;
    },
  };
}

/**
 * Wrap a Web {@link ReadableStream} of bytes as an XlsxSource. `toStream`
 * returns the stream directly; `toBytes` drains it once and caches the
 * bytes. Streams can only be consumed once — calling `toBytes` after
 * `toStream` (or vice versa) on the same source throws.
 */
export function fromStream(stream: ReadableStream<Uint8Array>): XlsxSource {
  if (!(stream instanceof ReadableStream)) {
    throw new OpenXmlIoError('fromStream expects a ReadableStream<Uint8Array>');
  }
  let consumed = false;
  let bytes: Promise<Uint8Array> | undefined;
  return {
    async toBytes() {
      if (bytes) return bytes;
      if (consumed) {
        throw new OpenXmlIoError('fromStream: stream already consumed via toStream()');
      }
      consumed = true;
      bytes = (async () => {
        const chunks: Uint8Array[] = [];
        const reader = stream.getReader();
        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            if (value) chunks.push(value);
          }
        } finally {
          reader.releaseLock();
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
      if (bytes) {
        throw new OpenXmlIoError('fromStream: bytes already consumed via toBytes()');
      }
      if (consumed) {
        throw new OpenXmlIoError('fromStream: stream already consumed');
      }
      consumed = true;
      return stream;
    },
  };
}

/**
 * Wrap an ArrayBuffer or Uint8Array. The bytes are referenced (Uint8Array
 * input) or wrapped without copy (ArrayBuffer input).
 */
export function fromArrayBuffer(buf: ArrayBuffer | Uint8Array): XlsxSource {
  let bytes: Uint8Array;
  if (isUint8Array(buf)) {
    bytes = buf;
  } else if (isArrayBuffer(buf)) {
    bytes = new Uint8Array(buf);
  } else {
    throw new OpenXmlIoError('fromArrayBuffer expects an ArrayBuffer or Uint8Array');
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

const collectChunks = (chunks: Uint8Array[]): Uint8Array => {
  let total = 0;
  for (const c of chunks) total += c.byteLength;
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.byteLength;
  }
  return out;
};

/**
 * In-memory Blob sink. Convenience `result()` returns the accumulated
 * payload as a Blob with the supplied MIME type.
 */
export function toBlob(
  mime: string = 'application/octet-stream',
): XlsxSink & { toBytes(): BufferedSinkWriter; result(): Blob } {
  const chunks: Uint8Array[] = [];
  let finalised: Uint8Array | undefined;

  const finalise = (): Uint8Array => {
    if (finalised !== undefined) return finalised;
    finalised = collectChunks(chunks);
    chunks.length = 0;
    return finalised;
  };

  return {
    toBytes(): BufferedSinkWriter {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) throw new OpenXmlIoError('toBlob sink: write after finish');
          if (!isUint8Array(chunk)) throw new OpenXmlIoError('toBlob sink: chunk is not a Uint8Array');
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
    result(): Blob {
      // Copy the bytes' typed-array view into a fresh Blob to detach from
      // the writer's internal buffer.
      const bytes = finalise();
      return new Blob([bytes.slice()], { type: mime });
    },
  };
}

/** In-memory ArrayBuffer sink. */
export function toArrayBuffer(): XlsxSink & { toBytes(): BufferedSinkWriter; result(): ArrayBuffer } {
  const chunks: Uint8Array[] = [];
  let finalised: Uint8Array | undefined;

  const finalise = (): Uint8Array => {
    if (finalised !== undefined) return finalised;
    finalised = collectChunks(chunks);
    chunks.length = 0;
    return finalised;
  };

  return {
    toBytes(): BufferedSinkWriter {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) throw new OpenXmlIoError('toArrayBuffer sink: write after finish');
          if (!isUint8Array(chunk)) throw new OpenXmlIoError('toArrayBuffer sink: chunk is not a Uint8Array');
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
    result(): ArrayBuffer {
      const bytes = finalise();
      // bytes' underlying buffer may be larger than its byteLength (slice
      // semantics on Uint8Array). Materialise an exactly-sized copy.
      const out = new ArrayBuffer(bytes.byteLength);
      new Uint8Array(out).set(bytes);
      return out;
    },
  };
}
