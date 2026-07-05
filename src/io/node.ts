// In-memory Node helpers.
//
// `fromBuffer` / `toBuffer` rely only on the global `Buffer` symbol — no
// `node:*` imports — so they're safe to ship through the `@office-kit/xlsx/streaming`
// browser-targeted entry too. Filesystem + Readable / Writable helpers live in
// `./node-fs.ts` (re-exported via `@office-kit/xlsx/node`) where the `node:fs` /
// `node:stream` imports stay out of the browser-safe surface.

import { OpenXmlIoError } from '../utils/exceptions';
import type { BufferedSinkWriter, XlsxSink } from './sink';
import type { XlsxSource } from './source';

/**
 * Wrap a Buffer or Uint8Array as an XlsxSource. The underlying bytes are
 * referenced — no copy — so callers must not mutate them while the source is in
 * use.
 */
export function fromBuffer(buf: Buffer | Uint8Array): XlsxSource {
  if (!(buf instanceof Uint8Array)) {
    throw new OpenXmlIoError('fromBuffer expects a Buffer or Uint8Array');
  }
  // Buffer is a subclass of Uint8Array, so a single normalisation suffices.
  const bytes: Uint8Array = buf instanceof Buffer ? new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength) : buf;
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
 * In-memory Buffer sink. The buffered path concatenates appended chunks into a
 * single allocation when {@link BufferedSinkWriter.finish} resolves; the
 * convenience `result()` returns it as a Node Buffer.
 */
export function toBuffer(): XlsxSink & { toBytes(): BufferedSinkWriter; result(): Buffer } {
  const chunks: Uint8Array[] = [];
  let finalised: Uint8Array | undefined;

  const finalise = (): Uint8Array => {
    if (finalised !== undefined) return finalised;
    let total = 0;
    for (const c of chunks) total += c.byteLength;
    const out = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      out.set(c, off);
      off += c.byteLength;
    }
    finalised = out;
    chunks.length = 0;
    return out;
  };

  return {
    toBytes(): BufferedSinkWriter {
      return {
        write(chunk: Uint8Array): void {
          if (finalised !== undefined) {
            throw new OpenXmlIoError('toBuffer sink: write after finish');
          }
          if (!(chunk instanceof Uint8Array)) {
            throw new OpenXmlIoError('toBuffer sink: chunk is not a Uint8Array');
          }
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
    result(): Buffer {
      const bytes = finalise();
      return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    },
  };
}
