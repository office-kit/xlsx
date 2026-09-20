import { unzipSync } from 'fflate';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { OpenXmlIoError } from '../../src/utils/exceptions.js';
import { openZip } from '../../src/zip/reader.js';
import { createZipWriter } from '../../src/zip/writer.js';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');
const EMPTY_XLSX = resolve(FIXTURES, 'empty.xlsx');

const utf8 = (s: string): Uint8Array => new TextEncoder().encode(s);

describe('createZipWriter (basic)', () => {
  it('produces a zip readable by openZip with the same bytes', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('hello.txt', utf8('hello'));
    await writer.addEntry('a/b.txt', utf8('nested'));
    await writer.finalize();

    const zip = await openZip(fromBuffer(sink.result()));
    expect(zip.list()).toEqual(['a/b.txt', 'hello.txt']);
    expect(new TextDecoder().decode(zip.read('hello.txt'))).toBe('hello');
    expect(new TextDecoder().decode(zip.read('a/b.txt'))).toBe('nested');
  });

  it.each([0, 1, 6, 9] as const)('preserves sparse binary entries at compression level %i', async (compressionLevel) => {
    const payload = new Uint8Array(72);
    payload.set([116, 101], 4);
    const sink = toBuffer();
    const writer = createZipWriter(sink, { compressionLevel });
    await writer.addEntry('printerSettings.bin', payload);
    const bytes = await writer.finalize();
    expect(unzipSync(bytes)['printerSettings.bin']).toEqual(payload);
    const archive = await openZip(fromBuffer(bytes));
    try { expect(archive.read('printerSettings.bin')).toEqual(payload); }
    finally { archive.close(); }
  });

  it('honours compress: false (STORE) — bytes still round-trip', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    const payload = new Uint8Array(1024).fill(0xab);
    await writer.addEntry('xl/media/image1.png', payload, { compress: false });
    await writer.finalize();

    const zip = await openZip(fromBuffer(sink.result()));
    expect(zip.read('xl/media/image1.png')).toEqual(payload);
  });

  it('handles 0-byte entries', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('empty.bin', new Uint8Array(0));
    await writer.finalize();

    const zip = await openZip(fromBuffer(sink.result()));
    expect(zip.read('empty.bin').byteLength).toBe(0);
  });

  it('finalize() is idempotent and yields the same bytes', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('a', utf8('A'));
    const first = await writer.finalize();
    const second = await writer.finalize();
    expect(second).toBe(first);
  });

  it('rejects addEntry after finalize', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('a', utf8('A'));
    await writer.finalize();
    await expect(writer.addEntry('b', utf8('B'))).rejects.toBeInstanceOf(OpenXmlIoError);
  });

  it('rejects duplicate entry paths', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('a', utf8('A'));
    await expect(writer.addEntry('a', utf8('A2'))).rejects.toBeInstanceOf(OpenXmlIoError);
  });

  it('rejects ReadableStream payload (deferred to streaming writer)', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(utf8('hi'));
        controller.close();
      },
    });
    await expect(writer.addEntry('a', stream)).rejects.toBeInstanceOf(OpenXmlIoError);
  });

  it('rejects sinks without a buffered toBytes()', () => {
    // `XlsxSink.toBytes` is required by the type, so a sink without it can
    // only be constructed via a cast. The runtime then surfaces the missing
    // method as a TypeError on the first call.
    // biome-ignore lint/suspicious/noExplicitAny: deliberate type-erasing cast
    const stubSink = {} as any;
    expect(() => createZipWriter(stubSink)).toThrow(TypeError);
  });
});

describe('createZipWriter (streaming behaviour)', () => {
  it('writes chunks to the sink as addEntry runs, not all at once on finalize', async () => {
    // Use a custom sink that reports each write timing relative to finalize.
    let finalizeStarted = false;
    const writes: Array<{ during: 'addEntry' | 'finalize'; size: number }> = [];
    const customSink: Parameters<typeof createZipWriter>[0] = {
      toBytes() {
        const chunks: Uint8Array[] = [];
        return {
          write(chunk) {
            writes.push({ during: finalizeStarted ? 'finalize' : 'addEntry', size: chunk.byteLength });
            chunks.push(chunk);
          },
          async finish() {
            let total = 0;
            for (const c of chunks) total += c.byteLength;
            const out = new Uint8Array(total);
            let off = 0;
            for (const c of chunks) {
              out.set(c, off);
              off += c.byteLength;
            }
            return out;
          },
        };
      },
    };

    const writer = createZipWriter(customSink);
    // Each entry deflates and emits chunks as the local file header + data
    // are built — those writes must land before finalize is even called.
    const big = new Uint8Array(64 * 1024);
    for (let i = 0; i < big.byteLength; i++) big[i] = i & 0xff;
    await writer.addEntry('a.bin', big, { compress: false });
    await writer.addEntry('b.bin', big, { compress: true });
    expect(writes.some((w) => w.during === 'addEntry')).toBe(true);

    finalizeStarted = true;
    await writer.finalize();
    // Final central directory chunk arrives during finalize.
    expect(writes.some((w) => w.during === 'finalize')).toBe(true);
  });
});

describe('createZipWriter (round-trip via openZip)', () => {
  it('every entry of empty.xlsx round-trips byte-identically through the writer', async () => {
    const original = await openZip(fromBuffer(readFileSync(EMPTY_XLSX)));
    const paths = original.list();

    const sink = toBuffer();
    const writer = createZipWriter(sink);
    for (const path of paths) {
      // OOXML xml parts are deflate by default; use STORE for media + vba.
      const compress = !(path.startsWith('xl/media/') || path === 'xl/vbaProject.bin');
      await writer.addEntry(path, original.read(path), { compress });
    }
    await writer.finalize();

    const round = await openZip(fromBuffer(sink.result()));
    expect(round.list()).toEqual(paths);
    for (const path of paths) {
      expect(round.read(path)).toEqual(original.read(path));
    }
  });
});
