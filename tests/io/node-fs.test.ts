// Node filesystem + Readable/Writable I/O helpers.

import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable, Writable } from 'node:stream';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { fromFile, fromFileSync, fromReadable, toFile, toWritable } from '../../src/io/node-fs';
import { OpenXmlIoError } from '../../src/utils/exceptions';

let scratch: string;
beforeAll(() => {
  scratch = mkdtempSync(join(tmpdir(), 'office-kit-xlsx-io-'));
});
afterAll(() => {
  rmSync(scratch, { recursive: true, force: true });
});

describe('fromFile', () => {
  it('reads bytes via toBytes()', async () => {
    const path = join(scratch, 'a.bin');
    writeFileSync(path, Buffer.from('hello fromFile', 'utf8'));
    const src = fromFile(path);
    const bytes = await src.toBytes();
    expect(new TextDecoder().decode(bytes)).toBe('hello fromFile');
  });

  it('streams via toStream() and yields the same bytes', async () => {
    const path = join(scratch, 'b.bin');
    const payload = new Uint8Array(64 * 1024);
    for (let i = 0; i < payload.byteLength; i++) payload[i] = i & 0xff;
    writeFileSync(path, payload);
    const src = fromFile(path);
    const stream = src.toStream?.();
    if (!stream) throw new Error('expected stream');
    const chunks: Uint8Array[] = [];
    const reader = stream.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    let total = 0;
    for (const c of chunks) total += c.byteLength;
    const out = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      out.set(c, off);
      off += c.byteLength;
    }
    expect(out).toEqual(payload);
  });

  it('rejects empty paths with OpenXmlIoError', () => {
    expect(() => fromFile('')).toThrowError(OpenXmlIoError);
  });

  it('surfaces ENOENT through toBytes() as OpenXmlIoError', async () => {
    const src = fromFile(join(scratch, 'does-not-exist.bin'));
    await expect(src.toBytes()).rejects.toBeInstanceOf(OpenXmlIoError);
  });
});

describe('fromFileSync', () => {
  it('reads file synchronously', async () => {
    const path = join(scratch, 'sync.bin');
    writeFileSync(path, Buffer.from('sync hello', 'utf8'));
    const src = fromFileSync(path);
    const bytes = await src.toBytes();
    expect(new TextDecoder().decode(bytes)).toBe('sync hello');
  });

  it('throws OpenXmlIoError on missing file', () => {
    expect(() => fromFileSync(join(scratch, 'nope.bin'))).toThrowError(OpenXmlIoError);
  });
});

describe('toFile', () => {
  it('writes incoming chunks to disk and returns the path via result()', async () => {
    const path = join(scratch, 'out.bin');
    const sink = toFile(path);
    const w = sink.toBytes();
    w.write(new Uint8Array([1, 2, 3]));
    w.write(new Uint8Array([4, 5]));
    await w.finish();
    expect(sink.result()).toBe(path);
    expect(Array.from(readFileSync(path))).toEqual([1, 2, 3, 4, 5]);
  });

  it('rejects writes after finish()', async () => {
    const path = join(scratch, 'closed.bin');
    const sink = toFile(path);
    const w = sink.toBytes();
    w.write(new Uint8Array([1]));
    await w.finish();
    expect(() => w.write(new Uint8Array([2]))).toThrowError(OpenXmlIoError);
  });

  it('rejects empty paths', () => {
    expect(() => toFile('')).toThrowError(OpenXmlIoError);
  });

  it('honours writable backpressure: writable.writableLength stays at one chunk at a time', async () => {
    // Build a custom Writable with a small highWaterMark and a deliberately
    // slow _write so the sink's per-chunk drain wait is observable. Without
    // the queued-write change, all N sink.write() calls would call
    // writable.write() immediately, so the internal buffer would grow to
    // ~N*chunkSize. With the fix, writableLength never holds more than one
    // chunk because the next chunk only goes in after drain.
    const chunkSize = 1024;
    const chunkCount = 32;
    const maxLengthSamples: number[] = [];
    const writes: Uint8Array[] = [];
    const writable = new Writable({
      highWaterMark: chunkSize, // small budget — chunkSize exactly matches one write
      write(chunk, _enc, cb) {
        writes.push(chunk instanceof Uint8Array ? new Uint8Array(chunk) : new Uint8Array(chunk));
        // Capture the *post-acceptance* buffer depth: how many additional
        // chunks are waiting in the writable beyond the one we're processing.
        maxLengthSamples.push(writable.writableLength);
        setImmediate(cb);
      },
    });
    const sink = toWritable(writable);
    const w = sink.toBytes();
    for (let i = 0; i < chunkCount; i++) {
      w.write(new Uint8Array(chunkSize).fill(i & 0xff));
    }
    await w.finish();
    expect(writes).toHaveLength(chunkCount);
    // No more than one chunk should ever be queued in the writable beyond the
    // one currently being processed. (Allow a small slack to account for the
    // last chunk landing while the drain handler is still settling.)
    const maxQueued = Math.max(...maxLengthSamples);
    expect(maxQueued).toBeLessThanOrEqual(chunkSize);
  });
});

describe('fromReadable', () => {
  it('collects a Node Readable into a single Uint8Array', async () => {
    const r = Readable.from([Buffer.from('foo'), Buffer.from('bar'), Buffer.from('baz')]);
    const src = fromReadable(r);
    const bytes = await src.toBytes();
    expect(new TextDecoder().decode(bytes)).toBe('foobarbaz');
  });

  it('rejects non-Readable inputs', () => {
    // @ts-expect-error deliberately wrong
    expect(() => fromReadable('not a stream')).toThrowError(OpenXmlIoError);
  });
});

describe('toWritable', () => {
  it('forwards chunks to an underlying Writable', async () => {
    const collected: Uint8Array[] = [];
    const w = new Writable({
      write(chunk, _enc, cb) {
        collected.push(chunk instanceof Uint8Array ? new Uint8Array(chunk) : new Uint8Array(chunk));
        cb();
      },
    });
    const sink = toWritable(w);
    const sw = sink.toBytes();
    sw.write(new Uint8Array([1, 2]));
    sw.write(new Uint8Array([3, 4, 5]));
    await sw.finish();
    let total = 0;
    for (const c of collected) total += c.byteLength;
    expect(total).toBe(5);
    expect(sink.result()).toBe(w);
  });

  it('rejects non-Writable inputs', () => {
    // @ts-expect-error deliberately wrong
    expect(() => toWritable({})).toThrowError(OpenXmlIoError);
  });
});
