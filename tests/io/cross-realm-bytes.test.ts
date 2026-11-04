// Bytes built in another realm (jsdom, iframes, Node `vm`) fail `instanceof Uint8Array`.
// Saving and loading must still accept them.

import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import { fromArrayBuffer } from '../../src/io/browser.js';
import { loadWorkbook, workbookToBytes } from '../../src/io/index.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { isArrayBuffer, isUint8Array } from '../../src/utils/bytes.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/index.js';
import { setCell } from '../../src/worksheet/index.js';
import { createZipWriter } from '../../src/zip/index.js';

const foreignCopy = (bytes: Uint8Array): Uint8Array => {
  const copy: unknown = runInNewContext('new Uint8Array(source)', { source: Array.from(bytes) });
  if (!isUint8Array(copy)) throw new Error('expected a Uint8Array');
  return copy;
};

describe('cross-realm byte arrays', () => {
  it('recognises a Uint8Array and an ArrayBuffer from another realm', () => {
    const foreign = foreignCopy(new Uint8Array([1, 2]));
    expect(foreign instanceof Uint8Array).toBe(false);
    expect(isUint8Array(foreign)).toBe(true);
    expect(isUint8Array(Buffer.from([1]))).toBe(true);
    expect(isUint8Array(new Uint16Array(1))).toBe(false);
    expect(isUint8Array(new DataView(new ArrayBuffer(1)))).toBe(false);
    expect(isArrayBuffer(runInNewContext('new ArrayBuffer(4)'))).toBe(true);
    expect(isArrayBuffer(new Uint8Array(4))).toBe(false);
  });

  it('writes a zip entry whose bytes come from another realm', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('a.txt', foreignCopy(new TextEncoder().encode('hello')));
    await writer.finalize();
    expect(sink.result().byteLength).toBeGreaterThan(0);
  });

  it('loads a workbook from bytes that come from another realm', async () => {
    const wb = createWorkbook();
    setCell(addWorksheet(wb, 'S'), 1, 1, 'hello');
    const foreign = foreignCopy(await workbookToBytes(wb));
    const fromNode = await loadWorkbook(fromBuffer(foreign));
    const fromBrowser = await loadWorkbook(fromArrayBuffer(foreign));
    for (const loaded of [fromNode, fromBrowser]) {
      expect(getSheet(loaded, 'S')?.rows.get(1)?.get(1)?.value).toBe('hello');
    }
  });
});
