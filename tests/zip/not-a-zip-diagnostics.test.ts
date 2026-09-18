// What openZip says about input that is not a zip. Upload paths validate by
// file extension, so a CSV or a PDF renamed to .xlsx reaches the loader and
// "archive is not a valid zip" is true but tells its user nothing.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import {
  OpenXmlError,
  OpenXmlIoError,
  OpenXmlNotImplementedError,
  OpenXmlSchemaError,
} from '../../src/utils/exceptions.js';
import { openZip } from '../../src/zip/reader.js';
import { createZipWriter } from '../../src/zip/writer.js';

const ascii = (s: string): Uint8Array => new TextEncoder().encode(s);

const openError = async (bytes: Uint8Array): Promise<Error> => {
  try {
    await openZip(fromBuffer(bytes));
  } catch (err) {
    return err as Error;
  }
  throw new Error('expected openZip to reject');
};

describe('openZip names what the leading bytes look like', () => {
  it('CSV content reads as plain text', async () => {
    const err = await openError(ascii('id,name,amount\n1,Widget,4.50\n2,Gadget,9.00\n'));
    expect(err).toBeInstanceOf(OpenXmlIoError);
    expect(err.message).toMatch(/not a valid zip/);
    expect(err.message).toMatch(/plain text/);
    expect(err.message).toMatch(/CSV/);
  });

  it('a UTF-8 byte-order mark is named as such', async () => {
    const bom = new Uint8Array([0xef, 0xbb, 0xbf, ...ascii('id,name\n1,Widget\n')]);
    const err = await openError(bom);
    expect(err.message).toMatch(/UTF-8 byte-order mark/);
  });

  it('a UTF-16 byte-order mark is named as such', async () => {
    const utf16 = new Uint8Array([0xff, 0xfe, 0x69, 0x00, 0x64, 0x00, 0x0a, 0x00]);
    const err = await openError(utf16);
    expect(err.message).toMatch(/UTF-16 byte-order mark/);
  });

  it('a PDF is named as a PDF', async () => {
    const pdf = new Uint8Array([...ascii('%PDF-1.7\n'), 0x80, 0x01, 0x02, 0x03]);
    const err = await openError(pdf);
    expect(err.message).toMatch(/a PDF/);
  });

  it('a zip prefix with no central directory reads as truncated', async () => {
    const truncated = new Uint8Array(64);
    truncated.set([0x50, 0x4b, 0x03, 0x04]);
    const err = await openError(truncated);
    expect(err.message).toMatch(/truncated or partially uploaded/);
  });

  it('says nothing extra about bytes no magic number identifies', async () => {
    const random = new Uint8Array(64);
    for (let i = 0; i < random.length; i++) random[i] = (i * 37 + 129) & 0xff;
    const err = await openError(random);
    expect(err.message).toMatch(/not a valid zip/);
    expect(err.message).not.toMatch(/leading bytes/);
  });

  it('an empty file reports its length against the minimum record size', async () => {
    const err = await openError(new Uint8Array(0));
    expect(err).toBeInstanceOf(OpenXmlIoError);
    expect(err.message).toMatch(/is 0 bytes/);
    expect(err.message).toMatch(/22-byte minimum/);
  });

  it('a too-short text file is still named as text', async () => {
    const err = await openError(ascii('a,b\n'));
    expect(err.message).toMatch(/22-byte minimum/);
    expect(err.message).toMatch(/plain text/);
  });

  it('an OLE compound-document container names both formats it can be', async () => {
    const cfb = new Uint8Array(64);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    const err = await openError(cfb);
    expect(err).toBeInstanceOf(OpenXmlNotImplementedError);
    expect(err.message).toMatch(/Encrypted xlsx/);
    expect(err.message).toMatch(/\.xls/);
  });

  it('the diagnosis survives the loadWorkbook path, not just openZip', async () => {
    await expect(loadWorkbook(fromBuffer(ascii('id,name\n1,Widget\n')))).rejects.toThrow(
      /plain text/,
    );
  });
});

// Pins the classes the README's error table and the loadWorkbook docstring
// promise, so the documented contract cannot drift without a test failing.
describe('the class loadWorkbook throws is the contract', () => {
  it('input that is not a zip is an OpenXmlIoError', async () => {
    await expect(loadWorkbook(fromBuffer(ascii('id,name\n1,Widget\n')))).rejects.toBeInstanceOf(
      OpenXmlIoError,
    );
  });

  it('an OLE compound-document container is an OpenXmlNotImplementedError', async () => {
    const cfb = new Uint8Array(64);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    await expect(loadWorkbook(fromBuffer(cfb))).rejects.toBeInstanceOf(OpenXmlNotImplementedError);
  });

  it('a readable zip that is not an OPC package is an OpenXmlSchemaError', async () => {
    const sink = toBuffer();
    const writer = createZipWriter(sink);
    await writer.addEntry('hello.txt', ascii('a zip, but not a package'));
    await writer.finalize();
    await expect(loadWorkbook(fromBuffer(sink.result()))).rejects.toBeInstanceOf(
      OpenXmlSchemaError,
    );
  });

  it('every one of them is an OpenXmlError, which is what a caller catches', async () => {
    const err = await openError(ascii('id,name\n1,Widget\n'));
    expect(err).toBeInstanceOf(OpenXmlError);
  });
});
