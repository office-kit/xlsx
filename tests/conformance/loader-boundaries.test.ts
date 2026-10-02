import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { expect, it } from 'vitest';
import { fromStream } from '../../src/io/browser.js';
import { loadWorkbook, type LoadOptions } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { OpenXmlContentLimitError, OpenXmlDecompressionBombError, OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { getCell } from '../../src/worksheet/worksheet.js';
import { packageFor } from './corpus.js';
import { required } from './required.js';

const sheetPath = 'xl/worksheets/sheet1.xml';
function input(): Uint8Array {
  return packageFor({ id: 'loader-boundaries', clause: 'public loader resource contract', rowsXml: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>' });
}
function change(before: string, after: string): Uint8Array {
  const parts = unzipSync(input());
  parts[sheetPath] = strToU8(strFromU8(required(parts[sheetPath])).replace(before, after));
  return zipSync(parts);
}
for (const mode of ['model', 'stream'] as const) for (const chunk of [0, 1, 7, 257]) {
  const read = async (bytes: Uint8Array, options: LoadOptions = {}): Promise<unknown[]> => {
    let offset = 0;
    const source = chunk === 0 ? fromBuffer(bytes) : fromStream(new ReadableStream<Uint8Array>({ pull(controller) {
      if (offset === bytes.length) controller.close();
      else { controller.enqueue(bytes.slice(offset, offset + chunk)); offset = Math.min(bytes.length, offset + chunk); }
    } }));
    if (mode === 'model') {
      const wb = await loadWorkbook(source, options); const ref = required(wb.sheets[0]);
      if (ref.kind !== 'worksheet') throw new Error('Expected worksheet');
      return [getCell(ref.sheet, 1, 1)?.value, getCell(ref.sheet, 1, 2)?.value, getCell(ref.sheet, 2, 1)?.value, getCell(ref.sheet, 2, 2)?.value];
    }
    const wb = await loadWorkbookStream(source, options);
    try { const values: unknown[] = []; for await (const row of wb.openWorksheet('Audit').iterValues()) values.push(...row); return values; }
    finally { await wb.close(); }
  };
  it(`${mode}/${chunk}: accepts exact content limits and ignores misleading dimensions`, async () => {
    const bytes = change('ref="A1:C2"', 'ref="A1:XFD1048576"');
    expect(await read(bytes, { contentLimits: { maxCells: 4, maxRows: 2 } })).toEqual([1, 2, 3, 4]);
    await expect(read(bytes, { contentLimits: { maxCells: 3 } })).rejects.toBeInstanceOf(OpenXmlContentLimitError);
    await expect(read(bytes, { contentLimits: { maxRows: 1 } })).rejects.toBeInstanceOf(OpenXmlContentLimitError);
  });
  it(`${mode}/${chunk}: applies ZIP entry, total and ratio caps at the public boundary`, async () => {
    const bytes = input(); const sizes = Object.values(unzipSync(bytes)).map(b => b.length);
    const largest = Math.max(...sizes); const total = sizes.reduce((a, b) => a + b, 0);
    expect(await read(bytes, { decompressionLimits: { maxEntryUncompressedBytes: largest, maxTotalUncompressedBytes: total } })).toEqual([1, 2, 3, 4]);
    for (const limits of [{ maxEntryUncompressedBytes: largest - 1 }, { maxTotalUncompressedBytes: total - 1 }, { maxCompressionRatio: 1 }]) {
      await expect(read(bytes, { decompressionLimits: limits })).rejects.toBeInstanceOf(OpenXmlDecompressionBombError);
    }
  });
  it.each([
    ['DTD', '<worksheet ', '<!DOCTYPE worksheet [<!ENTITY q "blocked">]><worksheet '],
    ['mismatched tags', '</sheetData>', '</wrong>'],
    ['invalid character reference', '<v>1</v>', '<v>&#0;</v>'],
  ])(`${mode}/${chunk}: rejects %s with the schema error contract`, async (_label, before, after) => {
    await expect(read(change(before, after))).rejects.toBeInstanceOf(OpenXmlSchemaError);
  });
}
