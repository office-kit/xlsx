import { readFileSync } from 'node:fs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { normalizeStrictArchive } from '../../src/io/strict.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { OpenXmlContentLimitError, OpenXmlNotImplementedError, OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { getCell } from '../../src/worksheet/worksheet.js';
import { parseXml } from '../../src/xml/parser.js';
import { iterParse } from '../../src/xml/iterparse.js';
import { openZip, type ZipArchive } from '../../src/zip/reader.js';

const strictNs = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const transitionalNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const fixture = readFileSync(new URL('./fixtures/strict/SimpleStrict.xlsx', import.meta.url));
const sheetPath = 'xl/worksheets/sheet2.xml';
const edit = (parts: Record<string, string>): Uint8Array => {
  const zip = unzipSync(fixture);
  for (const [path, xml] of Object.entries(parts)) zip[path] = strToU8(xml);
  return zipSync(zip);
};
const original = (path: string): string => {
  const bytes = unzipSync(fixture)[path];
  if (!bytes) throw new Error(`missing fixture part ${path}`);
  return strFromU8(bytes);
};
const dated = (value: string, attrs = '', formula = ''): string =>
  original(sheetPath).replace(/<c r="A4"[^>]*>.*?<\/c>/, `<c r="A4" s="1" t="d" ${attrs}>${formula}<v>${value}</v></c>`);

async function values(bytes: Uint8Array): Promise<unknown[]> {
  const wb = await loadWorkbook(fromBuffer(bytes));
  const ref = wb.sheets[1];
  if (ref?.kind !== 'worksheet') throw new Error('expected worksheet fixture');
  const result: unknown[] = [getCell(ref.sheet, 4, 1)?.value];
  const stream = await loadWorkbookStream(fromBuffer(bytes));
  try {
    for (const options of [{}, { minRow: 4, maxRow: 4 }]) {
      for await (const row of stream.openWorksheet('Sheet Number 2').iterRows(options)) {
        const cell = row.find((c) => c.row === 4 && c.col === 1);
        if (cell) result.push(cell.value);
      }
    }
  } finally { await stream.close(); }
  return result;
}

async function rejectsBoth(bytes: Uint8Array, error: typeof OpenXmlSchemaError | typeof OpenXmlNotImplementedError): Promise<void> {
  await expect(loadWorkbook(fromBuffer(bytes))).rejects.toBeInstanceOf(error);
  const stream = await loadWorkbookStream(fromBuffer(bytes));
  try {
    for (const options of [{}, { minRow: 4, maxRow: 4 }]) {
      await expect((async () => {
        for await (const _row of stream.openWorksheet('Sheet Number 2').iterRows(options)) { /* drain */ }
      })()).rejects.toBeInstanceOf(error);
    }
  } finally { await stream.close(); }
}

describe('Strict dates use the existing serial value model', () => {
  it.each([
    ['0', '1', 32874], ['false', 'true', 32874],
    ['1', '1', 31412], ['true', 'false', 32874],
  ])('dateCompatibility=%s date1904=%s', async (compat, epoch, expected) => {
    const workbook = original('xl/workbook.xml').replace('dateCompatibility="0"', `dateCompatibility="${compat}" date1904="${epoch}"`);
    const bytes = edit({ 'xl/workbook.xml': workbook });
    expect(await values(bytes)).toEqual([expected, expected, expected]);
    const wb = await loadWorkbook(fromBuffer(bytes));
    expect(await values(await workbookToBytes(wb))).toEqual([expected, expected, expected]);
  });
  it.each(['1990-01-01T12:00:00', '1990-01-01T12:00:00Z', '1990-01-01T14:00:00+02:00'])('reads %s in UTC', async (date) => {
    expect(await values(edit({ [sheetPath]: dated(date) }))).toEqual([32874.5, 32874.5, 32874.5]);
  });
  it('converts an ISO formula cache without altering the formula', async () => {
    const result = await values(edit({ [sheetPath]: dated('1990-01-01', '', '<f>DATE(1990,1,1)</f>') }));
    expect(result[0]).toMatchObject({ kind: 'formula', formula: 'DATE(1990,1,1)', cachedValue: 32874 });
    expect(result.slice(1)).toEqual([32874, 32874]);
  });
  it.each(['1990-02-30', '2019-02-29', '1990-01-01T25:00:00', '1990-01-01T12:00:00+00:70', '1990-01-01T12:00:00+14:01', 'not a date'])('rejects malformed ISO date %s', async (date) => {
    await rejectsBoth(edit({ [sheetPath]: dated(date) }), OpenXmlSchemaError);
  });
  it.each(['1899-12-31', '1900-02-28', '1990-01-01T12:00:00.0001'])('refuses unrepresentable date %s', async (date) => {
    await rejectsBoth(edit({ [sheetPath]: dated(date) }), OpenXmlNotImplementedError);
  });
  it('refuses ambiguous early numeric dates under the ISO date base', async () => {
    await rejectsBoth(edit({ [sheetPath]: dated('60').replace('t="d"', 't="n"') }), OpenXmlNotImplementedError);
  });
  it('keeps backward-compatible serials when dateCompatibility is true', async () => {
    const bytes = edit({
      [sheetPath]: dated('60').replace('t="d"', 't="n"'),
      'xl/workbook.xml': original('xl/workbook.xml').replace('dateCompatibility="0"', 'dateCompatibility="1"'),
    });
    expect(await values(bytes)).toEqual([60, 60, 60]);
  });
});

it('leaves the general XML parser and event parser namespace-faithful', async () => {
  const xml = `<worksheet xmlns="${strictNs}"/>`;
  expect(parseXml(xml).name).toBe(`{${strictNs}}worksheet`);
  const events = [];
  for await (const event of iterParse(xml)) events.push(event);
  expect(JSON.stringify(events)).toContain(strictNs);
});

it('preserves namespace-looking strings and custom XML attributes', async () => {
  const uri = 'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet';
  const xml = original(sheetPath).replace(/<c r="A4"[^>]*>.*?<\/c>/, `<c r="A4" t="inlineStr"><is><t>${uri} &amp; café 東京</t></is></c>`);
  expect(await values(edit({ [sheetPath]: xml }))).toEqual(new Array(3).fill(`${uri} & café 東京`));
  const raw = await openZip(fromBuffer(edit({ 'customXml/test.xml': `<data xmlns:p="${strictNs}" target="${uri}"><p:unknown><![CDATA[${uri}]]></p:unknown></data>` })));
  try {
    const normalized = strFromU8(normalizeStrictArchive(raw).read('customXml/test.xml'));
    expect(normalized).toContain(`target="${uri}"`);
    expect(parseXml(normalized).children[0]?.text).toBe(uri);
  } finally { raw.close(); }
});

it('normalizes escaped namespace declarations', async () => {
  const xml = original(sheetPath).replaceAll(strictNs, strictNs.replace('http:', 'http&#58;'));
  expect(await values(edit({ [sheetPath]: xml }))).toEqual([32874, 32874, 32874]);
});

it('keeps content limits active for Strict input', async () => {
  await expect(loadWorkbook(fromBuffer(fixture), { contentLimits: { maxCells: 1 } })).rejects.toBeInstanceOf(OpenXmlContentLimitError);
  const stream = await loadWorkbookStream(fromBuffer(fixture), { contentLimits: { maxCells: 1 } });
  try {
    await expect((async () => {
      for await (const _row of stream.openWorksheet('Sheet Number 2').iterRows()) { /* drain */ }
    })()).rejects.toBeInstanceOf(OpenXmlContentLimitError);
  } finally { await stream.close(); }
});

function chunkedArchive(xml: string, chunkSize: number, cancelled: () => void = () => {}): ZipArchive {
  const bytes = strToU8(xml);
  return {
    list: () => ['sheet.xml'], has: (p) => p === 'sheet.xml', close() {},
    read: () => bytes.slice(), readAsync: async () => bytes.slice(),
    readStream: () => {
      let offset = 0;
      return new ReadableStream({
        pull(controller) {
          if (offset === bytes.length) { controller.close(); return; }
          controller.enqueue(bytes.slice(offset, offset + chunkSize));
          offset = Math.min(bytes.length, offset + chunkSize);
        },
        cancel: cancelled,
      });
    },
  };
}

it('handles namespace declarations and UTF-8 split across single-byte chunks', async () => {
  const xml = `<s:worksheet xmlns:s="${strictNs}"><s:sheetData><s:row r="1"><s:c r="A1" t="inlineStr"><s:is><s:t>café 東京 &amp; hi</s:t></s:is></s:c></s:row></s:sheetData></s:worksheet>`;
  const archive = normalizeStrictArchive(chunkedArchive(xml, 1));
  const result = await new Response(archive.readStream('sheet.xml')).text();
  expect(parseXml(result)).toEqual(parseXml(xml.replaceAll(strictNs, transitionalNs)));
});

it('propagates cancellation without consuming the whole worksheet', async () => {
  let cancelled = false;
  const xml = `<worksheet xmlns="${strictNs}"><sheetData>${'<row r="1"/>'.repeat(10000)}</sheetData></worksheet>`;
  const stream = normalizeStrictArchive(chunkedArchive(xml, 128, () => { cancelled = true; })).readStream('sheet.xml');
  const reader = stream.getReader();
  await reader.read();
  await reader.cancel();
  await new Promise((resolve) => { setTimeout(resolve, 0); });
  expect(cancelled).toBe(true);
});

it.each([
  `<!DOCTYPE worksheet [<!ENTITY x "secret">]><worksheet xmlns="${strictNs}">&x;</worksheet>`,
  `<worksheet xmlns="${strictNs}"><broken></worksheet>`,
])('rejects unsafe or malformed XML before and after stream chunk boundaries', async (xml) => {
  const archive = normalizeStrictArchive(chunkedArchive(xml, 1));
  expect(() => archive.read('sheet.xml')).toThrow(OpenXmlSchemaError);
  await expect(new Response(archive.readStream('sheet.xml')).text()).rejects.toBeInstanceOf(OpenXmlSchemaError);
});

it.each([
  `<worksheet xmlns="http://purl.oclc.org/ooxml/unknown/main"/>`,
  `<pivotCacheDefinition xmlns="${strictNs}"/>`,
  `<worksheet xmlns="${strictNs}"><alignment horizontal="start"/></worksheet>`,
])('refuses unsupported Strict content explicitly', async (xml) => {
  const archive = normalizeStrictArchive(chunkedArchive(xml, 7));
  expect(() => archive.read('sheet.xml')).toThrow(OpenXmlNotImplementedError);
  await expect(new Response(archive.readStream('sheet.xml')).text()).rejects.toBeInstanceOf(OpenXmlNotImplementedError);
});

it.each(['1990-01-01T00:00:00.029Z', '1990-01-01T00:00:00.057000Z'])('keeps millisecond precision for %s', async (date) => {
  const expected = (Date.parse(date) - Date.UTC(1899, 11, 30)) / 86400000;
  for (const value of await values(edit({ [sheetPath]: dated(date) }))) expect(value).toBeCloseTo(expected, 9);
});

it('preserves external relationship targets and binary parts', async () => {
  const uri = 'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet';
  const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink" Target="${uri}" TargetMode="External"/></Relationships>`;
  const raw = await openZip(fromBuffer(edit({ 'custom.rels': rels })));
  try {
    const archive = normalizeStrictArchive(raw);
    expect(parseXml(archive.read('custom.rels')).children[0]?.attrs).toMatchObject({
      Target: uri, Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    });
    for (const path of raw.list().filter((p) => p.endsWith('.bin'))) expect(archive.read(path)).toEqual(raw.read(path));
  } finally { raw.close(); }
});

it.each(['h:mm:ss', '[h]:mm:ss'])('keeps calendar-independent numeric times with %s', async (format) => {
  const styles = original('xl/styles.xml').replace('numFmtId="14"', `numFmtId="${format === 'h:mm:ss' ? 21 : 46}"`);
  expect(styles).not.toBe(original('xl/styles.xml'));
  expect(await values(edit({ 'xl/styles.xml': styles, [sheetPath]: dated('0.5').replace('t="d"', 't="n"') }))).toEqual([0.5, 0.5, 0.5]);
});

it('maps Strict border start/end to Transitional left/right', async () => {
  const styles = original('xl/styles.xml').replace('<start/>', '<start style="thin"><color rgb="FFFF0000"/></start>').replace('<end/>', '<end style="double"/>');
  const wb = await loadWorkbook(fromBuffer(edit({ 'xl/styles.xml': styles })));
  expect(wb.styles.borders[0]).toMatchObject({ left: { style: 'thin', color: { rgb: 'FFFF0000' } }, right: { style: 'double' } });
});

it('refuses unit-bearing DrawingML coordinates instead of dropping the unit', async () => {
  const xml = '<a:theme xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"><a:off x="1cm" y="0"/></a:theme>';
  const archive = normalizeStrictArchive(chunkedArchive(xml, 3));
  expect(() => archive.read('sheet.xml')).toThrow(OpenXmlNotImplementedError);
  await expect(new Response(archive.readStream('sheet.xml')).text()).rejects.toBeInstanceOf(OpenXmlNotImplementedError);
});

it.each(['12:30:00', 'PT1H'])('explicitly rejects unsupported ISO value %s', async (value) => {
  await rejectsBoth(edit({ [sheetPath]: dated(value) }), OpenXmlNotImplementedError);
});
