import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromArrayBuffer, fromBlob, fromResponse, fromStream } from '../../src/io/browser.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { getCellDate } from '../../src/styles/display-text.js';
import { getCellBorder, getCellFont } from '../../src/styles/cell-style.js';
import { getSheet } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';
import { required } from './required.js';
import { canonicalValue, corpus, packageFor, SML } from './corpus.js';
import { validateXlsx } from './validate.js';

const output = process.env['QA_CORPUS_OUTPUT'];
if (output) {
  mkdirSync(output, { recursive: true });
  writeFileSync(join(output, 'manifest.json'), JSON.stringify(corpus));
}
function chunks(bytes: Uint8Array, size: number): ReadableStream<Uint8Array> {
  let offset = 0;
  return new ReadableStream({ pull(controller) {
    if (offset >= bytes.length) { controller.close(); return; }
    controller.enqueue(bytes.slice(offset, offset + size)); offset += size;
  } });
}

describe('normative corpus: independent expected semantics', () => {
  it.each(corpus)('$id [$clause]', async c => {
    const bytes = packageFor(c);
    expect((await validateXlsx(bytes)).status).toBe('valid');
    const wb = await loadWorkbook(fromBuffer(bytes));
    const sheetName = c.sheetName?.replace('&apos;', "'") ?? 'Audit';
    const ws = required(getSheet(wb, sheetName));
    const cell = required(getCell(ws, 1, 1));
    const expected = c.expectedValue ?? 'audit';
    expect(canonicalValue(cell.value)).toEqual(expected);
    if (c.expectedDate) expect(getCellDate(wb, cell)?.toISOString()).toBe(c.expectedDate);
    if (c.expectedNames) expect(wb.definedNames).toEqual(c.expectedNames);
    if (c.expectedDate1904 !== undefined) expect(wb.date1904).toBe(c.expectedDate1904);
    if (c.expectedFont) expect(getCellFont(wb, cell)).toMatchObject(c.expectedFont);
    for (const key of c.expectedMissingFont ?? []) expect(getCellFont(wb, cell)[key]).toBeUndefined();
    if (c.expectedBorder) expect(getCellBorder(wb, cell).left?.style).toBe(c.expectedBorder);
    // An unrelated edit must preserve semantics that were independently specified.
    setCell(ws, 2, 2, 'edited');
    const saved = await workbookToBytes(wb);
    expect((await validateXlsx(saved)).status).toBe('valid');
    const reloaded = await loadWorkbook(fromBuffer(saved));
    const reloadedSheet = required(getSheet(reloaded, sheetName));
    const original = required(getCell(reloadedSheet, 1, 1));
    expect(canonicalValue(original.value)).toEqual(expected);
    if (c.expectedDate) expect(getCellDate(reloaded, original)?.toISOString()).toBe(c.expectedDate);
    if (c.expectedNames) expect(reloaded.definedNames).toEqual(c.expectedNames);
    if (c.expectedDate1904 !== undefined) expect(reloaded.date1904).toBe(c.expectedDate1904);
    for (const [i, value] of (c.expectedValues ?? [expected]).entries()) {
      expect(canonicalValue(getCell(reloadedSheet, i + 1, 1)?.value)).toEqual(value);
    }
    if (c.expectedFont) expect(getCellFont(reloaded, original)).toMatchObject(c.expectedFont);
    for (const key of c.expectedMissingFont ?? []) expect(getCellFont(reloaded, original)[key]).toBeUndefined();
    if (c.expectedBorder) expect(getCellBorder(reloaded, original).left?.style).toBe(c.expectedBorder);
    const stream = await loadWorkbookStream(fromStream(chunks(bytes, 7)));
    const values = [];
    try { for await (const row of stream.openWorksheet(sheetName).iterRows()) values.push(canonicalValue(row[0]?.value)); }
    finally { await stream.close(); }
    expect(values).toEqual(c.expectedStreamValues ?? c.expectedValues ?? [expected]);
    if (c.expectedDate1904 !== undefined) expect(stream.date1904).toBe(c.expectedDate1904);
    if (output) { writeFileSync(join(output, `${c.id}.input.xlsx`), bytes); writeFileSync(join(output, `${c.id}.output.xlsx`), saved); }
  });
  it.each([1, 7, 257])('all browser adapters agree with expected text at chunk size %i', async size => {
    const c = required(corpus.find(candidate => candidate.id === 'inline-character-references'));
    const bytes = packageFor(c);
    for (const source of [fromArrayBuffer(bytes), fromBlob(new Blob([new Uint8Array(bytes).buffer])), fromResponse(new Response(chunks(bytes, size))), fromStream(chunks(bytes, size))]) {
      const wb = await loadWorkbook(source);
      expect(getCell(required(getSheet(wb, c.sheetName?.replace('&apos;', "'") ?? 'Audit')), 1, 1)?.value).toBe(c.expectedValue);
    }
  });
  it('namespace prefixes and quoting preserve semantics', async () => {
    const c = required(corpus.find(candidate => candidate.id === 'inline-character-references'));
    const parts = unzipSync(packageFor(c));
    const path = 'xl/worksheets/sheet1.xml';
    parts[path] = strToU8(strFromU8(required(parts[path])).replace(`xmlns="${SML}"`, `xmlns:s="${SML}"`).replace(/<(\/?)(worksheet|dimension|sheetData|row|c|is|t)\b/g, '<$1s:$2').replace(/"/g, "'"));
    const bytes = zipSync(parts);
    expect((await validateXlsx(bytes)).status).toBe('valid');
    const wb = await loadWorkbook(fromBuffer(bytes));
    expect(getCell(required(getSheet(wb, c.sheetName?.replace('&apos;', "'") ?? 'Audit')), 1, 1)?.value).toBe(c.expectedValue);
  });
});
