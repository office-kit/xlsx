import { describe, expect, it } from 'vitest';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { getCell } from '../../src/worksheet/worksheet.js';
import { openZip } from '../../src/zip/reader.js';
import { validateXlsx } from '../conformance/validate.js';

describe('write-only string retention', () => {
  it('stops admitting new strings after 100,000 entries, but reuses existing indices across sheets', async () => {
    const sink = toBuffer();
    const wb = await createWriteOnlyWorkbook(sink);
    const first = await wb.addWorksheet('First');
    for (let i = 0; i < 100_000; i++) await first.appendRow([`v${i}`]);
    await first.close();
    const second = await wb.addWorksheet('Second');
    await second.appendRow(['v0', 'new', { kind: 'rich-text', runs: [{ text: 'rich', font: { b: true } }] }]);
    const special = ' \t\n\r\u0001<&>_x0041_😀 ';
    const rich = { kind: 'rich-text' as const, runs: [
      { text: special, font: { b: true } }, { text: '_x0' }, { text: '041_' },
    ] };
    await second.appendRow(['', special, { value: rich, style: { font: { italic: true } } }]);
    await second.close();
    await wb.finalize();
    const zip = await openZip(fromBuffer(sink.result()));
    try {
      const xml = new TextDecoder().decode(await zip.read('xl/worksheets/sheet2.xml'));
      expect(xml).toContain('<c r="A1" t="s"><v>0</v></c>');
      expect(xml).toContain('<c r="B1" t="inlineStr"><is><t>new</t></is></c>');
      expect(xml).toContain('<c r="C1" t="inlineStr"><is><r><rPr><b/></rPr><t>rich</t></r></is></c>');
      expect(xml).toContain('<c r="C2" s="1" t="inlineStr">');
      expect(xml).toContain('xml:space="preserve"');
      const sst = new TextDecoder().decode(await zip.read('xl/sharedStrings.xml'));
      expect(sst.match(/<si>/g)).toHaveLength(100_000);
    } finally {
      zip.close();
    }
    const bytes = sink.result();
    const validation = await validateXlsx(bytes);
    expect(validation.issues).toEqual([]);
    const loaded = await loadWorkbook(fromBuffer(bytes));
    const sheet = loaded.sheets[1];
    if (!sheet || sheet.kind !== 'worksheet') throw new Error('missing worksheet');
    expect(getCell(sheet.sheet, 2, 1)?.value).toBe('');
    expect(getCell(sheet.sheet, 2, 2)?.value).toBe(special);
    expect(getCell(sheet.sheet, 2, 3)?.value).toEqual(rich);
    const streamed = await loadWorkbookStream(fromBuffer(bytes));
    try {
      const values = [];
      for await (const row of streamed.openWorksheet('Second').iterValues()) values.push(row);
      expect(values).toEqual([['v0', 'new', 'rich'], ['', special, `${special}_x0041_`]]);
    } finally {
      await streamed.close();
    }
  }, 30_000);

  it('bounds long strings by payload before reaching the entry limit, with deterministic output', async () => {
    const write = async () => {
      const sink = toBuffer();
      const wb = await createWriteOnlyWorkbook(sink, { compressionLevel: 0, mtime: new Date('2020-01-01T00:00:00Z') });
      const ws = await wb.addWorksheet('Long strings');
      for (let i = 0; i < 100; i++) await ws.appendRow([`${i}`.padEnd(32_000, 'x')]);
      await ws.close();
      await wb.finalize();
      return sink.result();
    };
    const bytes = await write();
    expect(bytes.equals(await write())).toBe(true);
    const zip = await openZip(fromBuffer(bytes));
    try {
      const xml = new TextDecoder().decode(zip.read('xl/worksheets/sheet1.xml'));
      expect(xml).toContain('<c r="A1" t="s">');
      expect(xml).toContain('<c r="A100" t="inlineStr">');
      const sst = new TextDecoder().decode(zip.read('xl/sharedStrings.xml'));
      expect(sst.match(/<si>/g)?.length).toBeLessThan(100);
    } finally {
      zip.close();
    }
  });
});
