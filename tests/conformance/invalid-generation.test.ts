import fc from 'fast-check';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { packageFor, SML } from './corpus.js';
import { required } from './required.js';
import { stripIgnorableMarkup } from './mc-strip.js';
import { validateXlsx } from './validate.js';
import { checkZipEnvelope } from './zip-envelope.js';

describe('independently authored invalid input and oracle calibration', () => {
  it('rejects generated out-of-grid references, with a valid boundary control', async () => {
    const at = (ref: string, row: number) => {
      const parts = unzipSync(packageFor({ id: 'boundary', clause: '18.3.1.4' }));
      parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${SML}"><dimension ref="${ref}"/><sheetData><row r="${row}"><c r="${ref}"><v>1</v></c></row></sheetData></worksheet>`);
      return zipSync(parts);
    };
    expect((await validateXlsx(at('XFD1048576', 1048576))).status).toBe('valid');
    const numRuns = Number(process.env['QA_FUZZ_RUNS'] ?? 30);
    const seed = Number(process.env['QA_FUZZ_SEED'] ?? 376263);
    if (!Number.isInteger(numRuns) || numRuns < 1 || !Number.isInteger(seed)) throw new Error('Invalid fuzz configuration');
    await fc.assert(fc.asyncProperty(fc.integer({ min: 1048577, max: 2_000_000 }), async row => {
      const result = await validateXlsx(at(`A${row}`, row));
      expect(result.status, JSON.stringify(result)).toBe('invalid');
      expect(result.issues.some(i => i.tier === 'semantic')).toBe(true);
    }), { numRuns, seed });
    expect((await validateXlsx(at('XFE1', 1))).status).toBe('invalid');
  }, 600_000);
  it.each(['../../../xl/workbook.xml', 'xl/%ZZ.xml', 'https://example.org/xl/workbook.xml', '//xl/workbook.xml', 'xl\\workbook.xml'])('rejects internal target %s', async target => {
    const parts = unzipSync(packageFor({ id: 'target', clause: 'OPC 9.2' }));
    parts['_rels/.rels'] = strToU8(strFromU8(required(parts['_rels/.rels'])).replace('Target="xl/workbook.xml"', `Target="${target}"`));
    expect((await validateXlsx(zipSync(parts))).issues.some(i => i.tier === 'opc')).toBe(true);
  });
  it('detects duplicate names before inflation can collapse them', () => {
    const bytes = zipSync({ 'a.xml': strToU8('<a/>'), 'b.xml': strToU8('<b/>') });
    const mutated = bytes.slice();
    for (let i = 0; i < mutated.length - 4; i++) {
      if (strFromU8(mutated.subarray(i, i + 5)) === 'b.xml') mutated[i] = 97;
    }
    expect(() => checkZipEnvelope(mutated)).toThrow(/Duplicate ZIP/);
    expect(checkZipEnvelope(bytes)).toBe('zip32');
  });
  it('rejects advertised oversized inflation before decompression', () => {
    const bytes = zipSync({ 'a.xml': strToU8('<a/>') });
    const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    for (let i = 0; i < bytes.length - 46; i++) if (v.getUint32(i, true) === 0x02014b50) v.setUint32(i + 24, 17 * 1024 * 1024, true);
    expect(() => checkZipEnvelope(bytes)).toThrow(/inflation budget/);
  });
  it('never certifies ZIP64 outside the bounded profile', async () => {
    const bytes = packageFor({ id: 'zip64', clause: 'OPC' });
    const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    v.setUint16(bytes.length - 12, 0xffff, true);
    expect(await validateXlsx(bytes)).toMatchObject({ ok: false, status: 'incomplete' });
  });
  const wrap = (body: string) => `<root xmlns:s="${SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${body}</root>`;
  it.each([
    '<mc:Choice Requires="s"><kept/></mc:Choice><mc:Fallback/><mc:Fallback/>',
    '<mc:Fallback/><mc:Choice Requires="s"/>',
    '<mc:Choice Requires="s"><kept/></mc:Choice><mc:Choice/>',
    '<mc:Choice Requires="s"><kept/></mc:Choice><mc:Choice Requires="missing"/>',
    '<mc:Fallback/>',
    '<mc:Choice Requires="s"/>unexpected',
  ])('validates the entire AlternateContent sequence: %s', body => {
    expect(() => stripIgnorableMarkup(wrap(`<mc:AlternateContent>${body}</mc:AlternateContent>`))).toThrow();
  });
  it('validates xml:space before normalization', () => {
    expect(() => stripIgnorableMarkup(wrap('<s:t xml:space="invalid">x</s:t>'))).toThrow(/xml:space/);
    expect(stripIgnorableMarkup(wrap('<s:t xml:space="preserve"> x </s:t>'))).toContain(' x ');
  });
});
