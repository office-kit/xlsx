import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { workbookToBytes } from '../../src/io/save.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { setCell } from '../../src/worksheet/worksheet.js';
import { required } from './required.js';
import { stripIgnorableMarkup } from './mc-strip.js';
import { validateXlsx } from './validate.js';

const SHEET = 'xl/worksheets/sheet1.xml';
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
type Parts = Record<string, Uint8Array>;
let original: Uint8Array;
beforeAll(async () => {
  const wb = createWorkbook();
  setCell(addWorksheet(wb, 'Audit'), 1, 1, 'hello');
  original = await workbookToBytes(wb);
});
function edit(parts: Parts, path: string, transform: (xml: string) => string): void {
  parts[path] = strToU8(transform(strFromU8(required(parts[path]))));
}
function prefixed(xml: string): string {
  return xml.replace('<worksheet ', `<worksheet xmlns:s="${SML}" `)
    .replace(/<(\/?)(row|c|v)\b/g, '<$1s:$2');
}

describe('validator calibration: single-fault package mutations', () => {
  const faults: Array<[string, (parts: Parts) => void, string]> = [
    ['missing package officeDocument relationship', p => edit(p, '_rels/.rels', s => s.replace(/<Relationship[^>]*officeDocument[^>]*\/>/, '')), 'opc'],
    ['wrong drawing target content type', p => {
      edit(p, SHEET, s => s.replace('</worksheet>', '<drawing r:id="wrong"/></worksheet>'));
      p['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="wrong" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../styles.xml"/></Relationships>');
    }, 'opc'],
    ['font index beyond actual fonts', p => edit(p, 'xl/styles.xml', s => s.replace(/fontId="0"/g, 'fontId="998"')), 'semantic'],
    ['prefixed wrong parent row', p => edit(p, SHEET, s => prefixed(s.replace('r="A1"', 'r="A99"'))), 'semantic'],
    ['lying style count', p => {
      edit(p, 'xl/styles.xml', s => s.replace(/<cellXfs count="\d+"/, '<cellXfs count="999"'));
      edit(p, SHEET, s => s.replace('<c r="A1"', '<c s="998" r="A1"'));
    }, 'semantic'],
    ['style index without count', p => {
      edit(p, 'xl/styles.xml', s => s.replace(/<cellXfs count="\d+"/, '<cellXfs'));
      edit(p, SHEET, s => s.replace('<c r="A1"', '<c s="998" r="A1"'));
    }, 'semantic'],
    ['shared string index without uniqueCount', p => {
      edit(p, 'xl/sharedStrings.xml', s => s.replace(/ uniqueCount="\d+"/, ''));
      edit(p, SHEET, s => s.replace('<v>0</v>', '<v>998</v>'));
    }, 'semantic'],
    ['dimension excludes cell', p => edit(p, SHEET, s => s.replace(/<dimension ref="[^"]+"/, '<dimension ref="Z99"')), 'semantic'],
    ['missing drawing relationship', p => edit(p, SHEET, s => s.replace('</worksheet>', '<drawing r:id="missing"/></worksheet>')), 'opc'],
    ['wrong drawing relationship type', p => {
      edit(p, SHEET, s => s.replace('</worksheet>', '<drawing r:id="wrong"/></worksheet>'));
      p['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="wrong" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../styles.xml"/></Relationships>');
    }, 'opc'],
    ['malformed unknown XML', p => {
      p['xl/unknown.xml'] = strToU8('<not-closed>');
      edit(p, '[Content_Types].xml', s => s.replace('</Types>', '<Override PartName="/xl/unknown.xml" ContentType="application/x-audit+xml"/></Types>'));
    }, 'xml'],
    ['missing content type default', p => edit(p, '[Content_Types].xml', s => s.replace(/<Default[^>]*Extension="rels"[^>]*\/>/, '')), 'opc'],
    ['duplicate relationship id', p => edit(p, '_rels/.rels', s => s.replace('</Relationships>', '<Relationship Id="rId1" Type="urn:audit" Target="xl/workbook.xml"/></Relationships>')), 'opc'],
    ['malformed XML before MC normalization', p => edit(p, SHEET, s => s.replace('</sheetData>', '')), 'xml'],
  ];
  it.each(faults)('rejects %s', async (_name, mutate, tier) => {
    const parts = unzipSync(original);
    mutate(parts);
    const result = await validateXlsx(zipSync(parts));
    expect(result.ok, JSON.stringify(result)).toBe(false);
    expect(result.issues.some(i => i.tier === tier), JSON.stringify(result)).toBe(true);
  });
  it('does not pass when XSD validation is skipped', async () => {
    expect(await validateXlsx(original, { skipXsd: true })).toMatchObject({ ok: false, status: 'incomplete' });
  });
  it('reports a missing XSD executable as inconclusive', async () => {
    vi.stubEnv('PATH', '/nonexistent-qa-tool-directory');
    try {
      expect(await validateXlsx(original)).toMatchObject({ ok: false, status: 'inconclusive' });
    } finally { vi.unstubAllEnvs(); }
  });
  it('accepts equivalent prefixed XML and single-quoted attributes', async () => {
    const parts = unzipSync(original);
    edit(parts, SHEET, s => prefixed(s).replace(/"/g, "'"));
    expect((await validateXlsx(zipSync(parts))).issues).toEqual([]);
  });
  it('reports unknown well-formed XML as incomplete instead of fully validated', async () => {
    const parts = unzipSync(original);
    parts['xl/unknown.xml'] = strToU8('<extension/>');
    edit(parts, '[Content_Types].xml', s => s.replace('</Types>', '<Override PartName="/xl/unknown.xml" ContentType="application/x-audit+xml"/></Types>'));
    const result = await validateXlsx(zipSync(parts));
    expect(result).toMatchObject({ ok: false, status: 'incomplete', issues: [], skipped: ['xl/unknown.xml'] });
  });
});

describe('MC consumer profile', () => {
  const wrap = (body: string) => `<root xmlns:s="${SML}" xmlns:u="urn:unsupported" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="u">${body}</root>`;
  it('requires every namespace to be understood', () => {
    const result = stripIgnorableMarkup(wrap('<mc:AlternateContent><mc:Choice Requires="s u"><choice/></mc:Choice><mc:Fallback><fallback/></mc:Fallback></mc:AlternateContent>'));
    expect(result).toContain('<fallback');
    expect(result).not.toContain('<choice');
  });
  it('understands supported namespaces even when marked ignorable', () => {
    expect(stripIgnorableMarkup(wrap('<s:t mc:Ignorable="s">kept</s:t>'))).toContain('kept');
  });
  it('carries namespace bindings from a selected Choice', () => {
    const result = stripIgnorableMarkup(wrap(`<mc:AlternateContent><mc:Choice xmlns:z="${SML}" Requires="z"><z:t>kept</z:t></mc:Choice><mc:Fallback/></mc:AlternateContent>`));
    expect(result).toContain('<z:t');
    expect(result).toContain(`xmlns:z="${SML}"`);
  });
  it('processes children of an ignored ProcessContent wrapper', () => {
    expect(stripIgnorableMarkup(wrap('<u:wrapper mc:ProcessContent="u:wrapper"><s:t>kept</s:t></u:wrapper>'))).toContain('kept');
  });
  it('rejects unknown MustUnderstand namespaces', () => {
    expect(() => stripIgnorableMarkup(wrap('<s:t mc:MustUnderstand="u"/>'))).toThrow(/MustUnderstand/);
  });
});
