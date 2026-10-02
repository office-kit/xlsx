import { readFileSync } from 'node:fs';
import { strToU8, zipSync } from 'fflate';
import type { Font } from '../../src/styles/fonts.js';
export interface CorpusCase {
  id: string; clause: string; fontXml?: string; expectedFont?: Partial<Font>; expectedMissingFont?: Array<keyof Font>;
  borderXml?: string; expectedBorder?: string; cellXml?: string;
  expectedValue?: string | boolean; omitRowIndex?: boolean;
}
export const corpus = (JSON.parse(readFileSync(new URL('./corpus/manifest.json', import.meta.url), 'utf8')) as { cases: CorpusCase[] }).cases;
export const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
/** Deliberately independent of production serialization and metadata enums. */
export function packageFor(c: CorpusCase): Uint8Array {
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${CT}.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="${CT}.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${CT}.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${SML}" xmlns:r="${R}"><sheets><sheet name="Audit" sheetId="1" r:id="sheet"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="sheet" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `<styleSheet xmlns="${SML}"><fonts count="1"><font>${c.fontXml ?? ''}</font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border>${c.borderXml ?? ''}</border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${SML}"><dimension ref="A1"/><sheetData><row${c.omitRowIndex ? '' : ' r="1"'}>${c.cellXml ?? '<c r="A1" t="inlineStr" s="0"><is><t>audit</t></is></c>'}</row></sheetData></worksheet>`,
  };
  return zipSync(Object.fromEntries(Object.entries(parts).map(([p, xml]) => [p, strToU8(xml)])));
}
