import { readFileSync } from 'node:fs';
import { strToU8, zipSync } from 'fflate';
import { required } from './required.js';
import type { CellValue } from '../../src/cell/cell.js';
import type { DefinedName } from '../../src/workbook/defined-names.js';
import type { Font } from '../../src/styles/fonts.js';
export interface CorpusCase {
  id: string; clause: string; fontXml?: string; expectedFont?: Partial<Font>; expectedMissingFont?: Array<keyof Font>;
  borderXml?: string; expectedBorder?: string; cellXml?: string;
  expectedValue?: CellValue | { date: string }; omitRowIndex?: boolean;
  rowsXml?: string; sharedStringsXml?: string; sheetName?: string; workbookBeforeSheets?: string; workbookAfterSheets?: string;
  expectedStreamValues?: CellValue[];
  expectedValues?: Array<CellValue | { date: string }>; expectedNames?: DefinedName[]; expectedDate1904?: boolean;
  expectedDate?: string; numFmtId?: number;
  officeValue?: string | number | boolean; officeValues?: Array<string | number | boolean>;
}
export const corpus = (JSON.parse(readFileSync(new URL('./corpus/manifest.json', import.meta.url), 'utf8')) as { cases: CorpusCase[] }).cases;
export const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export function canonicalValue(value: unknown): unknown {
  return value instanceof Date ? { date: value.toISOString() } : value;
}
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
/** Deliberately independent of production serialization and metadata enums. */
export function packageFor(c: CorpusCase): Uint8Array {
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${CT}.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="${CT}.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${CT}.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${SML}" xmlns:r="${R}">${c.workbookBeforeSheets ?? ''}<sheets><sheet name="${c.sheetName ?? 'Audit'}" sheetId="1" r:id="sheet"/></sheets>${c.workbookAfterSheets ?? ''}</workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="sheet" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `<styleSheet xmlns="${SML}"><fonts count="1"><font>${c.fontXml ?? ''}</font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border>${c.borderXml ?? ''}</border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${c.numFmtId === undefined ? 1 : 2}">${c.numFmtId === undefined ? '' : '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'}<xf numFmtId="${c.numFmtId ?? 0}" fontId="0" fillId="0" borderId="0" xfId="0"${c.numFmtId === undefined ? '' : ' applyNumberFormat="1"'}/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${SML}"><dimension ref="${c.rowsXml ? 'A1:C2' : 'A1'}"/><sheetData>${c.rowsXml ?? `<row${c.omitRowIndex ? '' : ' r="1"'}>${c.cellXml ?? '<c r="A1" t="inlineStr" s="0"><is><t>audit</t></is></c>'}</row>`}</sheetData></worksheet>`,
  };
  if (c.sharedStringsXml) {
    parts['xl/sharedStrings.xml'] = `<sst xmlns="${SML}">${c.sharedStringsXml}</sst>`;
    parts['[Content_Types].xml'] = required(parts['[Content_Types].xml']).replace('</Types>', `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT}.sharedStrings+xml"/></Types>`);
    parts['xl/_rels/workbook.xml.rels'] = required(parts['xl/_rels/workbook.xml.rels']).replace('</Relationships>', `<Relationship Id="strings" Type="${R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
  }
  return zipSync(Object.fromEntries(Object.entries(parts).map(([p, xml]) => [p, strToU8(xml)])));
}
