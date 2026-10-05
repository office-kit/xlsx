// Serialise a computed PivotTable into its three OOXML parts:
// xl/pivotCache/pivotCacheDefinitionN.xml, xl/pivotCache/pivotCacheRecordsN.xml
// and xl/pivotTables/pivotTableN.xml (ECMA-376 Part 1 §18.10).

import { escapeXmlAttr } from '../utils/escape.js';
import { REL_NS, SHEET_MAIN_NS } from '../xml/namespaces.js';
import {
  type PivotAxisEntry,
  type PivotCacheField,
  type PivotCacheValue,
  type PivotStructure,
  type PivotTable,
} from './pivot-table.js';

// Version stamps Excel compares against its own: a cache whose
// refreshedVersion is older than the running Excel is upgraded silently, and
// minRefreshableVersion 3 keeps the parts readable by Excel 2007.
const CREATED_VERSION = 6;
const MIN_REFRESHABLE_VERSION = 3;
const PIVOT_STYLE = 'PivotStyleLight16';

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const encoder = new TextEncoder();

/**
 * `<pivotCacheDefinition>`. `refreshOnLoad` makes Excel rebuild the cache and
 * report from the source when the file opens, so a report whose source cells
 * changed after the last refresh still opens correct.
 */
export function pivotCacheDefinitionToBytes(pt: PivotTable, c: PivotStructure, recordsRId: string): Uint8Array {
  const parts: string[] = [
    XML_DECL,
    `<pivotCacheDefinition xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}" r:id="${recordsRId}" refreshOnLoad="1"`,
    ` createdVersion="${CREATED_VERSION}" refreshedVersion="${CREATED_VERSION}" minRefreshableVersion="${MIN_REFRESHABLE_VERSION}" recordCount="${c.recordCount}">`,
    `<cacheSource type="worksheet"><worksheetSource ref="${escapeXmlAttr(pt.source.ref)}" sheet="${escapeXmlAttr(pt.source.sheet)}"/></cacheSource>`,
    `<cacheFields count="${c.fields.length}">`,
  ];
  for (const f of c.fields) parts.push(cacheFieldXml(f));
  parts.push('</cacheFields></pivotCacheDefinition>');
  return encoder.encode(parts.join(''));
}

function cacheFieldXml(f: PivotCacheField): string {
  const s = f.stats;
  // <sharedItems> attribute defaults describe a text-only field; anything
  // else has to be spelled out or Excel flags the cache as corrupt.
  let attrs = '';
  const nonText = s.hasNumber || s.hasBoolean || s.hasError;
  if (!s.hasString && !s.hasBlank && nonText) attrs += ' containsSemiMixedTypes="0"';
  if (!s.hasString) attrs += ' containsString="0"';
  if (s.hasBlank) attrs += ' containsBlank="1"';
  const kinds = Number(s.hasString) + Number(s.hasNumber) + Number(s.hasBoolean) + Number(s.hasError);
  if (kinds > 1) attrs += ' containsMixedTypes="1"';
  if (s.hasNumber) {
    attrs += ' containsNumber="1"';
    if (s.allIntegers) attrs += ' containsInteger="1"';
    attrs += ` minValue="${s.min}" maxValue="${s.max}"`;
  }
  if (f.items === undefined) {
    return `<cacheField name="${escapeXmlAttr(f.name)}" numFmtId="0"><sharedItems${attrs}/></cacheField>`;
  }
  const items = f.items.map(valueXml).join('');
  return `<cacheField name="${escapeXmlAttr(f.name)}" numFmtId="0"><sharedItems${attrs} count="${f.items.length}">${items}</sharedItems></cacheField>`;
}

function valueXml(v: PivotCacheValue): string {
  switch (v.t) {
    case 's':
      return `<s v="${escapeXmlAttr(v.v)}"/>`;
    case 'n':
      return `<n v="${v.v}"/>`;
    case 'b':
      return `<b v="${v.v ? 1 : 0}"/>`;
    case 'e':
      return `<e v="${escapeXmlAttr(v.v)}"/>`;
    case 'm':
      return '<m/>';
  }
}

/** `<pivotCacheRecords>`: one `<r>` per source row, items by index, other fields inline. */
export function pivotCacheRecordsToBytes(c: PivotStructure): Uint8Array {
  const parts: string[] = [XML_DECL, `<pivotCacheRecords xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}" count="${c.recordCount}">`];
  for (let i = 0; i < c.recordCount; i++) {
    let row = '<r>';
    for (const f of c.fields) {
      const x = f.itemIndex?.[i];
      if (x !== undefined) row += `<x v="${x}"/>`;
      else {
        const v = f.records[i];
        row += v === undefined ? '<m/>' : valueXml(v);
      }
    }
    parts.push(`${row}</r>`);
  }
  parts.push('</pivotCacheRecords>');
  return encoder.encode(parts.join(''));
}

/** `<pivotTableDefinition>` for `pt`, bound to the workbook-level `cacheId`. */
export function pivotTableToBytes(pt: PivotTable, c: PivotStructure, cacheId: number): Uint8Array {
  const compact = pt.layout === 'compact';
  let attrs = ` name="${escapeXmlAttr(pt.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0"`;
  attrs += ' applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1"';
  attrs += ` dataCaption="Values" updatedVersion="${CREATED_VERSION}" minRefreshableVersion="${MIN_REFRESHABLE_VERSION}"`;
  attrs += ' useAutoFormatting="1" itemPrintTitles="1"';
  attrs += ` createdVersion="${CREATED_VERSION}" indent="0"`;
  if (!compact) attrs += ' compact="0" compactData="0"';
  if (pt.layout !== 'tabular') attrs += ' outline="1" outlineData="1"';
  if (!pt.rowGrandTotals) attrs += ' rowGrandTotals="0"';
  if (!pt.columnGrandTotals) attrs += ' colGrandTotals="0"';
  attrs += ' multipleFieldFilters="0"';

  const loc = c.location;
  let locAttrs = ` ref="${loc.ref}" firstHeaderRow="${loc.firstHeaderRow}" firstDataRow="${loc.firstDataRow}" firstDataCol="${loc.firstDataCol}"`;
  if (pt.filters.length > 0) locAttrs += ` rowPageCount="${pt.filters.length}" colPageCount="1"`;

  const parts: string[] = [
    XML_DECL,
    `<pivotTableDefinition xmlns="${SHEET_MAIN_NS}"${attrs}>`,
    `<location${locAttrs}/>`,
    `<pivotFields count="${c.fields.length}">`,
  ];
  const rowSet = new Set(pt.rows);
  const colSet = new Set(pt.columns);
  const pageSet = new Set(pt.filters.map((f) => f.field));
  const dataSet = new Set(pt.values.map((v) => v.field));
  c.fields.forEach((f, idx) => {
    const axis = rowSet.has(idx) ? 'axisRow' : colSet.has(idx) ? 'axisCol' : pageSet.has(idx) ? 'axisPage' : undefined;
    let fa = '';
    if (axis !== undefined) fa += ` axis="${axis}"`;
    if (dataSet.has(idx)) fa += ' dataField="1"';
    if (!compact) fa += ' compact="0"';
    if (pt.layout === 'tabular') fa += ' outline="0"';
    fa += ' showAll="0"';
    if (axis === undefined || f.items === undefined) {
      parts.push(`<pivotField${fa}/>`);
      return;
    }
    const subtotal = pt.subtotals !== 'off' || axis === 'axisPage';
    if (!subtotal) fa += ' defaultSubtotal="0"';
    if (pt.subtotals === 'bottom' && axis === 'axisRow') fa += ' subtotalTop="0"';
    const items = f.items.map((_, i) => `<item x="${i}"/>`).join('');
    const count = f.items.length + (subtotal ? 1 : 0);
    parts.push(`<pivotField${fa}><items count="${count}">${items}${subtotal ? '<item t="default"/>' : ''}</items></pivotField>`);
  });
  parts.push('</pivotFields>');

  if (pt.rows.length > 0) {
    parts.push(`<rowFields count="${pt.rows.length}">${pt.rows.map((f) => `<field x="${f}"/>`).join('')}</rowFields>`);
  }
  if (hasReport(pt)) parts.push(axisItemsXml('rowItems', c.rowEntries));
  if (c.colFields.length > 0) {
    parts.push(`<colFields count="${c.colFields.length}">${c.colFields.map((f) => `<field x="${f}"/>`).join('')}</colFields>`);
  }
  if (hasReport(pt)) parts.push(axisItemsXml('colItems', c.colEntries));
  if (pt.filters.length > 0) {
    const pages = pt.filters.map((f) => {
      const field = c.fields[f.field];
      const idx = f.selected === undefined ? -1 : (field?.items?.findIndex((it) => sameItem(it, f.selected ?? null)) ?? -1);
      return `<pageField fld="${f.field}"${idx >= 0 ? ` item="${idx}"` : ''} hier="-1"/>`;
    });
    parts.push(`<pageFields count="${pages.length}">${pages.join('')}</pageFields>`);
  }
  if (pt.values.length > 0) {
    const dfs = pt.values.map((v, i) => {
      const sub = v.aggregate === 'sum' ? '' : ` subtotal="${v.aggregate}"`;
      return `<dataField name="${escapeXmlAttr(c.valueNames[i] ?? '')}" fld="${v.field}"${sub} baseField="0" baseItem="0"/>`;
    });
    parts.push(`<dataFields count="${dfs.length}">${dfs.join('')}</dataFields>`);
  }
  parts.push(
    `<pivotTableStyleInfo name="${PIVOT_STYLE}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>`,
    '</pivotTableDefinition>',
  );
  return encoder.encode(parts.join(''));
}

const hasReport = (pt: PivotTable): boolean => pt.rows.length + pt.columns.length + pt.values.length > 0;

function sameItem(item: PivotCacheValue, v: string | number | boolean | null): boolean {
  if (v === null) return item.t === 'm';
  if (typeof v === 'string') return item.t === 's' && item.v.toLowerCase() === v.toLowerCase();
  return (item.t === 'n' || item.t === 'b') && item.v === v;
}

function axisItemsXml(tag: 'rowItems' | 'colItems', entries: PivotAxisEntry[]): string {
  const out: string[] = [`<${tag} count="${entries.length}">`];
  for (const e of entries) {
    let attrs = '';
    if (e.type !== 'data') attrs += ` t="${e.type}"`;
    if (e.r > 0) attrs += ` r="${e.r}"`;
    if (e.i > 0) attrs += ` i="${e.i}"`;
    const xs = e.x.map((v) => (v === 0 ? '<x/>' : `<x v="${v}"/>`)).join('');
    out.push(xs.length === 0 ? `<i${attrs}/>` : `<i${attrs}>${xs}</i>`);
  }
  out.push(`</${tag}>`);
  return out.join('');
}
