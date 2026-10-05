// Read PivotTables back out of the passthrough parts the loader captured.
//
// The loader carries every pivot part verbatim. `liftPivotTables` then turns
// the pivots this library can express without losing anything into
// `ws.pivotTables` definitions (dropping their passthrough parts, which the
// writer regenerates), so a pivot written by `saveWorkbook` comes back
// editable. A pivot using anything the model lacks (hidden items, grouping,
// calculated fields, custom styles or captions, number formats, "show values
// as", slicers, …) stays passthrough and round-trips byte for byte;
// `listPassthroughPivotTables` still describes it for read-only display.

import type { Workbook } from '../workbook/workbook.js';
import { getSheet } from '../workbook/workbook.js';
import { relsFromBytes } from '../packaging/relationships.js';
import { resolveRelTarget } from '../packaging/part-name.js';
import { tupleToCoordinate } from '../utils/coordinate.js';
import { MARKUP_COMPAT_NS, REL_NS, SHEET_MAIN_NS, X16_NS } from '../xml/namespaces.js';
import { parseXml } from '../xml/parser.js';
import { findChild, findChildren, type XmlNode } from '../xml/tree.js';
import { parseRange } from './cell-range.js';
import {
  PIVOT_VALUES_FIELD,
  type PivotAggregate,
  type PivotItemValue,
  type PivotLayout,
  type PivotSubtotals,
  type PivotTable,
  sourceFieldNames,
} from './pivot-table.js';
import type { Worksheet } from './worksheet.js';

/** What a pivot shows, by field name — enough to draw a read-only field list. */
export interface PivotTableSummary {
  name: string;
  /** Table body as `<location ref>` records it (report filters excluded). */
  ref: string;
  /** Every cache field name, in cache order. */
  fields: string[];
  rows: string[];
  columns: string[];
  filters: string[];
  /** Value field captions, e.g. "Sum of Sales". */
  values: string[];
}

const PIVOT_TABLE_REL = `${REL_NS}/pivotTable`;
const PIVOT_CACHE_DEFINITION_REL = `${REL_NS}/pivotCacheDefinition`;
const PIVOT_CACHE_RECORDS_REL = `${REL_NS}/pivotCacheRecords`;
// Worksheet rels targets are relative to the sheet part; every producer we
// know of keeps worksheets here. A pivot whose target doesn't resolve to a
// captured part just stays passthrough.
const SHEET_PART_BASE = 'xl/worksheets/sheet.xml';

const q = (local: string): string => `{${SHEET_MAIN_NS}}${local}`;

const AGGREGATES: ReadonlySet<string> = new Set<PivotAggregate>(['sum', 'count', 'average', 'max', 'min']);

// Attributes the writer regenerates or the model carries. Anything else on
// these elements changes what Excel shows, so a pivot using it is not lifted.
const TABLE_ATTRS = new Set([
  'name',
  'cacheId',
  'applyNumberFormats',
  'applyBorderFormats',
  'applyFontFormats',
  'applyPatternFormats',
  'applyAlignmentFormats',
  'applyWidthHeightFormats',
  'dataCaption',
  'updatedVersion',
  'minRefreshableVersion',
  'useAutoFormatting',
  'itemPrintTitles',
  'createdVersion',
  'indent',
  'outline',
  'outlineData',
  'compact',
  'compactData',
  'rowGrandTotals',
  'colGrandTotals',
  'multipleFieldFilters',
  `{${X16_NS}}uid`,
  `{${MARKUP_COMPAT_NS}}Ignorable`,
]);
const TABLE_CHILDREN = new Set(
  ['location', 'pivotFields', 'rowFields', 'rowItems', 'colFields', 'colItems', 'pageFields', 'dataFields', 'pivotTableStyleInfo', 'extLst'].map(q),
);
const FIELD_ATTRS = new Set(['axis', 'dataField', 'showAll', 'compact', 'outline', 'defaultSubtotal', 'subtotalTop']);
const DATA_FIELD_ATTRS = new Set(['name', 'fld', 'subtotal', 'baseField', 'baseItem']);
// Excel 2016+ tags every pivot it saves with this empty marker ext.
const PIVOT_DEFAULT_LAYOUT_EXT = '{747A6164-185A-40DC-8AA5-F01512510D54}';
// ... and every cache with this one, empty unless slicers attach.
const X14_CACHE_EXT = '{725AE2AE-9491-48be-B2B4-4EB974FC3084}';
const DEFAULT_STYLE = 'PivotStyleLight16';

interface LocatedPivot {
  ws: Worksheet;
  relId: string;
  tablePath: string;
  cachePath: string | undefined;
  table: XmlNode;
  cache: XmlNode | undefined;
}

function locatePivots(wb: Workbook, ws: Worksheet): LocatedPivot[] {
  const out: LocatedPivot[] = [];
  const parts = wb.passthrough;
  if (parts === undefined) return out;
  for (const rel of ws.relsExtras ?? []) {
    if (rel.type !== PIVOT_TABLE_REL) continue;
    const tablePath = resolveRelTarget(SHEET_PART_BASE, rel.target);
    const tableBytes = parts.get(tablePath);
    if (tableBytes === undefined) continue;
    const relsBytes = parts.get(relsPathFor(tablePath));
    const cacheRel = relsBytes ? relsFromBytes(relsBytes).rels.find((r) => r.type === PIVOT_CACHE_DEFINITION_REL) : undefined;
    const cachePath = cacheRel ? resolveRelTarget(tablePath, cacheRel.target) : undefined;
    const cacheBytes = cachePath === undefined ? undefined : parts.get(cachePath);
    out.push({
      ws,
      relId: rel.id,
      tablePath,
      cachePath,
      table: parseXml(tableBytes),
      cache: cacheBytes === undefined ? undefined : parseXml(cacheBytes),
    });
  }
  return out;
}

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
}

const cacheFieldNames = (cache: XmlNode | undefined): string[] =>
  cache === undefined
    ? []
    : findChildren(findChild(cache, q('cacheFields')) ?? cache, q('cacheField')).map((f) => f.attrs['name'] ?? '');

const fieldIndices = (table: XmlNode, tag: 'rowFields' | 'colFields'): number[] => {
  const node = findChild(table, q(tag));
  return node === undefined ? [] : findChildren(node, q('field')).map((f) => Number(f.attrs['x'] ?? '0'));
};

function summarize(p: LocatedPivot): PivotTableSummary {
  const fields = cacheFieldNames(p.cache);
  const name = (i: number): string => (i === PIVOT_VALUES_FIELD ? 'Σ Values' : (fields[i] ?? `#${i}`));
  const pageNode = findChild(p.table, q('pageFields'));
  const dataNode = findChild(p.table, q('dataFields'));
  return {
    name: p.table.attrs['name'] ?? '',
    ref: findChild(p.table, q('location'))?.attrs['ref'] ?? '',
    fields,
    rows: fieldIndices(p.table, 'rowFields').map(name),
    columns: fieldIndices(p.table, 'colFields').map(name),
    filters: pageNode ? findChildren(pageNode, q('pageField')).map((f) => name(Number(f.attrs['fld'] ?? '0'))) : [],
    values: dataNode ? findChildren(dataNode, q('dataField')).map((f) => f.attrs['name'] ?? name(Number(f.attrs['fld'] ?? '0'))) : [],
  };
}

/** The pivots on `ws` that ride passthrough because the model can't express them. */
export function listPassthroughPivotTables(wb: Workbook, ws: Worksheet): PivotTableSummary[] {
  return locatePivots(wb, ws).map(summarize);
}

// `xmlns:*` entries are prefix declarations the parser keeps for an
// `mc:Ignorable`, not attributes of the element.
const onlyKnownAttrs = (node: XmlNode, allowed: ReadonlySet<string>): boolean =>
  Object.keys(node.attrs).every((k) => allowed.has(k) || k.startsWith('xmlns:'));

const isTrue = (v: string | undefined, dflt: boolean): boolean => (v === undefined ? dflt : v === '1' || v === 'true');

const onlyMarkerExt = (node: XmlNode, uri: string): boolean => {
  const ext = findChild(node, q('extLst'));
  if (ext === undefined) return true;
  return ext.children.every((e) => e.attrs['uri'] === uri && e.children.every((c) => c.children.length === 0 && Object.keys(c.attrs).length === 0));
};

function cacheItemValue(node: XmlNode | undefined): PivotItemValue | undefined {
  if (node === undefined) return undefined;
  const v = node.attrs['v'];
  switch (node.name) {
    case q('s'):
      return v ?? '';
    case q('n'):
      return Number(v);
    case q('b'):
      return v === '1' || v === 'true';
    case q('m'):
      return null;
    default:
      return undefined;
  }
}

/** The definition `p` stands for, or undefined when the model can't hold it losslessly. */
function toDefinition(wb: Workbook, p: LocatedPivot): PivotTable | undefined {
  const { table, cache } = p;
  if (cache === undefined) return undefined;
  if (!onlyKnownAttrs(table, TABLE_ATTRS) || !table.children.every((c) => TABLE_CHILDREN.has(c.name))) return undefined;
  if ((table.attrs['dataCaption'] ?? 'Values') !== 'Values') return undefined;
  if (!onlyMarkerExt(table, PIVOT_DEFAULT_LAYOUT_EXT) || !onlyMarkerExt(cache, X14_CACHE_EXT)) return undefined;
  const style = findChild(table, q('pivotTableStyleInfo'));
  if (style !== undefined && style.attrs['name'] !== DEFAULT_STYLE) return undefined;

  const source = findChild(findChild(cache, q('cacheSource')) ?? cache, q('worksheetSource'));
  const sheet = source?.attrs['sheet'];
  const ref = source?.attrs['ref'];
  if (sheet === undefined || ref === undefined || getSheet(wb, sheet) === undefined) return undefined;
  if (findChild(cache, q('cacheSource'))?.attrs['type'] !== 'worksheet') return undefined;
  if (!cache.children.every((c) => c.name === q('cacheSource') || c.name === q('cacheFields') || c.name === q('extLst'))) return undefined;
  const cacheFields = findChildren(findChild(cache, q('cacheFields')) ?? cache, q('cacheField'));
  if (cacheFields.some((f) => f.attrs['formula'] !== undefined || findChild(f, q('fieldGroup')) !== undefined)) return undefined;

  // Field indices are source column offsets in the model, so the cache has to
  // line up with the source header as it stands now.
  const names = sourceFieldNames(wb, { sheet, ref });
  if (names === undefined) return undefined;
  const cacheNames = cacheFields.map((f) => (f.attrs['name'] ?? '').toLowerCase());
  if (names.length !== cacheNames.length || names.some((n, i) => n.toLowerCase() !== cacheNames[i])) return undefined;

  const pivotFields = findChildren(findChild(table, q('pivotFields')) ?? table, q('pivotField'));
  if (pivotFields.length !== names.length) return undefined;
  for (const f of pivotFields) {
    if (!onlyKnownAttrs(f, FIELD_ATTRS)) return undefined;
    if (!f.children.every((c) => c.name === q('items'))) return undefined;
    const items = findChild(f, q('items'));
    if (items && !items.children.every((it) => Object.keys(it.attrs).every((k) => k === 'x' || (k === 't' && it.attrs['t'] === 'default')))) {
      return undefined;
    }
  }

  const rows = fieldIndices(table, 'rowFields');
  const colFields = fieldIndices(table, 'colFields');
  if (rows.includes(PIVOT_VALUES_FIELD)) return undefined;
  const sigma = colFields.indexOf(PIVOT_VALUES_FIELD);
  if (sigma >= 0 && sigma !== colFields.length - 1) return undefined;
  const columns = colFields.filter((f) => f !== PIVOT_VALUES_FIELD);

  const values: PivotTable['values'] = [];
  for (const df of findChildren(findChild(table, q('dataFields')) ?? table, q('dataField'))) {
    if (!onlyKnownAttrs(df, DATA_FIELD_ATTRS) || df.children.length > 0) return undefined;
    const aggregate = df.attrs['subtotal'] ?? 'sum';
    if (!isAggregate(aggregate)) return undefined;
    const field = Number(df.attrs['fld']);
    values.push({ field, aggregate, ...(df.attrs['name'] !== undefined ? { name: df.attrs['name'] } : {}) });
  }

  const filters: PivotTable['filters'] = [];
  for (const pf of findChildren(findChild(table, q('pageFields')) ?? table, q('pageField'))) {
    if (Object.keys(pf.attrs).some((k) => k !== 'fld' && k !== 'item' && k !== 'hier')) return undefined;
    const field = Number(pf.attrs['fld']);
    const item = pf.attrs['item'];
    if (item === undefined) {
      filters.push({ field });
      continue;
    }
    const pfItems = findChildren(findChild(pivotFields[field] ?? table, q('items')) ?? table, q('item'));
    const shared = Number(pfItems[Number(item)]?.attrs['x'] ?? '-1');
    const sharedItems = findChild(cacheFields[field] ?? cache, q('sharedItems'));
    const selected = cacheItemValue(sharedItems?.children[shared]);
    if (selected === undefined) return undefined;
    filters.push({ field, selected });
  }

  const location = findChild(table, q('location'));
  if (location === undefined) return undefined;
  if ((location.attrs['colPageCount'] ?? '1') !== '1' && filters.length > 0) return undefined;
  const body = parseRange(location.attrs['ref'] ?? '');
  const anchorRow = body.minRow - (filters.length > 0 ? filters.length + 1 : 0);
  if (anchorRow < 1) return undefined;

  // Layout and subtotal placement are per field in the file but per report in
  // the model: the outermost row field decides.
  const lead = pivotFields[rows[0] ?? columns[0] ?? -1];
  const layout: PivotLayout = !isTrue(lead?.attrs['outline'], true)
    ? 'tabular'
    : !isTrue(lead?.attrs['compact'], true)
      ? 'outline'
      : 'compact';
  const subtotals: PivotSubtotals = !isTrue(lead?.attrs['defaultSubtotal'], true)
    ? 'off'
    : isTrue(lead?.attrs['subtotalTop'], true)
      ? 'top'
      : 'bottom';

  return {
    name: table.attrs['name'] ?? '',
    source: { sheet, ref },
    anchor: tupleToCoordinate(body.minCol, anchorRow),
    rows,
    columns,
    filters,
    values,
    layout,
    subtotals,
    rowGrandTotals: isTrue(table.attrs['rowGrandTotals'], true),
    columnGrandTotals: isTrue(table.attrs['colGrandTotals'], true),
    renderedRef: `${tupleToCoordinate(body.minCol, anchorRow)}:${tupleToCoordinate(
      Math.max(body.maxCol, filters.length > 0 ? body.minCol + 1 : body.maxCol),
      body.maxRow,
    )}`,
  };
}

const isAggregate = (s: string): s is PivotAggregate => AGGREGATES.has(s);

/**
 * Turn every passthrough pivot the model can hold into a `ws.pivotTables`
 * definition and drop the parts the writer will regenerate. A pivot cache
 * shared with a pivot that stays passthrough is kept.
 */
export function liftPivotTables(wb: Workbook): void {
  const located: Array<{ p: LocatedPivot; def: PivotTable | undefined }> = [];
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    for (const p of locatePivots(wb, ref.sheet)) located.push({ p, def: toDefinition(wb, p) });
  }
  if (!located.some((l) => l.def !== undefined)) return;
  const cachesInUse = new Set<string>();
  for (const { p, def } of located) if (def === undefined && p.cachePath !== undefined) cachesInUse.add(p.cachePath);

  const drop = (path: string): void => {
    wb.passthrough?.delete(path);
    wb.passthroughContentTypes?.delete(path);
  };
  const droppedCaches = new Set<string>();
  for (const { p, def } of located) {
    if (def === undefined) continue;
    (p.ws.pivotTables ??= []).push(def);
    p.ws.relsExtras = (p.ws.relsExtras ?? []).filter((r) => r.id !== p.relId);
    drop(p.tablePath);
    drop(relsPathFor(p.tablePath));
    if (p.cachePath === undefined || cachesInUse.has(p.cachePath) || droppedCaches.has(p.cachePath)) continue;
    droppedCaches.add(p.cachePath);
    const cacheRels = wb.passthrough?.get(relsPathFor(p.cachePath));
    for (const r of cacheRels ? relsFromBytes(cacheRels).rels : []) {
      if (r.type === PIVOT_CACHE_RECORDS_REL) drop(resolveRelTarget(p.cachePath, r.target));
    }
    drop(relsPathFor(p.cachePath));
    drop(p.cachePath);
  }
  // Unhook the dropped caches from workbook.xml and its rels.
  const droppedRIds = new Set(
    (wb.workbookRelsExtras ?? [])
      .filter((r) => r.type === PIVOT_CACHE_DEFINITION_REL && droppedCaches.has(resolveRelTarget('xl/workbook.xml', r.target)))
      .map((r) => r.id),
  );
  const relsExtras = (wb.workbookRelsExtras ?? []).filter((r) => !droppedRIds.has(r.id));
  if (relsExtras.length > 0) wb.workbookRelsExtras = relsExtras;
  else delete wb.workbookRelsExtras;
  const caches = (wb.pivotCaches ?? []).filter((c) => !droppedRIds.has(c.rId));
  if (caches.length > 0) wb.pivotCaches = caches;
  else delete wb.pivotCaches;
  for (const ref of wb.sheets) {
    if (ref.kind === 'worksheet' && ref.sheet.relsExtras?.length === 0) delete ref.sheet.relsExtras;
  }
}
