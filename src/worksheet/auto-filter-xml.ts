// `<autoFilter>` read / write, shared by the worksheet and table parts (both
// use the same CT_AutoFilter content in the main namespace).

import { escapeXmlAttr } from '../utils/escape.js';
import { SHEET_MAIN_NS } from '../xml/namespaces.js';
import { serializeXml } from '../xml/serializer.js';
import { findChild, findChildren, type XmlNode } from '../xml/tree.js';
import type { AutoFilter, CustomFilterCondition, FilterColumn } from './auto-filter.js';

const FILTER_COLUMN_TAG = `{${SHEET_MAIN_NS}}filterColumn`;
const FILTERS_TAG = `{${SHEET_MAIN_NS}}filters`;
const FILTER_TAG = `{${SHEET_MAIN_NS}}filter`;
const CUSTOM_FILTERS_TAG = `{${SHEET_MAIN_NS}}customFilters`;
const CUSTOM_FILTER_TAG = `{${SHEET_MAIN_NS}}customFilter`;
const TOP10_TAG = `{${SHEET_MAIN_NS}}top10`;
const DYNAMIC_FILTER_TAG = `{${SHEET_MAIN_NS}}dynamicFilter`;

const OPERATORS: ReadonlySet<string> = new Set(['equal', 'lessThan', 'lessThanOrEqual', 'notEqual', 'greaterThanOrEqual', 'greaterThan']);
const isOperator = (op: string): op is NonNullable<CustomFilterCondition['operator']> => OPERATORS.has(op);

const bool = (v: string | undefined): boolean | undefined => (v === undefined ? undefined : v === '1' || v === 'true');

const num = (v: string | undefined): number | undefined => {
  if (v === undefined || v.trim() === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
};

const parseColumn = (fc: XmlNode, colId: number): FilterColumn | undefined => {
  const filters = findChild(fc, FILTERS_TAG);
  if (filters) {
    const values = findChildren(filters, FILTER_TAG).flatMap((f) => (f.attrs['val'] === undefined ? [] : [f.attrs['val']]));
    const blank = bool(filters.attrs['blank']);
    // Date groups (<dateGroupItem>) are not modelled; keep such a list verbatim.
    if (filters.children.every((c) => c.name === FILTER_TAG)) return { kind: 'filters', colId, values, ...(blank !== undefined ? { blank } : {}) };
  }
  const custom = findChild(fc, CUSTOM_FILTERS_TAG);
  if (custom) {
    const conditions = findChildren(custom, CUSTOM_FILTER_TAG).map((c): CustomFilterCondition => {
      const op = c.attrs['operator'];
      const condition: CustomFilterCondition = { val: c.attrs['val'] ?? '' };
      if (op !== undefined && op !== 'equal' && isOperator(op)) condition.operator = op;
      return condition;
    });
    const and = bool(custom.attrs['and']);
    return { kind: 'custom', colId, conditions, ...(and ? { and } : {}) };
  }
  const top10 = findChild(fc, TOP10_TAG);
  const top10Val = num(top10?.attrs['val']);
  if (top10 && top10Val !== undefined) {
    const top = bool(top10.attrs['top']);
    const percent = bool(top10.attrs['percent']);
    const filterVal = num(top10.attrs['filterVal']);
    return {
      kind: 'top10',
      colId,
      val: top10Val,
      ...(top === false ? { top } : {}),
      ...(percent ? { percent } : {}),
      ...(filterVal !== undefined ? { filterVal } : {}),
    };
  }
  const dynamic = findChild(fc, DYNAMIC_FILTER_TAG);
  const type = dynamic?.attrs['type'];
  if (dynamic && type) {
    const val = num(dynamic.attrs['val']);
    const maxVal = num(dynamic.attrs['maxVal']);
    return { kind: 'dynamic', colId, type, ...(val !== undefined ? { val } : {}), ...(maxVal !== undefined ? { maxVal } : {}) };
  }
  if (fc.children.length === 0) return undefined;
  const xml = fc.children.map((c) => new TextDecoder().decode(serializeXml(c, { xmlDeclaration: false }))).join('');
  return { kind: 'raw', colId, xml };
};

/** `<autoFilter>` → AutoFilter; undefined when it has no `ref`. */
export function parseAutoFilterNode(node: XmlNode): AutoFilter | undefined {
  const ref = node.attrs['ref'];
  if (!ref) return undefined;
  const filterColumns: FilterColumn[] = [];
  for (const fc of findChildren(node, FILTER_COLUMN_TAG)) {
    const colId = Number.parseInt(fc.attrs['colId'] ?? '', 10);
    if (!Number.isInteger(colId) || colId < 0) continue;
    const column = parseColumn(fc, colId);
    if (column) filterColumns.push(column);
  }
  return { ref, filterColumns };
}

const serializeColumn = (fc: FilterColumn): string => {
  switch (fc.kind) {
    case 'filters': {
      const attrs = fc.blank !== undefined ? ` blank="${fc.blank ? '1' : '0'}"` : '';
      if (fc.values.length === 0) return `<filters${attrs}/>`;
      return `<filters${attrs}>${fc.values.map((v) => `<filter val="${escapeXmlAttr(v)}"/>`).join('')}</filters>`;
    }
    case 'custom': {
      const items = fc.conditions.map((c) => `<customFilter${c.operator && c.operator !== 'equal' ? ` operator="${c.operator}"` : ''} val="${escapeXmlAttr(c.val)}"/>`);
      return `<customFilters${fc.and ? ' and="1"' : ''}>${items.join('')}</customFilters>`;
    }
    case 'top10': {
      let attrs = '';
      if (fc.top === false) attrs += ' top="0"';
      if (fc.percent) attrs += ' percent="1"';
      attrs += ` val="${fc.val}"`;
      if (fc.filterVal !== undefined) attrs += ` filterVal="${fc.filterVal}"`;
      return `<top10${attrs}/>`;
    }
    case 'dynamic': {
      let attrs = ` type="${escapeXmlAttr(fc.type)}"`;
      if (fc.val !== undefined) attrs += ` val="${fc.val}"`;
      if (fc.maxVal !== undefined) attrs += ` maxVal="${fc.maxVal}"`;
      return `<dynamicFilter${attrs}/>`;
    }
    case 'raw':
      return fc.xml;
  }
};

/** AutoFilter → `<autoFilter>` XML. */
export function serializeAutoFilterXml(filter: AutoFilter): string {
  const ref = ` ref="${escapeXmlAttr(filter.ref)}"`;
  if (filter.filterColumns.length === 0) return `<autoFilter${ref}/>`;
  const columns = filter.filterColumns.map((fc) => `<filterColumn colId="${fc.colId}">${serializeColumn(fc)}</filterColumn>`);
  return `<autoFilter${ref}>${columns.join('')}</autoFilter>`;
}
