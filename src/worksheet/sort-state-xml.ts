// `<sortState>` read / write, shared by the worksheet, its `<autoFilter>` and
// table parts (all use CT_SortState in the main namespace).

import { escapeXmlAttr } from '../utils/escape.js';
import { parseXsdBoolean } from '../utils/xsd-boolean.js';
import { SHEET_MAIN_NS } from '../xml/namespaces.js';
import { findChildren, type XmlNode } from '../xml/tree.js';
import type { SortBy, SortCondition, SortIconSet, SortMethod, SortState } from './sort-state.js';

const SORT_CONDITION_TAG = `{${SHEET_MAIN_NS}}sortCondition`;

const SORT_BY_VALUES: ReadonlyArray<SortBy> = ['value', 'cellColor', 'fontColor', 'icon'];
const SORT_METHODS: ReadonlyArray<SortMethod> = ['stroke', 'pinYin'];
const SORT_ICON_SETS: ReadonlyArray<SortIconSet> = [
  '3Arrows',
  '3ArrowsGray',
  '3Flags',
  '3TrafficLights1',
  '3TrafficLights2',
  '3Signs',
  '3Symbols',
  '3Symbols2',
  '4Arrows',
  '4ArrowsGray',
  '4RedToBlack',
  '4Rating',
  '4TrafficLights',
  '5Arrows',
  '5ArrowsGray',
  '5Rating',
  '5Quarters',
];

/** `<sortState>` → SortState; undefined when it has no `ref`. */
export const parseSortStateNode = (node: XmlNode): SortState | undefined => {
  const ref = node.attrs['ref'];
  if (!ref) return undefined;
  const out: SortState = { ref, conditions: [] };
  const cs = parseXsdBoolean(node.attrs['columnSort']);
  if (cs !== undefined) out.columnSort = cs;
  const cse = parseXsdBoolean(node.attrs['caseSensitive']);
  if (cse !== undefined) out.caseSensitive = cse;
  const sm = node.attrs['sortMethod'];
  if (sm && SORT_METHODS.includes(sm as SortMethod)) out.sortMethod = sm as SortMethod;

  for (const sc of findChildren(node, SORT_CONDITION_TAG)) {
    const cRef = sc.attrs['ref'];
    if (!cRef) continue;
    const c: SortCondition = { ref: cRef };
    const desc = parseXsdBoolean(sc.attrs['descending']);
    if (desc !== undefined) c.descending = desc;
    const sb = sc.attrs['sortBy'];
    if (sb && SORT_BY_VALUES.includes(sb as SortBy)) c.sortBy = sb as SortBy;
    if (sc.attrs['customList'] !== undefined) c.customList = sc.attrs['customList'];
    if (sc.attrs['dxfId'] !== undefined) {
      const n = Number.parseInt(sc.attrs['dxfId'], 10);
      if (Number.isInteger(n)) c.dxfId = n;
    }
    const is = sc.attrs['iconSet'];
    if (is && SORT_ICON_SETS.includes(is as SortIconSet)) c.iconSet = is as SortIconSet;
    if (sc.attrs['iconId'] !== undefined) {
      const n = Number.parseInt(sc.attrs['iconId'], 10);
      if (Number.isInteger(n)) c.iconId = n;
    }
    out.conditions.push(c);
  }
  return out;
};

/** SortState → `<sortState>` XML. */
export const serializeSortStateXml = (ss: SortState): string => {
  let attrs = ` ref="${escapeXmlAttr(ss.ref)}"`;
  if (ss.columnSort !== undefined) attrs += ` columnSort="${ss.columnSort ? '1' : '0'}"`;
  if (ss.caseSensitive !== undefined) attrs += ` caseSensitive="${ss.caseSensitive ? '1' : '0'}"`;
  if (ss.sortMethod !== undefined) attrs += ` sortMethod="${ss.sortMethod}"`;
  if (ss.conditions.length === 0) return `<sortState${attrs}/>`;
  const inner: string[] = [`<sortState${attrs}>`];
  for (const c of ss.conditions) inner.push(serializeSortCondition(c));
  inner.push('</sortState>');
  return inner.join('');
};

const serializeSortCondition = (c: SortCondition): string => {
  let attrs = '';
  if (c.descending !== undefined) attrs += ` descending="${c.descending ? '1' : '0'}"`;
  if (c.sortBy !== undefined) attrs += ` sortBy="${c.sortBy}"`;
  attrs += ` ref="${escapeXmlAttr(c.ref)}"`;
  if (c.customList !== undefined) attrs += ` customList="${escapeXmlAttr(c.customList)}"`;
  if (c.dxfId !== undefined) attrs += ` dxfId="${c.dxfId}"`;
  if (c.iconSet !== undefined) attrs += ` iconSet="${c.iconSet}"`;
  if (c.iconId !== undefined) attrs += ` iconId="${c.iconId}"`;
  return `<sortCondition${attrs}/>`;
};
