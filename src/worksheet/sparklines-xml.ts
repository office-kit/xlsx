// `<x14:sparklineGroups>` reader / writer. The element lives in a worksheet
// `<extLst><ext uri="{05C60535-…}">` entry, so the worksheet reader lifts that
// one entry out of the passthrough `<extLst>` and the writer splices it back.

import type { Color } from '../styles/colors.js';
import { escapeXmlAttr, escapeXmlText } from '../utils/escape.js';
import { X14_NS } from '../xml/namespaces.js';
import { findChild, findChildren, type XmlNode } from '../xml/tree.js';
import type { Sparkline, SparklineAxisType, SparklineEmptyCells, SparklineGroup, SparklineType } from './sparklines.js';

/** `<ext uri>` Excel registers for sparkline groups. */
const SPARKLINE_EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
/** The `xm:` namespace of the `<xm:f>` / `<xm:sqref>` leaves. */
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';

const X14 = (local: string): string => `{${X14_NS}}${local}`;
const XM = (local: string): string => `{${XM_NS}}${local}`;

const COLOR_SLOTS = [
  'colorSeries',
  'colorNegative',
  'colorAxis',
  'colorMarkers',
  'colorFirst',
  'colorLast',
  'colorHigh',
  'colorLow',
] as const;

const FLAG_ATTRS = [
  'dateAxis',
  'markers',
  'high',
  'low',
  'first',
  'last',
  'negative',
  'displayXAxis',
  'displayHidden',
  'rightToLeft',
] as const;

const TYPES: ReadonlyArray<SparklineType> = ['line', 'column', 'stacked'];
const EMPTY_CELLS: ReadonlyArray<SparklineEmptyCells> = ['gap', 'zero', 'span'];
const AXIS_TYPES: ReadonlyArray<SparklineAxisType> = ['individual', 'group', 'custom'];

const pick = <T extends string>(allowed: ReadonlyArray<T>, raw: string | undefined): T | undefined =>
  allowed.find((v) => v === raw);

const parseNumber = (raw: string | undefined): number | undefined => {
  if (raw === undefined) return undefined;
  const n = Number.parseFloat(raw);
  return Number.isFinite(n) ? n : undefined;
};

const parseColor = (el: XmlNode): Color => {
  const out: { rgb?: string; theme?: number; indexed?: number; auto?: boolean; tint?: number } = {};
  const rgb = el.attrs['rgb'];
  if (rgb !== undefined) out.rgb = rgb;
  const theme = parseNumber(el.attrs['theme']);
  if (theme !== undefined && Number.isInteger(theme)) out.theme = theme;
  const indexed = parseNumber(el.attrs['indexed']);
  if (indexed !== undefined && Number.isInteger(indexed)) out.indexed = indexed;
  const auto = el.attrs['auto'];
  if (auto === '1' || auto === 'true') out.auto = true;
  const tint = parseNumber(el.attrs['tint']);
  if (tint !== undefined) out.tint = tint;
  return out;
};

const parseSparklineGroup = (el: XmlNode): SparklineGroup => {
  const group: SparklineGroup = { sparklines: [] };
  const type = pick(TYPES, el.attrs['type']);
  if (type) group.type = type;
  for (const name of FLAG_ATTRS) {
    const raw = el.attrs[name];
    if (raw !== undefined) group[name] = raw === '1' || raw === 'true';
  }
  const empty = pick(EMPTY_CELLS, el.attrs['displayEmptyCellsAs']);
  if (empty) group.displayEmptyCellsAs = empty;
  const minAxis = pick(AXIS_TYPES, el.attrs['minAxisType']);
  if (minAxis) group.minAxisType = minAxis;
  const maxAxis = pick(AXIS_TYPES, el.attrs['maxAxisType']);
  if (maxAxis) group.maxAxisType = maxAxis;
  const lineWeight = parseNumber(el.attrs['lineWeight']);
  if (lineWeight !== undefined) group.lineWeight = lineWeight;
  const manualMin = parseNumber(el.attrs['manualMin']);
  if (manualMin !== undefined) group.manualMin = manualMin;
  const manualMax = parseNumber(el.attrs['manualMax']);
  if (manualMax !== undefined) group.manualMax = manualMax;
  for (const slot of COLOR_SLOTS) {
    const c = findChild(el, X14(slot));
    if (c) group[slot] = parseColor(c);
  }
  const dateFormula = findChild(el, XM('f'))?.text;
  if (dateFormula) group.dateFormula = dateFormula;
  const list = findChild(el, X14('sparklines'));
  for (const s of list ? findChildren(list, X14('sparkline')) : []) {
    const formula = findChild(s, XM('f'))?.text ?? '';
    const location = findChild(s, XM('sqref'))?.text;
    // A sparkline without a location has nowhere to draw; Excel drops it too.
    if (location) group.sparklines.push({ formula, location });
  }
  return group;
};

/** Parse an `<x14:sparklineGroups>` element. */
export function parseSparklineGroups(el: XmlNode): SparklineGroup[] {
  return findChildren(el, X14('sparklineGroup')).map(parseSparklineGroup);
}

const serializeColor = (tag: string, c: Color): string => {
  let attrs = '';
  if (c.auto) attrs += ' auto="1"';
  if (c.rgb !== undefined) attrs += ` rgb="${escapeXmlAttr(c.rgb)}"`;
  if (c.theme !== undefined) attrs += ` theme="${c.theme}"`;
  if (c.indexed !== undefined) attrs += ` indexed="${c.indexed}"`;
  if (c.tint !== undefined) attrs += ` tint="${c.tint}"`;
  return `<x14:${tag}${attrs}/>`;
};

const serializeSparkline = (s: Sparkline): string =>
  `<x14:sparkline><xm:f>${escapeXmlText(s.formula)}</xm:f><xm:sqref>${escapeXmlText(s.location)}</xm:sqref></x14:sparkline>`;

const serializeSparklineGroup = (g: SparklineGroup): string => {
  let attrs = '';
  if (g.manualMax !== undefined) attrs += ` manualMax="${g.manualMax}"`;
  if (g.manualMin !== undefined) attrs += ` manualMin="${g.manualMin}"`;
  if (g.lineWeight !== undefined) attrs += ` lineWeight="${g.lineWeight}"`;
  if (g.type !== undefined && g.type !== 'line') attrs += ` type="${g.type}"`;
  for (const name of FLAG_ATTRS) {
    const v = g[name];
    if (v !== undefined) attrs += ` ${name}="${v ? '1' : '0'}"`;
  }
  if (g.displayEmptyCellsAs !== undefined) attrs += ` displayEmptyCellsAs="${g.displayEmptyCellsAs}"`;
  if (g.minAxisType !== undefined) attrs += ` minAxisType="${g.minAxisType}"`;
  if (g.maxAxisType !== undefined) attrs += ` maxAxisType="${g.maxAxisType}"`;
  const parts: string[] = [`<x14:sparklineGroup${attrs}>`];
  for (const slot of COLOR_SLOTS) {
    const c = g[slot];
    if (c) parts.push(serializeColor(slot, c));
  }
  if (g.dateFormula !== undefined) parts.push(`<xm:f>${escapeXmlText(g.dateFormula)}</xm:f>`);
  parts.push('<x14:sparklines>', ...g.sparklines.map(serializeSparkline), '</x14:sparklines>');
  parts.push('</x14:sparklineGroup>');
  return parts.join('');
};

/** Serialise sparkline groups as the complete `<ext>` entry for a worksheet `<extLst>`. */
export function serializeSparklineExt(groups: ReadonlyArray<SparklineGroup>): string {
  return [
    `<ext uri="${SPARKLINE_EXT_URI}" xmlns:x14="${X14_NS}">`,
    `<x14:sparklineGroups xmlns:xm="${XM_NS}">`,
    ...groups.map(serializeSparklineGroup),
    '</x14:sparklineGroups>',
    '</ext>',
  ].join('');
}

/** The `<x14:sparklineGroups>` inside a worksheet `<ext>` entry, if that entry is the sparkline one. */
export function findSparklineGroupsInExt(ext: XmlNode): XmlNode | undefined {
  if (ext.attrs['uri'] !== SPARKLINE_EXT_URI) return undefined;
  return findChild(ext, X14('sparklineGroups'));
}
