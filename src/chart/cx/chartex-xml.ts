// xl/charts/chartN.xml read/write for the chartex (cx:) namespace.

import {
  parseShapeProperties,
  parseTextBody,
  serializeShapeProperties,
  serializeTextBody,
} from '../../drawing/dml/dml-xml.js';
import type { ShapeProperties } from '../../drawing/dml/shape-properties.js';
import type { TextBody } from '../../drawing/dml/text.js';
import { escapeXmlAttr, escapeXmlText } from '../../utils/escape.js';
import { OpenXmlSchemaError } from '../../utils/exceptions.js';
import { CX_NS, REL_NS } from '../../xml/namespaces.js';
import { parseXml } from '../../xml/parser.js';
import { findChild, findChildren, type XmlNode } from '../../xml/tree.js';
import type {
  CxAxis,
  CxChart,
  CxChartData,
  CxChartSpace,
  CxData,
  CxDataLabels,
  CxDim,
  CxLayoutId,
  CxLayoutPr,
  CxLegend,
  CxNumDim,
  CxPoint,
  CxSeries,
  CxStrDim,
  CxTitle,
} from './chartex.js';

const T = (local: string): string => `{${CX_NS}}${local}`;
const CHART_SPACE = T('chartSpace');
const CHART_DATA = T('chartData');
const EXTERNAL_DATA = T('externalData');
const DATA = T('data');
const NUM_DIM = T('numDim');
const STR_DIM = T('strDim');
const F = T('f');
const FORMAT_CODE = T('formatCode');
const LVL = T('lvl');
const PT = T('pt');
const CHART = T('chart');
const TITLE = T('title');
const TX = T('tx');
const TX_DATA = T('txData');
const V = T('v');
const PLOT_AREA = T('plotArea');
const PLOT_AREA_REGION = T('plotAreaRegion');
const SERIES = T('series');
const DATA_ID = T('dataId');
const LAYOUT_PR = T('layoutPr');
const SUBTOTALS = T('subtotals');
const SUBTOTAL = T('subtotal');
const BINNING = T('binning');
const PARENT_LABEL_LAYOUT = T('parentLabelLayout');
const VISIBILITY = T('visibility');
const STATISTICS = T('statistics');
const BIN_COUNT = T('binCount');
const BIN_SIZE = T('binSize');
const GEOGRAPHY = T('geography');
const REGION_LABEL_LAYOUT = T('regionLabelLayout');
const AXIS = T('axis');
const AXIS_ID = T('axisId');
const VAL_SCALING = T('valScaling');
const CAT_SCALING = T('catScaling');
const MAJOR_GRIDLINES = T('majorGridlines');
const DATA_LABELS = T('dataLabels');
const VISIBILITY_DL = T('visibility');
const LEGEND = T('legend');
const PLOT_VIS_ONLY = T('plotVisOnly');
const DISP_BLANKS_AS = T('dispBlanksAs');
const SP_PR = T('spPr');
const TX_PR = T('txPr');
const PLOT_SURFACE = T('plotSurface');
const A_T = '{http://schemas.openxmlformats.org/drawingml/2006/main}t';

const parseSpPrSlot = (parent: XmlNode): ShapeProperties | undefined => {
  const el = findChild(parent, SP_PR);
  return el ? parseShapeProperties(el) : undefined;
};

const parseTxPrSlot = (parent: XmlNode): TextBody | undefined => {
  const el = findChild(parent, TX_PR);
  return el ? parseTextBody(el) : undefined;
};

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeText = escapeXmlText;
const escapeAttr = escapeXmlAttr;

const valAttr = (n: XmlNode | undefined): string | undefined => n?.attrs['val'];
const intAttr = (n: XmlNode, name: string): number | undefined => {
  const v = n.attrs[name];
  if (v === undefined) return undefined;
  const x = Number.parseInt(v, 10);
  return Number.isInteger(x) ? x : undefined;
};
const floatAttr = (n: XmlNode, name: string): number | undefined => {
  const v = n.attrs[name];
  if (v === undefined) return undefined;
  const x = Number.parseFloat(v);
  return Number.isFinite(x) ? x : undefined;
};
const boolAttr = (n: XmlNode, name: string): boolean | undefined => {
  const v = n.attrs[name];
  if (v === undefined) return undefined;
  if (v === '1' || v === 'true') return true;
  if (v === '0' || v === 'false') return false;
  return undefined;
};

/** Best-effort sniff: returns true if `bytes` starts with a `cx:chartSpace` root. */
export function isChartExBytes(bytes: Uint8Array | string): boolean {
  const head = typeof bytes === 'string' ? bytes.slice(0, 4096) : new TextDecoder().decode(bytes.subarray(0, 4096));
  // Match an opening <cx:chartSpace ...> or <chartSpace xmlns="...chartex">.
  return /<\s*[A-Za-z][\w-]*:?chartSpace\b[^>]*chartex/.test(head);
}

// ---- parser ----------------------------------------------------------------

const parsePoints = (lvlEl: XmlNode | undefined): { ptCount?: number; pts: CxPoint[] } => {
  if (!lvlEl) return { pts: [] };
  const ptCount = intAttr(lvlEl, 'ptCount');
  const pts: CxPoint[] = [];
  for (const ptEl of findChildren(lvlEl, PT)) {
    const idx = intAttr(ptEl, 'idx');
    if (idx === undefined) continue;
    pts.push({ idx, v: ptEl.text ?? '' });
  }
  return { ...(ptCount !== undefined ? { ptCount } : {}), pts };
};

const parseDim = (el: XmlNode): CxDim => {
  const isNum = el.name === NUM_DIM;
  const type = el.attrs['type'] ?? (isNum ? 'val' : 'cat');
  const fEl = findChild(el, F);
  const fc = findChild(el, FORMAT_CODE);
  const dirRaw = fEl?.attrs['dir'];
  const dir = dirRaw === 'col' || dirRaw === 'row' ? dirRaw : undefined;
  const lvlEl = findChild(el, LVL);
  const { ptCount, pts } = parsePoints(lvlEl);
  const base = {
    type,
    ...(fEl ? { f: fEl.text ?? '' } : {}),
    ...(dir ? { dir } : {}),
    ...(ptCount !== undefined ? { ptCount } : {}),
    pts,
    ...(fc?.text ? { formatCode: fc.text } : {}),
  };
  return isNum ? ({ kind: 'num', ...base } as CxNumDim) : ({ kind: 'str', ...base } as CxStrDim);
};

const parseDataBlock = (el: XmlNode): CxData => {
  const id = intAttr(el, 'id') ?? 0;
  const dims: CxDim[] = [];
  for (const child of el.children) {
    if (typeof child === 'string') continue;
    if (child.name === NUM_DIM || child.name === STR_DIM) dims.push(parseDim(child));
  }
  return { id, dims };
};

const parseChartData = (el: XmlNode): CxChartData => {
  const data: CxData[] = [];
  for (const dEl of findChildren(el, DATA)) data.push(parseDataBlock(dEl));
  const ext = findChild(el, EXTERNAL_DATA);
  let externalData: CxChartData['externalData'];
  if (ext) {
    const rId = ext.attrs[`{${REL_NS}}id`];
    if (rId !== undefined) {
      const auto = ext.attrs['autoUpdate'];
      const autoUpdate = auto === '1' || auto === 'true' ? true : auto === '0' || auto === 'false' ? false : undefined;
      externalData = { rId, ...(autoUpdate !== undefined ? { autoUpdate } : {}) };
    }
  }
  return { ...(externalData ? { externalData } : {}), data };
};

const parseTitle = (el: XmlNode): CxTitle => {
  const posRaw = el.attrs['pos'];
  const pos = posRaw === 't' || posRaw === 'b' || posRaw === 'l' || posRaw === 'r' ? posRaw : undefined;
  const alignRaw = el.attrs['align'];
  const align = alignRaw === 'ctr' || alignRaw === 'l' || alignRaw === 'r' ? alignRaw : undefined;
  const overlay = boolAttr(el, 'overlay');
  // Plain text body inside <cx:tx><cx:rich>... <a:r><a:t>text</a:t></a:r>
  // </cx:rich></cx:tx>
  let text: string | undefined;
  let txDataRef: string | undefined;
  const tx = findChild(el, TX);
  if (tx) {
    const txData = findChild(tx, TX_DATA);
    if (txData) {
      const fEl = findChild(txData, F);
      const vEl = findChild(txData, V);
      if (fEl?.text) txDataRef = fEl.text;
      if (vEl?.text !== undefined) text = vEl.text;
    } else {
      // Walk into <cx:rich><a:p><a:r><a:t> if present.
      const rich = tx.children.find((c): c is XmlNode => typeof c !== 'string' && c.name === T('rich'));
      if (rich) {
        const collected: string[] = [];
        const walk = (n: XmlNode): void => {
          if (n.name === A_T && n.text) collected.push(n.text);
          for (const c of n.children) if (typeof c !== 'string') walk(c);
        };
        walk(rich);
        if (collected.length > 0) text = collected.join('');
      }
    }
  }
  const spPr = parseSpPrSlot(el);
  const txPr = parseTxPrSlot(el);
  return {
    ...(pos ? { pos } : {}),
    ...(align ? { align } : {}),
    ...(overlay !== undefined ? { overlay } : {}),
    ...(text !== undefined ? { text } : {}),
    ...(txDataRef !== undefined ? { txDataRef } : {}),
    ...(spPr ? { spPr } : {}),
    ...(txPr ? { txPr } : {}),
  };
};

const parseLegend = (el: XmlNode): CxLegend => {
  const posRaw = el.attrs['pos'];
  const pos =
    posRaw === 't' || posRaw === 'b' || posRaw === 'l' || posRaw === 'r' || posRaw === 'tr' ? posRaw : undefined;
  const alignRaw = el.attrs['align'];
  const align = alignRaw === 'ctr' || alignRaw === 'l' || alignRaw === 'r' ? alignRaw : undefined;
  const overlay = boolAttr(el, 'overlay');
  const spPr = parseSpPrSlot(el);
  const txPr = parseTxPrSlot(el);
  return {
    ...(pos ? { pos } : {}),
    ...(align ? { align } : {}),
    ...(overlay !== undefined ? { overlay } : {}),
    ...(spPr ? { spPr } : {}),
    ...(txPr ? { txPr } : {}),
  };
};

const parseLayoutPr = (el: XmlNode): CxLayoutPr | undefined => {
  const subEl = findChild(el, SUBTOTALS);
  if (subEl) {
    const idxs: number[] = [];
    for (const s of findChildren(subEl, SUBTOTAL)) {
      const i = intAttr(s, 'idx');
      if (i !== undefined) idxs.push(i);
    }
    return { kind: 'waterfall', subtotalIdx: idxs };
  }
  const binEl = findChild(el, BINNING);
  if (binEl) {
    const intervalRaw = binEl.attrs['intervalClosed'];
    const intervalClosed = intervalRaw === 'r' || intervalRaw === 'l' ? intervalRaw : undefined;
    // The bin count / size are child elements; neither means automatic bins.
    const countEl = findChild(binEl, BIN_COUNT);
    const sizeEl = findChild(binEl, BIN_SIZE);
    const binCount = countEl ? intAttr(countEl, 'val') : undefined;
    const binSize = sizeEl ? floatAttr(sizeEl, 'val') : undefined;
    return {
      kind: 'binning',
      ...(binCount !== undefined ? { binCount } : {}),
      ...(binSize !== undefined ? { binSize } : {}),
      ...(intervalClosed ? { intervalClosed } : {}),
      ...(floatAttr(binEl, 'underflow') !== undefined ? { underflow: floatAttr(binEl, 'underflow') as number } : {}),
      ...(floatAttr(binEl, 'overflow') !== undefined ? { overflow: floatAttr(binEl, 'overflow') as number } : {}),
    };
  }
  const plEl = findChild(el, PARENT_LABEL_LAYOUT);
  if (plEl) {
    const v = valAttr(plEl);
    const layout = v === 'banner' || v === 'none' ? v : 'overlapping';
    return { kind: 'parentLabel', layout };
  }
  const visEl = findChild(el, VISIBILITY);
  const statsEl = findChild(el, STATISTICS);
  if (visEl || statsEl) {
    const meanLine = visEl ? boolAttr(visEl, 'meanLine') : undefined;
    const meanMarker = visEl ? boolAttr(visEl, 'meanMarker') : undefined;
    const nonoutliers = visEl ? boolAttr(visEl, 'nonoutliers') : undefined;
    const outliers = visEl ? boolAttr(visEl, 'outliers') : undefined;
    const qmRaw = statsEl?.attrs['quartileMethod'];
    const quartileMethod = qmRaw === 'exclusive' || qmRaw === 'inclusive' ? qmRaw : undefined;
    return {
      kind: 'visibility',
      ...(meanLine !== undefined ? { meanLine } : {}),
      ...(meanMarker !== undefined ? { meanMarker } : {}),
      ...(nonoutliers !== undefined ? { nonoutliers } : {}),
      ...(outliers !== undefined ? { outliers } : {}),
      ...(quartileMethod ? { quartileMethod } : {}),
    };
  }
  const geoEl = findChild(el, GEOGRAPHY);
  if (geoEl) {
    const projRaw = geoEl.attrs['projectionType'];
    const proj =
      projRaw === 'automatic' || projRaw === 'mercator' || projRaw === 'miller' || projRaw === 'albers'
        ? projRaw
        : undefined;
    const rllEl = findChild(geoEl, REGION_LABEL_LAYOUT);
    const rllRaw = rllEl ? valAttr(rllEl) : undefined;
    const rll = rllRaw === 'none' || rllRaw === 'bestFit' || rllRaw === 'showAll' ? rllRaw : undefined;
    return {
      kind: 'region',
      ...(geoEl.attrs['cultureLanguage'] !== undefined
        ? { cultureLanguage: geoEl.attrs['cultureLanguage'] as string }
        : {}),
      ...(geoEl.attrs['cultureRegion'] !== undefined ? { cultureRegion: geoEl.attrs['cultureRegion'] as string } : {}),
      ...(proj ? { projectionType: proj } : {}),
      ...(rll ? { regionLabelLayout: rll } : {}),
    };
  }
  return undefined;
};

const parseDataLabels = (el: XmlNode): CxDataLabels => {
  const posRaw = el.attrs['pos'];
  const allowed: ReadonlyArray<string> = ['ctr', 'b', 'l', 'r', 't', 'inEnd', 'outEnd', 'inBase'];
  const pos = posRaw && allowed.includes(posRaw) ? (posRaw as CxDataLabels['pos']) : undefined;
  const visEl = findChild(el, VISIBILITY_DL);
  const out: CxDataLabels = { ...(pos ? { pos } : {}) };
  if (visEl) {
    const seriesName = boolAttr(visEl, 'seriesName');
    const categoryName = boolAttr(visEl, 'categoryName');
    const value = boolAttr(visEl, 'value');
    if (seriesName !== undefined || categoryName !== undefined || value !== undefined) {
      out.visibility = {
        ...(seriesName !== undefined ? { seriesName } : {}),
        ...(categoryName !== undefined ? { categoryName } : {}),
        ...(value !== undefined ? { value } : {}),
      };
    }
  }
  return out;
};

const parseSeries = (el: XmlNode): CxSeries => {
  const layoutIdRaw = el.attrs['layoutId'];
  if (!layoutIdRaw) {
    throw new OpenXmlSchemaError('parseChartExXml: <cx:series> missing layoutId');
  }
  const layoutId = layoutIdRaw as CxLayoutId;
  const hidden = boolAttr(el, 'hidden');
  const ownerIdx = intAttr(el, 'ownerIdx');
  const formatIdx = intAttr(el, 'formatIdx');
  const axisIds: number[] = [];
  for (const aEl of findChildren(el, AXIS_ID)) {
    const v = intAttr(aEl, 'val');
    if (v !== undefined) axisIds.push(v);
  }
  const dataIdEl = findChild(el, DATA_ID);
  const dataId = dataIdEl ? intAttr(dataIdEl, 'val') : undefined;
  const layoutPrEl = findChild(el, LAYOUT_PR);
  const layoutPr = layoutPrEl ? parseLayoutPr(layoutPrEl) : undefined;
  let tx: CxSeries['tx'];
  const txEl = findChild(el, TX);
  if (txEl) {
    const txData = findChild(txEl, TX_DATA);
    if (txData) {
      const fEl = findChild(txData, F);
      const vEl = findChild(txData, V);
      const f = fEl?.text;
      const v = vEl?.text;
      if (f !== undefined || v !== undefined) {
        tx = { ...(f !== undefined ? { f } : {}), ...(v !== undefined ? { v } : {}) };
      }
    }
  }
  const dlEl = findChild(el, DATA_LABELS);
  const dataLabels = dlEl ? parseDataLabels(dlEl) : undefined;
  const spPr = parseSpPrSlot(el);
  const txPr = parseTxPrSlot(el);
  return {
    layoutId,
    ...(hidden !== undefined ? { hidden } : {}),
    ...(ownerIdx !== undefined ? { ownerIdx } : {}),
    ...(formatIdx !== undefined ? { formatIdx } : {}),
    ...(axisIds.length > 0 ? { axisIds } : {}),
    ...(tx ? { tx } : {}),
    ...(dataLabels ? { dataLabels } : {}),
    ...(dataId !== undefined ? { dataId } : {}),
    ...(layoutPr ? { layoutPr } : {}),
    ...(spPr ? { spPr } : {}),
    ...(txPr ? { txPr } : {}),
  };
};

const parseAxis = (el: XmlNode): CxAxis => {
  const id = intAttr(el, 'id') ?? 0;
  const hidden = boolAttr(el, 'hidden');
  const valScalingEl = findChild(el, VAL_SCALING);
  let valScaling: CxAxis['valScaling'];
  if (valScalingEl) {
    const min = floatAttr(valScalingEl, 'min');
    const max = floatAttr(valScalingEl, 'max');
    if (min !== undefined || max !== undefined) {
      valScaling = {
        ...(min !== undefined ? { min } : {}),
        ...(max !== undefined ? { max } : {}),
      };
    }
  }
  const catScalingEl = findChild(el, CAT_SCALING);
  const catGap = catScalingEl ? floatAttr(catScalingEl, 'gapWidth') : undefined;
  const majorGridlines = findChild(el, MAJOR_GRIDLINES) !== undefined ? true : undefined;
  const titleEl = findChild(el, TITLE);
  const title = titleEl ? parseTitle(titleEl) : undefined;
  const spPr = parseSpPrSlot(el);
  const txPr = parseTxPrSlot(el);
  return {
    id,
    ...(hidden !== undefined ? { hidden } : {}),
    ...(valScaling ? { valScaling } : {}),
    ...(catGap !== undefined ? { catScalingGapWidth: catGap } : {}),
    ...(majorGridlines !== undefined ? { majorGridlines } : {}),
    ...(title ? { title } : {}),
    ...(spPr ? { spPr } : {}),
    ...(txPr ? { txPr } : {}),
  };
};

const parseChart = (el: XmlNode): CxChart => {
  const titleEl = findChild(el, TITLE);
  const title = titleEl ? parseTitle(titleEl) : undefined;
  const plotAreaEl = findChild(el, PLOT_AREA);
  if (!plotAreaEl) throw new OpenXmlSchemaError('parseChartExXml: <cx:chart> missing <cx:plotArea>');
  const region = findChild(plotAreaEl, PLOT_AREA_REGION);
  if (!region) throw new OpenXmlSchemaError('parseChartExXml: <cx:plotArea> missing <cx:plotAreaRegion>');
  const series: CxSeries[] = [];
  for (const sEl of findChildren(region, SERIES)) series.push(parseSeries(sEl));
  const axes: CxAxis[] = [];
  // Axes live as siblings of <cx:plotAreaRegion> under <cx:plotArea>, but
  // tolerate writers (including older revisions of this library) that nest them
  // inside the region.
  for (const aEl of findChildren(plotAreaEl, AXIS)) axes.push(parseAxis(aEl));
  for (const aEl of findChildren(region, AXIS)) axes.push(parseAxis(aEl));
  // <cx:plotSurface> sits inside plotAreaRegion and carries the plot background
  // spPr.
  const plotSurfaceEl = findChild(region, PLOT_SURFACE);
  const plotSurfaceSpPr = plotSurfaceEl ? parseShapeProperties(plotSurfaceEl) : undefined;
  const legendEl = findChild(el, LEGEND);
  const legend = legendEl ? parseLegend(legendEl) : undefined;
  const plotVisOnlyEl = findChild(el, PLOT_VIS_ONLY);
  const plotVisOnly = plotVisOnlyEl ? boolAttr(plotVisOnlyEl, 'val') : undefined;
  const dispRaw = valAttr(findChild(el, DISP_BLANKS_AS));
  const dispBlanksAs = dispRaw === 'gap' || dispRaw === 'zero' || dispRaw === 'span' ? dispRaw : undefined;
  return {
    plotArea: { series, axes, ...(plotSurfaceSpPr ? { spPr: plotSurfaceSpPr } : {}) },
    ...(title ? { title } : {}),
    ...(legend ? { legend } : {}),
    ...(plotVisOnly !== undefined ? { plotVisOnly } : {}),
    ...(dispBlanksAs ? { dispBlanksAs } : {}),
  };
};

export function parseChartExXml(bytes: Uint8Array | string): CxChartSpace {
  const root = parseXml(bytes);
  if (root.name !== CHART_SPACE) {
    throw new OpenXmlSchemaError(`parseChartExXml: root is "${root.name}", expected cx:chartSpace`);
  }
  const chartDataEl = findChild(root, CHART_DATA);
  if (!chartDataEl) throw new OpenXmlSchemaError('parseChartExXml: <cx:chartSpace> missing <cx:chartData>');
  const chartEl = findChild(root, CHART);
  if (!chartEl) throw new OpenXmlSchemaError('parseChartExXml: <cx:chartSpace> missing <cx:chart>');
  const spPr = parseSpPrSlot(root);
  const txPr = parseTxPrSlot(root);
  return {
    kind: 'cxChartSpace',
    chartData: parseChartData(chartDataEl),
    chart: parseChart(chartEl),
    ...(spPr ? { spPr } : {}),
    ...(txPr ? { txPr } : {}),
  };
}

// ---- serializer ------------------------------------------------------------

const serializePoint = (p: CxPoint): string => `<cx:pt idx="${p.idx}">${escapeText(p.v)}</cx:pt>`;

const serializeDim = (d: CxDim): string => {
  const tag = d.kind === 'num' ? 'cx:numDim' : 'cx:strDim';
  const parts: string[] = [`<${tag} type="${escapeAttr(d.type)}">`];
  if (d.f !== undefined) {
    const dirAttr = d.dir ? ` dir="${d.dir}"` : '';
    parts.push(`<cx:f${dirAttr}>${escapeText(d.f)}</cx:f>`);
  }
  if (d.formatCode !== undefined) parts.push(`<cx:formatCode>${escapeText(d.formatCode)}</cx:formatCode>`);
  const ptCount = d.ptCount ?? d.pts.length;
  if (d.pts.length > 0 || d.ptCount !== undefined) {
    parts.push(`<cx:lvl ptCount="${ptCount}">`);
    for (const p of d.pts) parts.push(serializePoint(p));
    parts.push('</cx:lvl>');
  }
  parts.push(`</${tag}>`);
  return parts.join('');
};

const serializeData = (d: CxData): string => {
  const parts: string[] = [`<cx:data id="${d.id}">`];
  for (const dim of d.dims) parts.push(serializeDim(dim));
  parts.push('</cx:data>');
  return parts.join('');
};

const serializeChartData = (cd: CxChartData): string => {
  const parts: string[] = ['<cx:chartData>'];
  if (cd.externalData) {
    const auto =
      cd.externalData.autoUpdate !== undefined ? ` cx:autoUpdate="${cd.externalData.autoUpdate ? '1' : '0'}"` : '';
    parts.push(`<cx:externalData r:id="${escapeAttr(cd.externalData.rId)}"${auto}/>`);
  }
  for (const d of cd.data) parts.push(serializeData(d));
  parts.push('</cx:chartData>');
  return parts.join('');
};

const serializeTitle = (t: CxTitle, tag: 'cx:title' = 'cx:title'): string => {
  const attrs: string[] = [];
  if (t.pos) attrs.push(`pos="${t.pos}"`);
  if (t.align) attrs.push(`align="${t.align}"`);
  if (t.overlay !== undefined) attrs.push(`overlay="${t.overlay ? '1' : '0'}"`);
  const open = `<${tag}${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}`;
  const innerParts: string[] = [];
  if (t.text !== undefined || t.txDataRef !== undefined) {
    const f = t.txDataRef !== undefined ? `<cx:f>${escapeText(t.txDataRef)}</cx:f>` : '';
    const v = t.text !== undefined ? `<cx:v>${escapeText(t.text)}</cx:v>` : '';
    innerParts.push(`<cx:tx><cx:txData>${f}${v}</cx:txData></cx:tx>`);
  }
  if (t.spPr) innerParts.push(serializeShapeProperties(t.spPr, 'cx:spPr'));
  if (t.txPr) innerParts.push(serializeTextBody(t.txPr, 'cx:txPr'));
  return innerParts.length === 0 ? `${open}/>` : `${open}>${innerParts.join('')}</${tag}>`;
};

const serializeLegend = (l: CxLegend): string => {
  const attrs: string[] = [];
  if (l.pos) attrs.push(`pos="${l.pos}"`);
  if (l.align) attrs.push(`align="${l.align}"`);
  if (l.overlay !== undefined) attrs.push(`overlay="${l.overlay ? '1' : '0'}"`);
  const open = `<cx:legend${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}`;
  const innerParts: string[] = [];
  if (l.spPr) innerParts.push(serializeShapeProperties(l.spPr, 'cx:spPr'));
  if (l.txPr) innerParts.push(serializeTextBody(l.txPr, 'cx:txPr'));
  return innerParts.length === 0 ? `${open}/>` : `${open}>${innerParts.join('')}</cx:legend>`;
};

const serializeLayoutPr = (lp: CxLayoutPr): string => {
  switch (lp.kind) {
    case 'waterfall': {
      if (lp.subtotalIdx.length === 0) return '<cx:layoutPr><cx:subtotals/></cx:layoutPr>';
      const parts = ['<cx:layoutPr><cx:subtotals>'];
      for (const i of lp.subtotalIdx) parts.push(`<cx:subtotal idx="${i}"/>`);
      parts.push('</cx:subtotals></cx:layoutPr>');
      return parts.join('');
    }
    case 'binning': {
      // ECMA's CT_Binning: the interval / flow bounds are attributes, the bin
      // count or size a child element. Excel asks to repair anything else.
      const a: string[] = [];
      if (lp.intervalClosed !== undefined) a.push(`intervalClosed="${lp.intervalClosed}"`);
      if (lp.underflow !== undefined) a.push(`underflow="${lp.underflow}"`);
      if (lp.overflow !== undefined) a.push(`overflow="${lp.overflow}"`);
      // Automatic binning (`binCountAuto`) is the default and has no element:
      // Excel refuses an empty `<cx:binCount/>`.
      const child =
        lp.binSize !== undefined
          ? `<cx:binSize val="${lp.binSize}"/>`
          : lp.binCount !== undefined
            ? `<cx:binCount val="${lp.binCount}"/>`
            : '';
      const open = `<cx:binning${a.length > 0 ? ` ${a.join(' ')}` : ''}`;
      return `<cx:layoutPr>${child ? `${open}>${child}</cx:binning>` : `${open}/>`}</cx:layoutPr>`;
    }
    case 'parentLabel':
      return `<cx:layoutPr><cx:parentLabelLayout val="${lp.layout}"/></cx:layoutPr>`;
    case 'visibility': {
      const va: string[] = [];
      if (lp.meanLine !== undefined) va.push(`meanLine="${lp.meanLine ? '1' : '0'}"`);
      if (lp.meanMarker !== undefined) va.push(`meanMarker="${lp.meanMarker ? '1' : '0'}"`);
      if (lp.nonoutliers !== undefined) va.push(`nonoutliers="${lp.nonoutliers ? '1' : '0'}"`);
      if (lp.outliers !== undefined) va.push(`outliers="${lp.outliers ? '1' : '0'}"`);
      const visTag = `<cx:visibility${va.length > 0 ? ` ${va.join(' ')}` : ''}/>`;
      const qmTag = lp.quartileMethod !== undefined ? `<cx:statistics quartileMethod="${lp.quartileMethod}"/>` : '';
      return `<cx:layoutPr>${visTag}${qmTag}</cx:layoutPr>`;
    }
    case 'region': {
      const a: string[] = [];
      if (lp.cultureLanguage !== undefined) a.push(`cultureLanguage="${escapeAttr(lp.cultureLanguage)}"`);
      if (lp.cultureRegion !== undefined) a.push(`cultureRegion="${escapeAttr(lp.cultureRegion)}"`);
      if (lp.projectionType !== undefined) a.push(`projectionType="${lp.projectionType}"`);
      const rll = lp.regionLabelLayout !== undefined ? `<cx:regionLabelLayout val="${lp.regionLabelLayout}"/>` : '';
      return `<cx:layoutPr><cx:geography${a.length > 0 ? ` ${a.join(' ')}` : ''}>${rll}</cx:geography></cx:layoutPr>`;
    }
  }
};

const serializeDataLabels = (dl: CxDataLabels): string => {
  const attrs: string[] = [];
  if (dl.pos) attrs.push(`pos="${dl.pos}"`);
  let inner = '';
  if (dl.visibility) {
    const va: string[] = [];
    if (dl.visibility.seriesName !== undefined) va.push(`seriesName="${dl.visibility.seriesName ? '1' : '0'}"`);
    if (dl.visibility.categoryName !== undefined) va.push(`categoryName="${dl.visibility.categoryName ? '1' : '0'}"`);
    if (dl.visibility.value !== undefined) va.push(`value="${dl.visibility.value ? '1' : '0'}"`);
    inner = `<cx:visibility${va.length > 0 ? ` ${va.join(' ')}` : ''}/>`;
  }
  const open = `<cx:dataLabels${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}`;
  return inner === '' ? `${open}/>` : `${open}>${inner}</cx:dataLabels>`;
};

const serializeSeries = (s: CxSeries): string => {
  const attrs: string[] = [`layoutId="${s.layoutId}"`];
  if (s.hidden !== undefined) attrs.push(`hidden="${s.hidden ? '1' : '0'}"`);
  if (s.ownerIdx !== undefined) attrs.push(`ownerIdx="${s.ownerIdx}"`);
  if (s.formatIdx !== undefined) attrs.push(`formatIdx="${s.formatIdx}"`);
  const parts: string[] = [`<cx:series ${attrs.join(' ')}>`];
  if (s.tx) {
    const f = s.tx.f !== undefined ? `<cx:f>${escapeText(s.tx.f)}</cx:f>` : '';
    const v = s.tx.v !== undefined ? `<cx:v>${escapeText(s.tx.v)}</cx:v>` : '';
    parts.push(`<cx:tx><cx:txData>${f}${v}</cx:txData></cx:tx>`);
  }
  if (s.spPr) parts.push(serializeShapeProperties(s.spPr, 'cx:spPr'));
  if (s.txPr) parts.push(serializeTextBody(s.txPr, 'cx:txPr'));
  if (s.dataLabels) parts.push(serializeDataLabels(s.dataLabels));
  if (s.dataId !== undefined) parts.push(`<cx:dataId val="${s.dataId}"/>`);
  if (s.layoutPr) parts.push(serializeLayoutPr(s.layoutPr));
  if (s.axisIds) for (const a of s.axisIds) parts.push(`<cx:axisId val="${a}"/>`);
  parts.push('</cx:series>');
  return parts.join('');
};

const serializeAxis = (a: CxAxis): string => {
  const attrs: string[] = [`id="${a.id}"`];
  if (a.hidden !== undefined) attrs.push(`hidden="${a.hidden ? '1' : '0'}"`);
  const parts: string[] = [`<cx:axis ${attrs.join(' ')}>`];
  // Excel rejects an empty `<cx:axis>` — every axis needs at least a scaling
  // child. Default to `<cx:valScaling/>` when the caller didn't pick one, which
  // is the safer fallback (numeric axes are far more common than category axes
  // in chartex layouts).
  let scalingEmitted = false;
  if (a.valScaling) {
    const va: string[] = [];
    if (a.valScaling.min !== undefined) va.push(`min="${a.valScaling.min}"`);
    if (a.valScaling.max !== undefined) va.push(`max="${a.valScaling.max}"`);
    parts.push(`<cx:valScaling${va.length > 0 ? ` ${va.join(' ')}` : ''}/>`);
    scalingEmitted = true;
  }
  if (a.catScalingGapWidth !== undefined) {
    parts.push(`<cx:catScaling gapWidth="${a.catScalingGapWidth}"/>`);
    scalingEmitted = true;
  }
  if (!scalingEmitted) parts.push('<cx:valScaling/>');
  if (a.majorGridlines) parts.push('<cx:majorGridlines/>');
  if (a.title) parts.push(serializeTitle(a.title));
  if (a.spPr) parts.push(serializeShapeProperties(a.spPr, 'cx:spPr'));
  if (a.txPr) parts.push(serializeTextBody(a.txPr, 'cx:txPr'));
  parts.push('</cx:axis>');
  return parts.join('');
};

const serializeChart = (c: CxChart): string => {
  const parts: string[] = ['<cx:chart>'];
  if (c.title) parts.push(serializeTitle(c.title));
  parts.push('<cx:plotArea><cx:plotAreaRegion>');
  // <cx:plotSurface> sits at the start of <cx:plotAreaRegion> and carries the
  // plot background spPr.
  if (c.plotArea.spPr) parts.push(serializeShapeProperties(c.plotArea.spPr, 'cx:plotSurface'));
  for (const s of c.plotArea.series) parts.push(serializeSeries(s));
  parts.push('</cx:plotAreaRegion>');
  // Axes are siblings of <cx:plotAreaRegion> under <cx:plotArea>.
  for (const a of c.plotArea.axes) parts.push(serializeAxis(a));
  parts.push('</cx:plotArea>');
  if (c.legend) parts.push(serializeLegend(c.legend));
  if (c.plotVisOnly !== undefined) parts.push(`<cx:plotVisOnly val="${c.plotVisOnly ? '1' : '0'}"/>`);
  if (c.dispBlanksAs !== undefined) parts.push(`<cx:dispBlanksAs val="${c.dispBlanksAs}"/>`);
  parts.push('</cx:chart>');
  return parts.join('');
};

export function serializeChartExSpace(space: CxChartSpace): string {
  const parts: string[] = [
    XML_HEADER,
    `<cx:chartSpace xmlns:cx="${CX_NS}" xmlns:r="${REL_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`,
    serializeChartData(space.chartData),
    serializeChart(space.chart),
  ];
  // chartSpace-level spPr / txPr are siblings of <cx:chart>, emitted after.
  if (space.spPr) parts.push(serializeShapeProperties(space.spPr, 'cx:spPr'));
  if (space.txPr) parts.push(serializeTextBody(space.txPr, 'cx:txPr'));
  parts.push('</cx:chartSpace>');
  return parts.join('');
}

export function chartExToBytes(space: CxChartSpace): Uint8Array {
  return new TextEncoder().encode(serializeChartExSpace(space));
}
