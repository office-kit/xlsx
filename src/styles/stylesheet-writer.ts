// Stylesheet → xl/styles.xml writer.
//
// Pairs with parseStylesheetXml. The order of sections matches what Excel emits
// and what openpyxl writes — readers tolerate any order but Excel's
// diff-friendly layout helps when comparing fixtures.

import { toTree } from '../schema/serialize';
import { qname, SHEET_MAIN_NS } from '../xml/namespaces';
import { serializeXml } from '../xml/serializer';
import { el, type XmlNode } from '../xml/tree';
import { AlignmentSchema } from './alignment.schema';
import { BorderSchema } from './borders.schema';
import { getDxfs, type StylesheetWithDxfs } from './differential';
import { fillToTree } from './fills.schema';
import { FontSchema } from './fonts.schema';
import { NumberFormatSchema } from './numbers.schema';
import { ProtectionSchema } from './protection.schema';
import type { CellXf, Stylesheet } from './stylesheet';

const STYLESHEET_TAG = qname(SHEET_MAIN_NS, 'styleSheet');
const FONTS_TAG = qname(SHEET_MAIN_NS, 'fonts');
const FILLS_TAG = qname(SHEET_MAIN_NS, 'fills');
const BORDERS_TAG = qname(SHEET_MAIN_NS, 'borders');
const NUMFMTS_TAG = qname(SHEET_MAIN_NS, 'numFmts');
const NUMFMT_TAG = qname(SHEET_MAIN_NS, 'numFmt');
const CELLSTYLEXFS_TAG = qname(SHEET_MAIN_NS, 'cellStyleXfs');
const CELLXFS_TAG = qname(SHEET_MAIN_NS, 'cellXfs');
const CELLSTYLES_TAG = qname(SHEET_MAIN_NS, 'cellStyles');
const CELLSTYLE_TAG = qname(SHEET_MAIN_NS, 'cellStyle');
const DXFS_TAG = qname(SHEET_MAIN_NS, 'dxfs');
const DXF_TAG = qname(SHEET_MAIN_NS, 'dxf');
const XF_TAG = qname(SHEET_MAIN_NS, 'xf');

/** Serialise a Stylesheet to its `xl/styles.xml` payload. */
export function stylesheetToBytes(ss: Stylesheet): Uint8Array {
  return serializeXml(buildStylesheetTree(ss));
}

function buildStylesheetTree(ss: Stylesheet): XmlNode {
  const root = el(STYLESHEET_TAG);

  // numFmts (custom only — built-ins are implicit).
  if (ss.numFmts.size > 0) {
    const numFmtsEl = el(NUMFMTS_TAG, { count: String(ss.numFmts.size) });
    const ids = [...ss.numFmts.keys()].sort((a, b) => a - b);
    for (const id of ids) {
      const code = ss.numFmts.get(id);
      if (code === undefined) continue;
      numFmtsEl.children.push(el(NUMFMT_TAG, { numFmtId: String(id), formatCode: code }));
    }
    root.children.push(numFmtsEl);
  }

  // fonts
  const fontsEl = el(FONTS_TAG, { count: String(ss.fonts.length) });
  for (const f of ss.fonts) fontsEl.children.push(toTree(f, FontSchema));
  root.children.push(fontsEl);

  // fills (use fillToTree — fillFromTree's symmetric writer)
  const fillsEl = el(FILLS_TAG, { count: String(ss.fills.length) });
  for (const fill of ss.fills) fillsEl.children.push(fillToTree(fill));
  root.children.push(fillsEl);

  // borders
  const bordersEl = el(BORDERS_TAG, { count: String(ss.borders.length) });
  for (const b of ss.borders) bordersEl.children.push(toTree(b, BorderSchema));
  root.children.push(bordersEl);

  // cellStyleXfs (always emitted — Excel rejects styles.xml that omits it when
  // cellXfs reference an xfId).
  const cellStyleXfsEl = el(CELLSTYLEXFS_TAG, {
    count: String(Math.max(ss.cellStyleXfs.length, 1)),
  });
  if (ss.cellStyleXfs.length === 0) {
    cellStyleXfsEl.children.push(el(XF_TAG, { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0' }));
  } else {
    for (const xf of ss.cellStyleXfs) cellStyleXfsEl.children.push(cellXfToTree(xf));
  }
  root.children.push(cellStyleXfsEl);

  // cellXfs — same fallback for empty pools.
  const cellXfsEl = el(CELLXFS_TAG, {
    count: String(Math.max(ss.cellXfs.length, 1)),
  });
  if (ss.cellXfs.length === 0) {
    cellXfsEl.children.push(el(XF_TAG, { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0', xfId: '0' }));
  } else {
    for (const xf of ss.cellXfs) cellXfsEl.children.push(cellXfToTree(xf));
  }
  root.children.push(cellXfsEl);

  // cellStyles — named styles entries pointing at cellStyleXfs.
  if (ss.namedStyles && ss.namedStyles.length > 0) {
    const cellStylesEl = el(CELLSTYLES_TAG, { count: String(ss.namedStyles.length) });
    for (const ns of ss.namedStyles) {
      const attrs: Record<string, string> = { name: ns.name, xfId: String(ns.xfId) };
      if (ns.builtinId !== undefined) attrs['builtinId'] = String(ns.builtinId);
      if (ns.iLevel !== undefined) attrs['iLevel'] = String(ns.iLevel);
      if (ns.hidden) attrs['hidden'] = '1';
      if (ns.customBuiltin) attrs['customBuiltin'] = '1';
      cellStylesEl.children.push(el(CELLSTYLE_TAG, attrs));
    }
    root.children.push(cellStylesEl);
  }

  // dxfs — differential styles referenced by conditional formatting and tables.
  // Emitted even when empty: Excel writes `<dxfs count="0"/>` into every
  // workbook, and a `dxfId` is an index into this list, so the element is the
  // anchor conditional-formatting rules count from.
  const dxfs = getDxfs(ss as StylesheetWithDxfs);
  {
    const dxfsEl = el(DXFS_TAG, { count: String(dxfs.length) });
    for (const dxf of dxfs) {
      const dxfEl = el(DXF_TAG);
      // Order matches openpyxl: font, numFmt, fill, alignment, border,
      // protection.
      if (dxf.font) dxfEl.children.push(toTree(dxf.font, FontSchema));
      if (dxf.numFmt) dxfEl.children.push(toTree(dxf.numFmt, NumberFormatSchema));
      if (dxf.fill) dxfEl.children.push(fillToTree(dxf.fill));
      if (dxf.alignment) dxfEl.children.push(toTree(dxf.alignment, AlignmentSchema));
      if (dxf.border) dxfEl.children.push(toTree(dxf.border, BorderSchema));
      if (dxf.protection) dxfEl.children.push(toTree(dxf.protection, ProtectionSchema));
      dxfsEl.children.push(dxfEl);
    }
    root.children.push(dxfsEl);
  }

  // tableStyles / colors / extLst, carried over verbatim from the loaded file.
  // CT_Stylesheet (ECMA-376 §18.8.39) puts all three after <dxfs>.
  if (ss.stylesXmlTail) {
    for (const node of ss.stylesXmlTail) root.children.push(node);
  }

  return root;
}

const cellXfToTree = (xf: CellXf): XmlNode => {
  const attrs: Record<string, string> = {
    numFmtId: String(xf.numFmtId),
    fontId: String(xf.fontId),
    fillId: String(xf.fillId),
    borderId: String(xf.borderId),
  };
  if (xf.xfId !== undefined) attrs['xfId'] = String(xf.xfId);
  if (xf.applyFont) attrs['applyFont'] = '1';
  if (xf.applyFill) attrs['applyFill'] = '1';
  if (xf.applyBorder) attrs['applyBorder'] = '1';
  if (xf.applyNumberFormat) attrs['applyNumberFormat'] = '1';
  if (xf.applyAlignment) attrs['applyAlignment'] = '1';
  if (xf.applyProtection) attrs['applyProtection'] = '1';
  if (xf.pivotButton) attrs['pivotButton'] = '1';
  if (xf.quotePrefix) attrs['quotePrefix'] = '1';
  const node = el(XF_TAG, attrs);
  if (xf.alignment) node.children.push(toTree(xf.alignment, AlignmentSchema));
  if (xf.protection) node.children.push(toTree(xf.protection, ProtectionSchema));
  return node;
};
