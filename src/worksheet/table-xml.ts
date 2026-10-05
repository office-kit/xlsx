// xl/tables/tableN.xml read/write.
//
// Hand-rolled (no schema) because the table element's attribute set is large
// and conditional, and we want minimum bundle weight. Pairs with the
// loader/writer wiring in src/public/{load,save}.ts.

import { escapeXmlAttr } from '../utils/escape.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { SHEET_MAIN_NS } from '../xml/namespaces.js';
import { parseXml } from '../xml/parser.js';
import { findChild, findChildren, type XmlNode } from '../xml/tree.js';
import { parseAutoFilterNode, serializeAutoFilterXml } from './auto-filter-xml.js';
import type { TableColumn, TableDefinition, TableStyleInfo } from './table.js';
import { makeTableColumn, makeTableDefinition } from './table.js';

const TABLE_TAG = `{${SHEET_MAIN_NS}}table`;
const TABLE_COLUMNS_TAG = `{${SHEET_MAIN_NS}}tableColumns`;
const TABLE_COLUMN_TAG = `{${SHEET_MAIN_NS}}tableColumn`;
const TABLE_STYLE_INFO_TAG = `{${SHEET_MAIN_NS}}tableStyleInfo`;
const AUTOFILTER_TAG = `{${SHEET_MAIN_NS}}autoFilter`;

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeAttr = escapeXmlAttr;

const parseBool = (v: string | undefined): boolean | undefined => {
  if (v === undefined) return undefined;
  if (v === '1' || v === 'true') return true;
  if (v === '0' || v === 'false') return false;
  return undefined;
};

const parseInteger = (v: string | undefined): number | undefined => {
  if (v === undefined) return undefined;
  const n = Number.parseInt(v, 10);
  return Number.isInteger(n) ? n : undefined;
};

/** Parse a `xl/tables/tableN.xml` payload. */
export function parseTableXml(bytes: Uint8Array | string): TableDefinition {
  const root = parseXml(bytes);
  if (root.name !== TABLE_TAG) {
    throw new OpenXmlSchemaError(`parseTableXml: root is "${root.name}", expected table`);
  }
  const idRaw = root.attrs['id'];
  const displayName = root.attrs['displayName'];
  const ref = root.attrs['ref'];
  if (!idRaw || !displayName || !ref) {
    throw new OpenXmlSchemaError('parseTableXml: <table> missing required @id / @displayName / @ref');
  }
  const id = Number.parseInt(idRaw, 10);
  if (!Number.isInteger(id) || id < 1) {
    throw new OpenXmlSchemaError(`parseTableXml: <table id="${idRaw}"> is not a positive integer`);
  }

  const columns: TableColumn[] = [];
  const colsEl = findChild(root, TABLE_COLUMNS_TAG);
  if (colsEl) {
    for (const c of findChildren(colsEl, TABLE_COLUMN_TAG)) columns.push(parseTableColumn(c));
  }

  const styleEl = findChild(root, TABLE_STYLE_INFO_TAG);
  const autoFilterEl = findChild(root, AUTOFILTER_TAG);

  const totalsRowCount = parseInteger(root.attrs['totalsRowCount']);
  const headerRowCount = parseInteger(root.attrs['headerRowCount']);
  const totalsRowShown = parseBool(root.attrs['totalsRowShown']);

  const opts: Parameters<typeof makeTableDefinition>[0] = {
    id,
    displayName,
    ref,
    columns,
  };
  if (root.attrs['name']) opts.name = root.attrs['name'];
  if (headerRowCount !== undefined) opts.headerRowCount = headerRowCount;
  if (totalsRowCount !== undefined) opts.totalsRowCount = totalsRowCount;
  if (totalsRowShown !== undefined) opts.totalsRowShown = totalsRowShown;
  if (styleEl) {
    const style = parseTableStyleInfo(styleEl);
    if (style) opts.styleInfo = style;
  }
  if (autoFilterEl) {
    const af = parseAutoFilterNode(autoFilterEl);
    if (af) opts.autoFilter = af;
  }
  return makeTableDefinition(opts);
}

const parseTableColumn = (node: XmlNode): TableColumn => {
  const id = parseInteger(node.attrs['id']);
  const name = node.attrs['name'];
  if (id === undefined || !name) {
    throw new OpenXmlSchemaError('parseTableXml: <tableColumn> missing @id / @name');
  }
  const col: TableColumn = makeTableColumn({ id, name });
  const fn = node.attrs['totalsRowFunction'];
  if (fn) col.totalsRowFunction = fn as NonNullable<TableColumn['totalsRowFunction']>;
  if (node.attrs['totalsRowLabel']) col.totalsRowLabel = node.attrs['totalsRowLabel'];
  return col;
};

const parseTableStyleInfo = (node: XmlNode): TableStyleInfo | undefined => {
  const info: TableStyleInfo = {};
  if (node.attrs['name']) info.name = node.attrs['name'];
  const sFirst = parseBool(node.attrs['showFirstColumn']);
  if (sFirst !== undefined) info.showFirstColumn = sFirst;
  const sLast = parseBool(node.attrs['showLastColumn']);
  if (sLast !== undefined) info.showLastColumn = sLast;
  const sRow = parseBool(node.attrs['showRowStripes']);
  if (sRow !== undefined) info.showRowStripes = sRow;
  const sCol = parseBool(node.attrs['showColumnStripes']);
  if (sCol !== undefined) info.showColumnStripes = sCol;
  // Empty styleInfo is meaningless on the wire; collapse it.
  return Object.keys(info).length > 0 ? info : undefined;
};

/** Serialise a TableDefinition to its `xl/tables/tableN.xml` bytes. */
export function tableToBytes(table: TableDefinition): Uint8Array {
  return new TextEncoder().encode(serializeTable(table));
}

function serializeTable(table: TableDefinition): string {
  let attrs = ` id="${table.id}"`;
  if (table.name !== undefined) attrs += ` name="${escapeAttr(table.name)}"`;
  attrs += ` displayName="${escapeAttr(table.displayName)}"`;
  attrs += ` ref="${escapeAttr(table.ref)}"`;
  if (table.headerRowCount !== undefined) attrs += ` headerRowCount="${table.headerRowCount}"`;
  if (table.totalsRowCount !== undefined) attrs += ` totalsRowCount="${table.totalsRowCount}"`;
  if (table.totalsRowShown !== undefined) attrs += ` totalsRowShown="${table.totalsRowShown ? '1' : '0'}"`;

  const parts: string[] = [XML_HEADER, `<table xmlns="${SHEET_MAIN_NS}"${attrs}>`];
  if (table.autoFilter) parts.push(serializeAutoFilterXml(table.autoFilter));
  parts.push(`<tableColumns count="${table.columns.length}">`);
  for (const col of table.columns) parts.push(serializeTableColumn(col));
  parts.push('</tableColumns>');
  if (table.styleInfo) parts.push(serializeTableStyleInfo(table.styleInfo));
  parts.push('</table>');
  return parts.join('');
}

const serializeTableColumn = (col: TableColumn): string => {
  let attrs = ` id="${col.id}" name="${escapeAttr(col.name)}"`;
  if (col.totalsRowFunction) attrs += ` totalsRowFunction="${col.totalsRowFunction}"`;
  if (col.totalsRowLabel !== undefined) attrs += ` totalsRowLabel="${escapeAttr(col.totalsRowLabel)}"`;
  return `<tableColumn${attrs}/>`;
};

const serializeTableStyleInfo = (info: TableStyleInfo): string => {
  let attrs = '';
  if (info.name !== undefined) attrs += ` name="${escapeAttr(info.name)}"`;
  if (info.showFirstColumn !== undefined) attrs += ` showFirstColumn="${info.showFirstColumn ? '1' : '0'}"`;
  if (info.showLastColumn !== undefined) attrs += ` showLastColumn="${info.showLastColumn ? '1' : '0'}"`;
  if (info.showRowStripes !== undefined) attrs += ` showRowStripes="${info.showRowStripes ? '1' : '0'}"`;
  if (info.showColumnStripes !== undefined) attrs += ` showColumnStripes="${info.showColumnStripes ? '1' : '0'}"`;
  return `<tableStyleInfo${attrs}/>`;
};

