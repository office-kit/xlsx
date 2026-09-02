// OOXML namespace and well-known package-path constants.
//
// Mirrors openpyxl/openpyxl/xml/constants.py so the test corpus stays directly
// comparable. Where openpyxl exposes a single big dict (NAMESPACES) we keep
// them as discrete `const` exports — bundlers can then drop the ones a given
// build path never imports.

// ---- W3C / Dublin Core ------------------------------------------------------

export const XML_NS = 'http://www.w3.org/XML/1998/namespace';
export const XSI_NS = 'http://www.w3.org/2001/XMLSchema-instance';

export const DCORE_NS = 'http://purl.org/dc/elements/1.1/';
export const DCTERMS_NS = 'http://purl.org/dc/terms/';
export const DCMITYPE_NS = 'http://purl.org/dc/dcmitype/';
export const DCTERMS_PREFIX = 'dcterms';

// ---- ECMA-376 Office Document core ------------------------------------------

const DOC_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/';
export const REL_NS = `${DOC_NS}relationships`;
export const COMMENTS_NS = `${REL_NS}/comments`;
export const IMAGE_NS = `${REL_NS}/image`;
export const VML_NS = `${REL_NS}/vmlDrawing`;
export const EXTERNAL_LINK_NS = `${REL_NS}/externalLink`;
export const VTYPES_NS = `${DOC_NS}docPropsVTypes`;
export const XPROPS_NS = `${DOC_NS}extended-properties`;
export const CUSTPROPS_NS = `${DOC_NS}custom-properties`;
export const CPROPS_FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';

// ---- Package layer ----------------------------------------------------------

const PKG_NS = 'http://schemas.openxmlformats.org/package/2006/';
export const PKG_REL_NS = `${PKG_NS}relationships`;
export const COREPROPS_NS = `${PKG_NS}metadata/core-properties`;
export const CONTYPES_NS = `${PKG_NS}content-types`;

// ---- SpreadsheetML & DrawingML ----------------------------------------------

export const SHEET_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const SHEET_DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
export const CHART_DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing';
export const PICTURE_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

// ---- Microsoft extensions ---------------------------------------------------

export const CUSTOMUI_NS = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility';
export const MARKUP_COMPAT_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
export const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
export const X15_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
// "Author capabilities" — the namespace of the `<x15ac:absPath>` Excel writes
// inside workbook.xml's `mc:Choice Requires="x15"`. Only needed for its
// canonical prefix, so it isn't exported.
const X15AC_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac';
export const X16_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';
export const C14_NS = 'http://schemas.microsoft.com/office/drawing/2010/chart';
export const C15_NS = 'http://schemas.microsoft.com/office/drawing/2012/chart';
export const C16_NS = 'http://schemas.microsoft.com/office/drawing/2017/03/chart';
export const CX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
export const THREADED_COMMENTS_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';

// ---- Default prefix map -----------------------------------------------------
//
// Used when serialising XmlNode trees back to text: prefer these prefixes for
// namespaces we recognise so output matches what Excel / openpyxl
// conventionally emit. Entries are intentionally short — full prefix discovery
// falls back to auto-generated names (`ns0`, `ns1`, ...) for namespaces not
// listed here.

export const DEFAULT_PREFIXES: Readonly<Record<string, string>> = Object.freeze({
  [XML_NS]: 'xml',
  [XSI_NS]: 'xsi',
  [DCORE_NS]: 'dc',
  [DCTERMS_NS]: DCTERMS_PREFIX,
  [DCMITYPE_NS]: 'dcmitype',
  [REL_NS]: 'r',
  [VTYPES_NS]: 'vt',
  // app.xml uses XPROPS_NS as its default namespace; custom.xml uses
  // CUSTPROPS_NS the same way. Mark them as '' so serialised output matches
  // Office / openpyxl convention out of the box.
  [XPROPS_NS]: '',
  [CUSTPROPS_NS]: '',
  [PKG_REL_NS]: '',
  [COREPROPS_NS]: 'cp',
  [CONTYPES_NS]: '',
  [SHEET_MAIN_NS]: '',
  [CHART_NS]: 'c',
  [DRAWING_NS]: 'a',
  [SHEET_DRAWING_NS]: 'xdr',
  [CHART_DRAWING_NS]: 'cdr',
  [PICTURE_NS]: 'pic',
  [MARKUP_COMPAT_NS]: 'mc',
  [X14_NS]: 'x14',
  [X15_NS]: 'x15',
  [X15AC_NS]: 'x15ac',
  [X16_NS]: 'x16',
  [C14_NS]: 'c14',
  [C15_NS]: 'c15',
  [C16_NS]: 'c16',
  [CX_NS]: 'cx',
  [THREADED_COMMENTS_NS]: 'tc',
});

// ---- ZIP package paths ------------------------------------------------------

export const PACKAGE_PROPS = 'docProps';
export const PACKAGE_XL = 'xl';
export const PACKAGE_RELS = '_rels';
export const PACKAGE_THEME = `${PACKAGE_XL}/theme`;
export const PACKAGE_WORKSHEETS = `${PACKAGE_XL}/worksheets`;
export const PACKAGE_CHARTSHEETS = `${PACKAGE_XL}/chartsheets`;
export const PACKAGE_DRAWINGS = `${PACKAGE_XL}/drawings`;
export const PACKAGE_CHARTS = `${PACKAGE_XL}/charts`;
export const PACKAGE_IMAGES = `${PACKAGE_XL}/media`;
export const PACKAGE_PIVOT_TABLE = `${PACKAGE_XL}/pivotTables`;
export const PACKAGE_PIVOT_CACHE = `${PACKAGE_XL}/pivotCache`;
export const PACKAGE_WORKSHEET_RELS = `${PACKAGE_WORKSHEETS}/${PACKAGE_RELS}`;
export const PACKAGE_CHARTSHEETS_RELS = `${PACKAGE_CHARTSHEETS}/${PACKAGE_RELS}`;

export const ARC_CONTENT_TYPES = '[Content_Types].xml';
export const ARC_ROOT_RELS = `${PACKAGE_RELS}/.rels`;
export const ARC_WORKBOOK_RELS = `${PACKAGE_XL}/${PACKAGE_RELS}/workbook.xml.rels`;
export const ARC_CORE = `${PACKAGE_PROPS}/core.xml`;
export const ARC_APP = `${PACKAGE_PROPS}/app.xml`;
export const ARC_CUSTOM = `${PACKAGE_PROPS}/custom.xml`;
export const ARC_WORKBOOK = `${PACKAGE_XL}/workbook.xml`;
export const ARC_STYLE = `${PACKAGE_XL}/styles.xml`;
export const ARC_THEME = `${PACKAGE_THEME}/theme1.xml`;
export const ARC_SHARED_STRINGS = `${PACKAGE_XL}/sharedStrings.xml`;
export const ARC_CUSTOM_UI = 'customUI/customUI.xml';

// ---- MIME / content-type strings -------------------------------------------

const WORKBOOK_TPL = 'application/vnd.openxmlformats-officedocument.spreadsheetml.%s.main+xml';
const SPREADSHEET_TPL = 'application/vnd.openxmlformats-officedocument.spreadsheetml.%s+xml';
const WORKBOOK_MACRO_TPL = 'application/vnd.ms-excel.%s.macroEnabled.main+xml';
const fmt = (tpl: string, kind: string): string => tpl.replace('%s', kind);

export const XLSX_TYPE = fmt(WORKBOOK_TPL, 'sheet');
export const XLTX_TYPE = fmt(WORKBOOK_TPL, 'template');
export const XLSM_TYPE = fmt(WORKBOOK_MACRO_TPL, 'sheet');
export const XLTM_TYPE = fmt(WORKBOOK_MACRO_TPL, 'template');
export const SHARED_STRINGS_TYPE = fmt(SPREADSHEET_TPL, 'sharedStrings');
export const EXTERNAL_LINK_TYPE = fmt(SPREADSHEET_TPL, 'externalLink');
export const WORKSHEET_TYPE = fmt(SPREADSHEET_TPL, 'worksheet');
export const COMMENTS_TYPE = fmt(SPREADSHEET_TPL, 'comments');
export const STYLES_TYPE = fmt(SPREADSHEET_TPL, 'styles');
export const CHARTSHEET_TYPE = fmt(SPREADSHEET_TPL, 'chartsheet');
export const DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
export const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const CHARTEX_TYPE = 'application/vnd.ms-office.chartex+xml';
export const CHARTSHAPE_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml';
export const THEME_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';
export const CPROPS_TYPE = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
export const VBA_TYPE = 'application/vnd.ms-office.vbaProject';
export const ACTIVEX_TYPE = 'application/vnd.ms-office.activeX+xml';
export const CTRLPROPS_TYPE = 'application/vnd.ms-excel.controlproperties+xml';

// ---- QName helpers ----------------------------------------------------------

/**
 * Build a Clark-notation QName: `{namespace}localname`. When `namespace` is
 * empty / undefined the local name is returned as-is.
 */
export function qname(namespace: string | undefined, local: string): string {
  return namespace ? `{${namespace}}${local}` : local;
}

const QNAME_RE = /^\{([^}]*)\}(.+)$/;

/**
 * Inverse of {@link qname}. Returns `{ ns, local }`. For unprefixed names `ns`
 * is the empty string.
 */
export function parseQName(name: string): { ns: string; local: string } {
  const m = QNAME_RE.exec(name);
  if (m === null) return { ns: '', local: name };
  return { ns: m[1] ?? '', local: m[2] ?? '' };
}
