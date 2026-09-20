// Workbook-only normalization. Public XML/ZIP readers continue to report the
// original document. Opaque XML (notably themes) must cross this boundary too:
// copying Strict bytes into a Transitional package produces an invalid file.
import { SaxesParser } from 'saxes';
import { manifestFromBytes } from '../packaging/manifest.js';
import { BUILTIN_FORMATS, isDateFormat, isTimedeltaFormat } from '../styles/numbers.js';
import { parseStylesheetXml } from '../styles/stylesheet-reader.js';
import { dateToExcel } from '../utils/datetime.js';
import { escapeXmlAttr, escapeXmlText } from '../utils/escape.js';
import { OpenXmlNotImplementedError, OpenXmlSchemaError } from '../utils/exceptions.js';
import {
  ARC_CONTENT_TYPES, CHART_DRAWING_NS, CHART_NS, CUSTPROPS_NS, DRAWING_NS,
  PICTURE_NS, PKG_REL_NS, REL_NS, SHEET_DRAWING_NS, SHEET_MAIN_NS,
  STRICT_NS_ROOT, VTYPES_NS, XPROPS_NS,
} from '../xml/namespaces.js';
import type { ZipArchive } from '../zip/reader.js';

const STRICT_SHEET = `${STRICT_NS_ROOT}spreadsheetml/main`;
const STRICT_REL = `${STRICT_NS_ROOT}officeDocument/relationships`;
const NAMESPACES = new Map([
  [STRICT_SHEET, SHEET_MAIN_NS], [STRICT_REL, REL_NS],
  [`${STRICT_NS_ROOT}drawingml/main`, DRAWING_NS],
  [`${STRICT_NS_ROOT}drawingml/chart`, CHART_NS],
  [`${STRICT_NS_ROOT}drawingml/spreadsheetDrawing`, SHEET_DRAWING_NS],
  [`${STRICT_NS_ROOT}drawingml/chartDrawing`, CHART_DRAWING_NS],
  [`${STRICT_NS_ROOT}drawingml/picture`, PICTURE_NS],
  [`${STRICT_NS_ROOT}officeDocument/extendedProperties`, XPROPS_NS],
  [`${STRICT_NS_ROOT}officeDocument/customProperties`, CUSTPROPS_NS],
  [`${STRICT_NS_ROOT}officeDocument/docPropsVTypes`, VTYPES_NS],
]);
// Exact relationship types, not a rewrite of arbitrary URIs or user text.
const RELATIONSHIPS = new Map([
  'officeDocument', 'worksheet', 'chartsheet', 'sharedStrings', 'styles', 'theme',
  'calcChain', 'table', 'comments', 'drawing', 'chart', 'chartUserShapes', 'image',
  'hyperlink', 'extended-properties', 'custom-properties', 'printerSettings',
  'pivotTable', 'pivotCacheDefinition', 'pivotCacheRecords', 'externalLink',
  'externalLinkPath', 'connections', 'queryTable', 'customXml', 'customXmlProps',
].map((name) => [`${STRICT_REL}/${name}`, `${REL_NS}/${name}`]));
RELATIONSHIPS.set(`${STRICT_REL}/extendedProperties`, `${REL_NS}/extended-properties`);
RELATIONSHIPS.set(`${STRICT_REL}/customProperties`, `${REL_NS}/custom-properties`);
const STRICT_ROOTS = new Map<string, ReadonlySet<string>>([
  [SHEET_MAIN_NS, new Set(['workbook', 'worksheet', 'styleSheet', 'sst', 'comments', 'table', 'chartsheet', 'calcChain'])],
  [DRAWING_NS, new Set(['theme', 'themeOverride'])],
  [CHART_NS, new Set(['chartSpace'])],
  [SHEET_DRAWING_NS, new Set(['wsDr'])],
  [CHART_DRAWING_NS, new Set(['userShapes'])],
  [XPROPS_NS, new Set(['Properties'])], [CUSTPROPS_NS, new Set(['Properties'])],
]);
const ENCODER = new TextEncoder();
const DECODER = new TextDecoder();
const PROBE_DECODER = new TextDecoder('latin1');
const FEED_SIZE = 8192;
const ROOT_FOUND = Symbol();
const XMLNS = 'http://www.w3.org/2000/xmlns/';
const unsupported = (path: string, feature: string): never => {
  throw new OpenXmlNotImplementedError(`ISO 29500 strict: ${path}: cannot convert ${feature} to Transitional XLSX. Re-save as Excel Workbook (.xlsx).`);
};
const mapped = (table: ReadonlyMap<string, string>, value: string, path: string): string => {
  const result = table.get(value);
  if (result !== undefined) return result;
  if (value.startsWith(STRICT_NS_ROOT)) unsupported(path, value);
  return value;
};
// Preserve referenced whitespace across the second XML parse.
const attrText = (value: string): string => escapeXmlAttr(value)
  .replace(/\r/g, '&#13;').replace(/\n/g, '&#10;').replace(/\t/g, '&#9;');
const textContent = (value: string): string => escapeXmlText(value).replace(/\r/g, '&#13;');

const hasStrictMarkup = (bytes: Uint8Array): boolean => {
  const text = PROBE_DECODER.decode(bytes);
  // Numeric character references are legal in namespace declarations too.
  return text.includes(STRICT_NS_ROOT) || text.includes('&#');
};

// Time-only and elapsed-time formats do not depend on the calendar epoch.
function isCalendarFormat(code: string | undefined): boolean {
  if (!isDateFormat(code) || isTimedeltaFormat(code)) return false;
  const tokens = (code ?? '').replace(/"[^"]*"|\[[^\]]*\]|\\.|_./g, '');
  return /[dy]/i.test(tokens) || (!/[hs]/i.test(tokens) && /m/i.test(tokens));
}

interface Context {
  strict: boolean;
  date1904: boolean;
  isoEpoch: boolean;
  dateStyles: Set<number>;
}

/** Strict dates become serials, just like dates read from Transitional cells. */
function dateSerial(raw: string, context: Context, path: string): string {
  const value = raw.trim();
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(\.\d+)?(Z|[+-]\d{2}:\d{2})?)?$/.exec(value);
  if (!m && /^(?:\d{2}:\d{2}:\d{2}|[+-]?P)/.test(value)) unsupported(path, 'ISO time-only or duration cells');
  if (!m) throw new OpenXmlSchemaError(`ISO 29500 strict: ${path}: invalid ISO date cell`);
  const year = Number(m[1]), month = Number(m[2]), day = Number(m[3]);
  const hour = Number(m[4] ?? 0), minute = Number(m[5] ?? 0), second = Number(m[6] ?? 0);
  const fractionalDigits = (m[7] ?? '').slice(1);
  if (/[1-9]/.test(fractionalDigits.slice(3))) unsupported(path, 'sub-millisecond ISO dates');
  const milliseconds = Number(fractionalDigits.slice(0, 3).padEnd(3, '0'));
  const date = new Date(0);
  date.setUTCFullYear(year, month - 1, day);
  date.setUTCHours(hour, minute, second, milliseconds);
  if (month < 1 || month > 12 || day < 1 || hour > 23 || minute > 59 || second > 59 ||
      date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new OpenXmlSchemaError(`ISO 29500 strict: ${path}: invalid ISO date cell`);
  }
  const zone = m[8];
  if (zone && zone !== 'Z') {
    const hours = Number(zone.slice(1, 3)), minutes = Number(zone.slice(4, 6));
    if (hours > 14 || minutes > 59 || (hours === 14 && minutes !== 0)) {
      throw new OpenXmlSchemaError(`ISO 29500 strict: ${path}: invalid ISO date timezone`);
    }
    date.setTime(date.getTime() - (zone[0] === '+' ? 1 : -1) * (hours * 60 + minutes) * 60_000);
  }
  // The full ISO calendar and Excel's leap-day fiction cannot be represented
  // interchangeably. Do not silently change a pre-March-1900 date on save.
  if (date.getTime() < Date.UTC(context.date1904 ? 1904 : 1900, context.date1904 ? 0 : 2, 1) ||
      date.getUTCFullYear() > 9999) {
    unsupported(path, 'date outside the supported millisecond-precision Excel date range');
  }
  return String(dateToExcel(date, { epoch: context.date1904 ? 'mac' : 'windows' }));
}

function converter(path: string, context: Context, emit: (text: string) => void) {
  const parser = new SaxesParser({ xmlns: true });
  const names: string[] = [];
  let cellType = '', cellStyle = 0;
  let value: string | undefined;
  let inCell = false;
  parser.on('error', (cause) => { throw new OpenXmlSchemaError(`ISO 29500 strict: malformed XML in ${path}`, { cause }); });
  parser.on('doctype', () => { throw new OpenXmlSchemaError(`ISO 29500 strict: DOCTYPE is not allowed in ${path}`); });
  parser.on('opentag', (node) => {
    const ns = mapped(NAMESPACES, node.uri, path);
    if (names.length === 0 && node.uri.startsWith(STRICT_NS_ROOT) && !STRICT_ROOTS.get(ns)?.has(node.local)) {
      unsupported(path, `part root ${node.name}`);
    }
    if (node.uri === STRICT_SHEET) context.strict = true;
    let local = node.local;
    if (node.uri === STRICT_SHEET && names.at(-1)?.split(':').at(-1) === 'border' && (local === 'start' || local === 'end')) local = local === 'start' ? 'left' : 'right';
    const name = node.prefix ? `${node.prefix}:${local}` : local;
    names.push(name);
    if (ns === SHEET_MAIN_NS && local === 'workbookPr') {
      const compatibility = node.attributes['dateCompatibility']?.value;
      if (compatibility !== undefined && !['0', '1', 'true', 'false'].includes(compatibility)) {
        throw new OpenXmlSchemaError(`ISO 29500 strict: invalid dateCompatibility in ${path}`);
      }
      context.isoEpoch = compatibility === '0' || compatibility === 'false';
      context.date1904 = !context.isoEpoch && ['1', 'true'].includes(node.attributes['date1904']?.value ?? '');
    }
    if (ns === SHEET_MAIN_NS && local === 'c') {
      inCell = true;
      cellType = node.attributes['t']?.value ?? 'n';
      if (cellType === 'd' && node.uri !== STRICT_SHEET) cellType = '';
      cellStyle = Number(node.attributes['s']?.value ?? 0);
    }
    if (inCell && ns === SHEET_MAIN_NS && local === 'v') value = '';
    emit(`<${name}`);
    for (const attr of Object.values(node.attributes)) {
      let v = attr.value;
      if (attr.uri === XMLNS) v = mapped(NAMESPACES, v, path);
      else if (ns === PKG_REL_NS && local === 'Relationship' && attr.name === 'Type') v = mapped(RELATIONSHIPS, v, path);
      else if (ns === SHEET_MAIN_NS) {
        if (node.uri === STRICT_SHEET && local === 'workbook' && attr.name === 'conformance') continue;
        if (local === 'workbookPr' && attr.name === 'dateCompatibility') continue;
        if (local === 'workbookPr' && attr.name === 'date1904' && context.isoEpoch) v = '0';
        if (local === 'c' && attr.name === 't' && cellType === 'd') v = 'n';
        if (node.uri === STRICT_SHEET && local === 'alignment' && attr.name === 'horizontal' && (v === 'start' || v === 'end')) {
          unsupported(path, 'direction-relative cell alignment');
        }
      }
      if (node.uri.startsWith(`${STRICT_NS_ROOT}drawingml/`) &&
          /^[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:mm|cm|in|pt|pc|pi)$/.test(v)) {
        unsupported(path, 'DrawingML universal measures');
      }
      mapped(NAMESPACES, attr.uri, path);
      emit(` ${attr.name}="${attrText(v)}"`);
    }
    emit('>');
  });
  const content = (text: string): void => {
    if (value !== undefined) value += text;
    else emit(textContent(text));
  };
  parser.on('text', content);
  parser.on('cdata', content);
  parser.on('comment', (text) => emit(`<!--${text}-->`));
  parser.on('processinginstruction', ({ target, body }) => emit(`<?${target} ${body}?>`));
  parser.on('closetag', (node) => {
    if ((node.uri === STRICT_SHEET || node.uri === SHEET_MAIN_NS) && node.local === 'v' && value !== undefined) {
      if (cellType === 'd') emit(value.trim() === '' ? '' : dateSerial(value, context, path));
      else {
        const number = Number(value);
        if (context.isoEpoch && cellType === 'n' && context.dateStyles.has(cellStyle) && number < 61) {
          unsupported(path, 'numeric date before March 1900 with dateCompatibility=false');
        }
        emit(textContent(value));
      }
      value = undefined;
    }
    emit(`</${names.pop()}>`);
    if ((node.uri === STRICT_SHEET || node.uri === SHEET_MAIN_NS) && node.local === 'c') inCell = false;
  });
  return parser;
}

/** Internal adapter: no public XML or ZIP semantics are changed. */
export function normalizeStrictArchive(archive: ZipArchive): ZipArchive {
  const context: Context = { strict: false, date1904: false, isoEpoch: false, dateStyles: new Set() };
  const xmlPaths = new Set<string>();
  const workbookPaths = new Set<string>();
  const stylePaths = new Set<string>();
  if (archive.has(ARC_CONTENT_TYPES)) {
    const manifest = manifestFromBytes(archive.read(ARC_CONTENT_TYPES));
    const extensions = new Set(manifest.defaults.filter((d) => /(?:\+xml|\/xml)$/.test(d.contentType)).map((d) => d.ext));
    for (const path of archive.list()) if (extensions.has(path.slice(path.lastIndexOf('.') + 1))) xmlPaths.add(path);
    for (const part of manifest.overrides) {
      const path = part.partName.replace(/^\//, '');
      if (/(?:\+xml|\/xml)$/.test(part.contentType)) xmlPaths.add(path);
      if (/\.(?:sheet|template)(?:\.macroEnabled)?\.main\+xml$/.test(part.contentType)) workbookPaths.add(path);
      if (part.contentType.endsWith('.styles+xml')) stylePaths.add(path);
    }
  }
  const isXml = (path: string): boolean => xmlPaths.has(path) || path.endsWith('.xml') || path.endsWith('.rels');
  const read = (path: string): Uint8Array => {
    const bytes = archive.read(path);
    if (!isXml(path)) return bytes;
    // Fast-path ordinary Transitional parts, including large worksheets.
    if (!workbookPaths.has(path) && !context.isoEpoch &&
        !hasStrictMarkup(bytes)) return bytes;
    const out: string[] = [];
    converter(path, context, (text) => out.push(text)).write(DECODER.decode(bytes)).close();
    const normalized = ENCODER.encode(out.join(''));
    if (context.isoEpoch && stylePaths.has(path)) {
      const styles = parseStylesheetXml(normalized);
      context.dateStyles = new Set(styles.cellXfs.flatMap((xf, i) =>
        isCalendarFormat(styles.numFmts.get(xf.numFmtId) ?? BUILTIN_FORMATS[xf.numFmtId]) ? [i] : []));
    }
    return normalized;
  };
  return {
    list: () => archive.list(), has: (path) => archive.has(path), close: () => archive.close(),
    read, readAsync: async (path) => read(path),
    readStream(path) {
      const source = archive.readStream(path);
      if (!isXml(path)) return source;
      const decoder = new TextDecoder();
      // Probe only the root of a Transitional sheet; its rows retain the
      // original fast path. A mixed package can still contain a Strict sheet.
      let normalize: boolean | undefined = context.strict ? true : undefined;
      let pending: Uint8Array[] = [];
      const probeDecoder = new TextDecoder();
      const probe = new SaxesParser({ xmlns: true });
      probe.on('error', (cause) => { throw new OpenXmlSchemaError(`Malformed XML in ${path}`, { cause }); });
      probe.on('doctype', () => { throw new OpenXmlSchemaError(`DOCTYPE is not allowed in ${path}`); });
      probe.on('opentag', (node) => {
        normalize = node.uri.startsWith(STRICT_NS_ROOT) || Object.values(node.ns).some((ns) => ns.startsWith(STRICT_NS_ROOT));
        throw ROOT_FOUND;
      });
      let output: string[] = [];
      const parser = converter(path, context, (text) => output.push(text));
      return source.pipeThrough(new TransformStream<Uint8Array, Uint8Array>({
        transform(chunk, controller) {
          if (normalize === undefined) {
            pending.push(chunk);
            try { probe.write(probeDecoder.decode(chunk, { stream: true })); }
            catch (cause) { if (cause !== ROOT_FOUND) throw cause; }
            if (normalize === undefined) return;
          } else pending.push(chunk);
          for (const bytes of pending) {
            if (!normalize) { controller.enqueue(bytes); continue; }
            for (let i = 0; i < bytes.length; i += FEED_SIZE) {
              parser.write(decoder.decode(bytes.subarray(i, i + FEED_SIZE), { stream: true }));
              if (output.length) controller.enqueue(ENCODER.encode(output.join('')));
              output = [];
            }
          }
          pending = [];
        },
        flush(controller) {
          if (normalize === undefined) probe.write(probeDecoder.decode()).close();
          if (!normalize) return;
          parser.write(decoder.decode()).close();
          if (output.length) controller.enqueue(ENCODER.encode(output.join('')));
        },
      }));
    },
  };
}
