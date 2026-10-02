// Independent QA oracle: raw XML → OPC graph → vendored XSD → semantics.
// No production ZIP, manifest, relationship, or worksheet parser is used.
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, posix } from 'node:path';
import { required } from './required.js';
import { unzipSync } from 'fflate';
import { stripIgnorableMarkup } from './mc-strip.js';
import { CONTENT_TYPES_SCHEMA, RELATIONSHIPS_SCHEMA, schemaFor } from './schema-map.js';
import { attribute as attr, children, descendants, parseDocument, textContent, type XmlElement } from './xml-tree.js';

type Tier = 'zip' | 'xml' | 'opc' | 'xsd' | 'semantic';
interface ValidationIssue { tier: Tier; part: string; message: string }
export interface ValidationResult {
  ok: boolean;
  status: 'valid' | 'invalid' | 'incomplete' | 'inconclusive';
  issues: ValidationIssue[];
  /** XML parts without an applied schema; opaque binary parts are not included. */
  skipped: string[];
}
export interface ValidateOptions { ignoreParts?: ReadonlySet<string>; skipXsd?: boolean }
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
interface Relationship { id: string; type: string; target: string; external: boolean }
const relPath = (part: string) => part ? posix.join(posix.dirname(part), '_rels', `${posix.basename(part)}.rels`) : '_rels/.rels';
function resolve(source: string, target: string): string {
  return posix.normalize(target.startsWith('/') ? decodeURI(target.slice(1).split('#')[0] ?? '') : posix.join(posix.dirname(source), decodeURI(target.split('#')[0] ?? '')));
}
function decode(bytes: Uint8Array): string {
  const encoding = bytes[0] === 0xff && bytes[1] === 0xfe || bytes[0] === 0x3c && bytes[1] === 0 ? 'utf-16le'
    : bytes[0] === 0xfe && bytes[1] === 0xff || bytes[0] === 0 && bytes[1] === 0x3c ? 'utf-16be' : 'utf-8';
  return new TextDecoder(encoding, { fatal: true }).decode(bytes);
}
/** A skipped stage or an unsupported XML part can never yield a full pass. */
export async function validateXlsx(bytes: Uint8Array, options: ValidateOptions = {}): Promise<ValidationResult> {
  const issues: ValidationIssue[] = [];
  const skipped: string[] = [];
  const issue = (tier: Tier, part: string, message: string) => { issues.push({ tier, part, message }); };
  const finish = (): ValidationResult => {
    const status = issues.some(i => i.part !== '<runner>') ? 'invalid' : issues.length ? 'inconclusive'
      : options.skipXsd || skipped.length ? 'incomplete' : 'valid';
    return { ok: status === 'valid', status, issues, skipped };
  };
  let parts: Record<string, Uint8Array>;
  try { parts = unzipSync(bytes); }
  catch (cause) { issue('zip', '<package>', String(cause)); return finish(); }
  const types = new Map<string, string>();
  const defaults = new Map<string, string>();
  const raw = new Map<string, XmlElement>();
  const texts = new Map<string, string>();
  const parse = (part: string) => {
    try { const xml = decode(required(parts[part])); texts.set(part, xml); raw.set(part, parseDocument(xml)); }
    catch (cause) { issue('xml', part, String(cause)); }
  };
  if (!parts['[Content_Types].xml']) { issue('opc', '[Content_Types].xml', 'package is missing [Content_Types].xml'); return finish(); }
  parse('[Content_Types].xml');
  const manifest = raw.get('[Content_Types].xml');
  if (!manifest) return finish();
  if (manifest.local !== 'Types' || manifest.uri !== CT) issue('opc', '[Content_Types].xml', 'Wrong manifest root namespace');
  for (const d of children(manifest, 'Default', CT)) {
    const extension = (attr(d, 'Extension') ?? '').toLowerCase();
    if (defaults.has(extension)) issue('opc', '[Content_Types].xml', `Duplicate Default ${extension}`);
    defaults.set(extension, attr(d, 'ContentType') ?? '');
  }
  for (const o of children(manifest, 'Override', CT)) {
    const name = attr(o, 'PartName') ?? '';
    const part = name.replace(/^\//, '');
    if (!name.startsWith('/') || types.has(part)) issue('opc', '[Content_Types].xml', `Invalid or duplicate Override ${name}`);
    types.set(part, attr(o, 'ContentType') ?? '');
    if (!parts[part]) issue('opc', name, 'Override targets a part that does not exist in the zip');
  }
  for (const part of Object.keys(parts)) {
    if (part.endsWith('/') || part === '[Content_Types].xml') continue;
    if (!types.has(part)) types.set(part, defaults.get((part.split('.').at(-1) ?? '').toLowerCase()) ?? '');
    const type = types.get(part) ?? '';
    if (!type) issue('opc', part, 'no Default extension or Override entry resolves a content type');
    if (/\.(xml|rels)$/i.test(part) || /(?:\+xml|\/xml)$/.test(type)) parse(part);
  }
  const rels = new Map<string, Map<string, Relationship>>();
  for (const [part, root] of raw) {
    if (!part.endsWith('.rels')) continue;
    const source = part === '_rels/.rels' ? '' : part.replace(/(^|\/)_rels\//, '$1').replace(/\.rels$/, '');
    if (source && !parts[source]) issue('opc', part, `Relationship source ${source} does not exist`);
    if (root.uri !== REL || root.local !== 'Relationships') issue('opc', part, 'Wrong relationship root namespace');
    const map = new Map<string, Relationship>();
    for (const node of children(root, 'Relationship', REL)) {
      const r: Relationship = { id: attr(node, 'Id') ?? '', type: attr(node, 'Type') ?? '', target: attr(node, 'Target') ?? '', external: attr(node, 'TargetMode') === 'External' };
      if (map.has(r.id)) issue('opc', part, `Duplicate Relationship Id ${r.id}`);
      map.set(r.id, r);
      if (!r.external) {
        try { if (!parts[resolve(source, r.target)]) issue('opc', part, `relationship ${r.id} does not resolve to a part`); }
        catch (cause) { issue('opc', part, `Invalid target: ${String(cause)}`); }
      }
    }
    rels.set(source, map);
  }
  const office = [...(rels.get('')?.values() ?? [])].filter(r => r.type.endsWith('/officeDocument'));
  if (office.length !== 1 || office[0]?.external) issue('opc', '_rels/.rels', 'Expected exactly one internal officeDocument relationship');
  const expectedContentTypes: Record<string, string[]> = {
    styles: ['application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'],
    sharedStrings: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'],
    worksheet: ['application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'],
    drawing: ['application/vnd.openxmlformats-officedocument.drawing+xml'],
    table: ['application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'],
    chart: ['application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
  };
  for (const [source, relationships] of rels) for (const r of relationships.values()) {
    if (r.external) continue;
    const expected = expectedContentTypes[r.type.split('/').at(-1) ?? ''];
    if (expected) {
      try {
        const target = resolve(source, r.target);
        if (parts[target] && !expected.includes(types.get(target) ?? '')) issue('opc', relPath(source), `Relationship ${r.id} target has wrong content type ${types.get(target)}`);
      } catch (cause) { issue('opc', relPath(source), String(cause)); }
    }
  }
  const processed = new Map<string, XmlElement>();
  const jobs: Array<{ part: string; schema: string; xml: string }> = [];
  for (const [part] of raw) {
    try {
      const xml = part === '[Content_Types].xml' || part.endsWith('.rels') ? required(texts.get(part)) : stripIgnorableMarkup(required(texts.get(part)));
      processed.set(part, parseDocument(xml));
      const schema = part === '[Content_Types].xml' ? CONTENT_TYPES_SCHEMA : part.endsWith('.rels') ? RELATIONSHIPS_SCHEMA : schemaFor(types.get(part) ?? '');
      if (!schema || options.ignoreParts?.has(part)) skipped.push(part);
      else if (!options.skipXsd) jobs.push({ part, schema, xml });
    } catch (cause) { issue('xsd', part, `MC preprocessing failed: ${String(cause)}`); }
  }
  // Validate references after MC selection: ignored branches need not resolve.
  const expectedTypes: Record<string, string[]> = { sheet: ['worksheet', 'chartsheet', 'dialogsheet'], drawing: ['drawing'], legacyDrawing: ['vmlDrawing'], legacyDrawingHF: ['vmlDrawing'], tablePart: ['table'], hyperlink: ['hyperlink'], blip: ['image'], chart: ['chart'] };
  for (const [part, root] of processed) for (const node of descendants(root)) {
    for (const a of node.attributes) {
      if (a.uri !== R && a.uri !== 'http://purl.oclc.org/ooxml/officeDocument/relationships') continue;
      const r = rels.get(part)?.get(a.value);
      if (!r) issue('opc', part, `${node.local} ${a.local}=${a.value} has no matching Relationship in ${relPath(part)}`);
      else if (expectedTypes[node.local] && !required(expectedTypes[node.local]).includes(r.type.split('/').at(-1) ?? '')) issue('opc', part, `${node.local} refers to wrong relationship type ${r.type}`);
    }
  }
  const count = (part: string | undefined, container: string, item: string): number | undefined => {
    if (!part) return undefined;
    const root = processed.get(part);
    const node = root && (root.local === container ? root : children(root, container)[0]);
    if (!node) return 0;
    const actual = children(node, item).length;
    const declared = attr(node, container === 'sst' ? 'uniqueCount' : 'count');
    if (declared !== undefined && Number(declared) !== actual) issue('semantic', part, `${container} declared count=${declared}, actual children=${actual}`);
    return actual;
  };
  for (const [part, root] of processed) {
    if (root.uri !== SML || root.local !== 'workbook') continue;
    const names = new Map<string, string>();
    const ids = new Set<string>();
    for (const sheet of children(children(root, 'sheets')[0] ?? root, 'sheet')) {
      const name = attr(sheet, 'name') ?? '';
      const prior = names.get(name.toLowerCase());
      if (prior !== undefined) issue('semantic', part, `duplicate <sheet name="${name}"> (case-insensitive collision with "${prior}")`);
      names.set(name.toLowerCase(), name);
      const id = attr(sheet, 'sheetId') ?? '';
      if (ids.has(id)) issue('semantic', part, `duplicate sheetId=${id}`);
      ids.add(id);
    }
    const relationships = [...(rels.get(part)?.values() ?? [])];
    const target = (type: string) => { const r = relationships.find(candidate => !candidate.external && candidate.type.endsWith(`/${type}`)); return r && resolve(part, r.target); };
    const stylePart = target('styles');
    const stylesheet = stylePart && processed.get(stylePart);
    if (stylesheet && stylePart) {
      const sizes = new Map(['fonts', 'fills', 'borders', 'cellStyleXfs'].map(container => [container, count(stylePart, container, container === 'cellStyleXfs' ? 'xf' : container.slice(0, -1)) ?? 0]));
      for (const container of ['cellXfs', 'cellStyleXfs']) for (const xf of children(children(stylesheet, container)[0] ?? stylesheet, 'xf')) {
        for (const [key, pool] of [['fontId', 'fonts'], ['fillId', 'fills'], ['borderId', 'borders'], ...(container === 'cellXfs' ? [['xfId', 'cellStyleXfs']] : [])]) {
          if (!key || !pool) continue;
          const value = attr(xf, key);
          if (value !== undefined && (!Number.isInteger(Number(value)) || Number(value) < 0 || Number(value) >= (sizes.get(pool) ?? 0))) issue('semantic', stylePart, `${container} ${key}=${value} exceeds actual ${pool}`);
        }
      }
    }
    const styleCount = count(target('styles'), 'cellXfs', 'xf') ?? 1;
    const stringCount = count(target('sharedStrings'), 'sst', 'si') ?? 0;
    for (const r of relationships) {
      if (r.external || !r.type.endsWith('/worksheet')) continue;
      const path = resolve(part, r.target);
      const sheet = processed.get(path);
      if (sheet) checkWorksheet(path, sheet, styleCount, stringCount, issue);
    }
  }
  if (jobs.length) {
    const temp = mkdtempSync(join(tmpdir(), 'xlsx-conformance-'));
    try {
      const groups = new Map<string, string[]>();
      const files = new Map<string, string>();
      for (const [i, job] of jobs.entries()) {
        const file = join(temp, `${i}.xml`); writeFileSync(file, job.xml); files.set(file, job.part);
        const group = groups.get(job.schema) ?? []; group.push(file); groups.set(job.schema, group);
      }
      for (const [schema, paths] of groups) {
        const result = spawnSync('xmllint', ['--nonet', '--noout', '--schema', schema, ...paths], { encoding: 'utf8', timeout: 30_000, maxBuffer: 16 * 1024 * 1024 });
        if (result.error) { issue('xsd', '<runner>', result.error.message); continue; }
        if (result.status === 0) continue;
        const diagnostics = result.stderr.split('\n').filter(line => line.trim() && !/validates$/.test(line));
        if (!diagnostics.length) issue('xsd', '<runner>', `xmllint failed without diagnostics: exit=${result.status}, signal=${result.signal}`);
        for (const line of diagnostics) {
          const file = /^(.+?\.xml)(?::| )/.exec(line)?.[1];
          issue('xsd', file ? files.get(file) ?? '<runner>' : '<runner>', line);
        }
      }
    } finally { rmSync(temp, { recursive: true, force: true }); }
  }
  return finish();
}
interface Coordinate { row: number; col: number }
function coordinate(ref: string): Coordinate | undefined {
  const m = /^([A-Z]{1,3})([1-9]\d*)$/.exec(ref);
  if (!m) return undefined;
  let col = 0; for (const c of required(m[1])) col = col * 26 + c.charCodeAt(0) - 64;
  const row = Number(m[2]);
  return col <= 16384 && row <= 1048576 ? { row, col } : undefined;
}
function range(ref: string): [Coordinate, Coordinate] | undefined {
  const [a, b = a] = ref.split(':'); const first = coordinate(a ?? ''); const last = coordinate(b ?? '');
  return first && last && first.row <= last.row && first.col <= last.col ? [first, last] : undefined;
}
function checkWorksheet(part: string, root: XmlElement, styles: number, strings: number, issue: (tier: Tier, part: string, message: string) => void): void {
  const fail = (message: string) => issue('semantic', part, message);
  const dimension = attr(children(root, 'dimension')[0] ?? root, 'ref');
  const bounds = dimension ? range(dimension) : undefined;
  if (dimension && !bounds) fail(`Invalid dimension ${dimension}`);
  let rowIndex = 0;
  for (const row of children(children(root, 'sheetData')[0] ?? root, 'row')) {
    rowIndex = Number(attr(row, 'r') ?? rowIndex + 1);
    let col = 0;
    for (const cell of children(row, 'c')) {
      const ref = attr(cell, 'r'); const pos = ref ? coordinate(ref) : { row: rowIndex, col: col + 1 };
      if (!pos) { fail(`Invalid cell coordinate ${ref}`); continue; }
      col = pos.col;
      if (pos.row !== rowIndex) fail(`cell r="${ref}" is inside <row r="${rowIndex}">`);
      if (bounds && (pos.row < bounds[0].row || pos.row > bounds[1].row || pos.col < bounds[0].col || pos.col > bounds[1].col)) fail(`dimension ${dimension} excludes cell ${ref}`);
      const style = attr(cell, 's');
      if (style !== undefined && (!Number.isInteger(Number(style)) || Number(style) < 0 || Number(style) >= styles)) fail(`cell ${ref} references styleId ${style} but actual cellXfs=${styles}`);
      if (attr(cell, 't') === 's') {
        const v = children(cell, 'v')[0]; const index = v ? Number(textContent(v)) : NaN;
        if (!Number.isInteger(index) || index < 0 || index >= strings) fail(`cell ${ref} references sharedString index ${index} but actual sst=${strings}`);
      }
    }
  }
  // Row sweep with a fixed 16,384-column occupancy tree: O(n log(columns)).
  const events: Array<{ row: number; delta: number; a: number; b: number }> = [];
  for (const merge of children(children(root, 'mergeCells')[0] ?? root, 'mergeCell')) {
    const ref = attr(merge, 'ref') ?? ''; const r = range(ref);
    if (!r) { fail(`Invalid merge range ${ref}`); continue; }
    events.push({ row: r[0].row, delta: 1, a: r[0].col, b: r[1].col }, { row: r[1].row + 1, delta: -1, a: r[0].col, b: r[1].col });
  }
  events.sort((a, b) => a.row - b.row || a.delta - b.delta);
  const max = new Int32Array(65536); const lazy = new Int32Array(65536);
  const update = (i: number, low: number, high: number, a: number, b: number, delta: number): void => {
    if (b < low || a > high) return;
    if (a <= low && high <= b) { max[i] = (max[i] ?? 0) + delta; lazy[i] = (lazy[i] ?? 0) + delta; return; }
    const mid = (low + high) >> 1;
    update(i * 2, low, mid, a, b, delta); update(i * 2 + 1, mid + 1, high, a, b, delta);
    max[i] = (lazy[i] ?? 0) + Math.max((max[i * 2] ?? 0), (max[i * 2 + 1] ?? 0));
  };
  for (const e of events) { update(1, 1, 16384, e.a, e.b, e.delta); if ((max[1] ?? 0) > 1) { fail('mergeCells overlap'); break; } }
}
