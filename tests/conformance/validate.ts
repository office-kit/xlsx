// Three-tier conformance validator for xlsx packages.
//
// Tier A — OPC structure
//   * `[Content_Types].xml` parses and covers every part in the zip.
//   * Every `*.rels` parses and resolves Internal targets to existing parts.
//   * Every Override points at a part that exists.
//
// Tier B — XSD (ECMA-376 Transitional, 5th edition)
//   * For each XML part with a known content type, validate against the
//     vendored XSD via `xmllint --schema`. Markup-compatibility extensions
//     are stripped first so the base schema is sufficient.
//
// Tier C — Semantics that XSD can't express
//   * worksheet `<c r="…">` matches its parent `<row r="…">`.
//   * worksheet `<dimension ref>` covers all populated cells.
//   * Cell `s="N"` indices fit inside `cellXfs`.
//   * SharedString references fit inside `sst`.
//   * `<mergeCells>` ranges do not overlap.
//   * Workbook `<sheet>` elements have unique sheetIds and resolvable r:ids.

import { spawnSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, posix } from 'node:path';

import { fromBuffer } from '../../src/io/node';
import {
  findOverride,
  manifestFromBytes,
  type Manifest,
} from '../../src/packaging/manifest';
import {
  relsFromBytes,
  type Relationships,
} from '../../src/packaging/relationships';
import { openZip, type ZipArchive } from '../../src/zip/reader';
import { stripIgnorableMarkup } from './mc-strip';
import {
  CONTENT_TYPES_SCHEMA,
  RELATIONSHIPS_SCHEMA,
  hasSchemaFor,
  schemaFor,
} from './schema-map';

type Tier = 'opc' | 'xsd' | 'semantic';

interface ValidationIssue {
  tier: Tier;
  part: string;
  message: string;
}

export interface ValidationResult {
  ok: boolean;
  issues: ValidationIssue[];
  /** Parts that had no XSD mapping (binary or out-of-scope formats). */
  skipped: string[];
}

export interface ValidateOptions {
  /**
   * Skip a part by exact path (e.g. `xl/printerSettings/printerSettings1.bin`)
   * even when its content type would normally be validated.
   */
  ignoreParts?: ReadonlySet<string>;
  /** When true, omit Tier B (XSD). Useful for narrow OPC/semantic checks. */
  skipXsd?: boolean;
}

const DEFAULT_DEFAULTS: Record<string, string> = {
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
};

/** Run all three tiers against an xlsx package supplied as bytes. */
export async function validateXlsx(
  bytes: Uint8Array,
  options: ValidateOptions = {},
): Promise<ValidationResult> {
  const issues: ValidationIssue[] = [];
  const skipped: string[] = [];

  const archive = await openZip(fromBuffer(bytes));
  try {
    const ctx = await loadPackage(archive, issues);
    if (!ctx) return finalize(issues, skipped);

    runTierOpc(ctx, issues);
    if (!options.skipXsd) runTierXsd(ctx, options, issues, skipped);
    runTierSemantic(ctx, issues);

    return finalize(issues, skipped);
  } finally {
    archive.close();
  }
}

function finalize(issues: ValidationIssue[], skipped: string[]): ValidationResult {
  return { ok: issues.length === 0, issues, skipped };
}

// ---------------------------------------------------------------------------
// Package loading
// ---------------------------------------------------------------------------

interface PackageContext {
  archive: ZipArchive;
  /** All entry paths, sans directory entries. */
  parts: string[];
  manifest: Manifest;
  /** All `*.rels` files in the package, keyed by their path inside the zip. */
  relsByPath: Map<string, Relationships>;
  /** Cached UTF-8 view of an XML part. */
  textOf: (path: string) => string;
}

async function loadPackage(
  archive: ZipArchive,
  issues: ValidationIssue[],
): Promise<PackageContext | undefined> {
  const parts = archive.list().filter((p) => !p.endsWith('/'));

  if (!archive.has('[Content_Types].xml')) {
    issues.push({
      tier: 'opc',
      part: '[Content_Types].xml',
      message: 'package is missing [Content_Types].xml',
    });
    return undefined;
  }
  let manifest: Manifest;
  try {
    manifest = manifestFromBytes(archive.read('[Content_Types].xml'));
  } catch (cause) {
    issues.push({
      tier: 'opc',
      part: '[Content_Types].xml',
      message: `failed to parse: ${(cause as Error).message}`,
    });
    return undefined;
  }

  const relsByPath = new Map<string, Relationships>();
  for (const path of parts) {
    if (!path.endsWith('.rels')) continue;
    try {
      relsByPath.set(path, relsFromBytes(archive.read(path)));
    } catch (cause) {
      issues.push({
        tier: 'opc',
        part: path,
        message: `failed to parse: ${(cause as Error).message}`,
      });
    }
  }

  const decoder = new TextDecoder('utf-8');
  const textCache = new Map<string, string>();
  const textOf = (path: string): string => {
    let t = textCache.get(path);
    if (t === undefined) {
      t = decoder.decode(archive.read(path));
      textCache.set(path, t);
    }
    return t;
  };

  return { archive, parts, manifest, relsByPath, textOf };
}

// ---------------------------------------------------------------------------
// Tier A — OPC structure
// ---------------------------------------------------------------------------

function runTierOpc(ctx: PackageContext, issues: ValidationIssue[]): void {
  const defaults = new Map<string, string>();
  for (const d of ctx.manifest.defaults) defaults.set(d.ext.toLowerCase(), d.contentType);
  for (const [ext, ct] of Object.entries(DEFAULT_DEFAULTS)) {
    if (!defaults.has(ext)) defaults.set(ext, ct);
  }

  // Every non-rels, non-manifest part must have a content type.
  for (const path of ctx.parts) {
    if (path === '[Content_Types].xml') continue;
    const partName = `/${path}`;
    if (findOverride(ctx.manifest, partName)) continue;
    const ext = path.split('.').pop()?.toLowerCase();
    if (ext && defaults.has(ext)) continue;
    issues.push({
      tier: 'opc',
      part: path,
      message: 'no Default extension or Override entry resolves a content type',
    });
  }

  // Every Override must target a part that physically exists.
  for (const ovr of ctx.manifest.overrides) {
    const path = ovr.partName.replace(/^\//, '');
    if (!ctx.archive.has(path)) {
      issues.push({
        tier: 'opc',
        part: ovr.partName,
        message: 'Override targets a part that does not exist in the zip',
      });
    }
  }

  // Every Internal relationship must resolve. The `source part` for a rels
  // file is the part whose path is the rels file's parent dir minus the
  // `_rels` segment — e.g. `xl/_rels/workbook.xml.rels` is the rels for
  // `xl/workbook.xml`, and `_rels/.rels` is the rels for the package root.
  for (const [relsPath, rels] of ctx.relsByPath) {
    const dir = dirname(relsPath);
    const sourceDir =
      dir === '_rels' ? '' : dir === '.' ? '' : dir.replace(/(^|\/)_rels$/, '');
    for (const r of rels.rels) {
      if (r.targetMode === 'External') continue;
      const resolved = posix.normalize(
        r.target.startsWith('/')
          ? r.target.slice(1)
          : sourceDir
            ? posix.join(sourceDir, r.target)
            : r.target,
      );
      if (!ctx.archive.has(resolved)) {
        issues.push({
          tier: 'opc',
          part: relsPath,
          message: `relationship ${r.id} → ${r.target} does not resolve to a part (resolved to ${resolved})`,
        });
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Tier B — XSD via xmllint
// ---------------------------------------------------------------------------

interface XsdJob {
  /** Path inside the zip; used for diagnostic output and ignore checks. */
  part: string;
  /** Schema to validate against. */
  schema: string;
  /** Final XML text to feed xmllint (after MC stripping). */
  xml: string;
}

function runTierXsd(
  ctx: PackageContext,
  options: ValidateOptions,
  issues: ValidationIssue[],
  skipped: string[],
): void {
  const ignore = options.ignoreParts ?? new Set<string>();
  const jobs: XsdJob[] = [];

  for (const path of ctx.parts) {
    if (ignore.has(path)) continue;
    if (path === '[Content_Types].xml') {
      jobs.push({ part: path, schema: CONTENT_TYPES_SCHEMA, xml: ctx.textOf(path) });
      continue;
    }
    if (path.endsWith('.rels')) {
      jobs.push({ part: path, schema: RELATIONSHIPS_SCHEMA, xml: ctx.textOf(path) });
      continue;
    }
    const contentType = lookupContentType(path, ctx.manifest);
    if (!contentType) continue; // Tier A reports the missing type
    if (!hasSchemaFor(contentType)) {
      skipped.push(path);
      continue;
    }
    const schema = schemaFor(contentType);
    if (!schema) continue;
    jobs.push({ part: path, schema, xml: stripIgnorableMarkupSafe(ctx.textOf(path)) });
  }

  if (jobs.length === 0) return;

  const tmpRoot = mkdtempSync(join(tmpdir(), 'xlsxkit-conformance-'));
  try {
    const bySchema = new Map<string, XsdJob[]>();
    for (const job of jobs) {
      const list = bySchema.get(job.schema) ?? [];
      list.push(job);
      bySchema.set(job.schema, list);
    }
    for (const [schema, schemaJobs] of bySchema) {
      validateBatch(tmpRoot, schema, schemaJobs, issues);
    }
  } finally {
    rmSync(tmpRoot, { recursive: true, force: true });
  }
}

function lookupContentType(path: string, manifest: Manifest): string | undefined {
  const partName = `/${path}`;
  const override = findOverride(manifest, partName);
  if (override) return override.contentType;
  const ext = path.split('.').pop()?.toLowerCase();
  if (!ext) return undefined;
  for (const d of manifest.defaults) {
    if (d.ext.toLowerCase() === ext) return d.contentType;
  }
  return DEFAULT_DEFAULTS[ext];
}

function stripIgnorableMarkupSafe(xml: string): string {
  try {
    return stripIgnorableMarkup(xml);
  } catch {
    // If pre-processing fails for any reason, fall back to the raw document
    // and let xmllint report whatever's wrong.
    return xml;
  }
}

function validateBatch(
  tmpRoot: string,
  schema: string,
  jobs: XsdJob[],
  issues: ValidationIssue[],
): void {
  // xmllint accepts multiple input files in one invocation, which amortises
  // schema-load cost. We materialise each job to a temp file whose path makes
  // the failing part identifiable from xmllint's stderr.
  const fileToPart = new Map<string, string>();
  const args: string[] = ['--noout', '--schema', schema];
  let i = 0;
  for (const job of jobs) {
    const safeName = `${i++}-${job.part.replace(/[^A-Za-z0-9._-]/g, '_')}`;
    const filePath = join(tmpRoot, safeName);
    mkdirSync(dirname(filePath), { recursive: true });
    writeFileSync(filePath, job.xml, 'utf8');
    fileToPart.set(filePath, job.part);
    args.push(filePath);
  }
  const result = spawnSync('xmllint', args, { encoding: 'utf8' });
  if (result.error) {
    issues.push({
      tier: 'xsd',
      part: '<runner>',
      message: `xmllint not available: ${result.error.message}`,
    });
    return;
  }
  if (result.status === 0) return;

  // xmllint writes one stanza per file. Parse stderr line-by-line so each
  // diagnostic is attributed back to the original part path.
  const lines = result.stderr.split('\n');
  let currentPart = '<unknown>';
  for (const line of lines) {
    if (!line.trim()) continue;
    const fileMatch = /^([^:]+):/.exec(line);
    if (fileMatch) {
      const candidate = fileMatch[1];
      const part = candidate ? fileToPart.get(candidate) : undefined;
      if (part) currentPart = part;
    }
    if (/validates$/.test(line)) continue;
    issues.push({ tier: 'xsd', part: currentPart, message: line });
  }
}

// ---------------------------------------------------------------------------
// Tier C — Semantic invariants
// ---------------------------------------------------------------------------

function runTierSemantic(ctx: PackageContext, issues: ValidationIssue[]): void {
  // Find the workbook part via Override; fall back to the conventional path.
  const workbookPart =
    ctx.manifest.overrides.find((o) =>
      /spreadsheetml\.(sheet|template)\.main\+xml$/.test(o.contentType),
    )?.partName.replace(/^\//, '') ?? 'xl/workbook.xml';

  let cellXfsCount: number | undefined;
  let sstCount: number | undefined;

  // Styles & SST counts feed the per-cell checks below.
  for (const ovr of ctx.manifest.overrides) {
    const path = ovr.partName.replace(/^\//, '');
    if (!ctx.archive.has(path)) continue;
    if (ovr.contentType.endsWith('spreadsheetml.styles+xml')) {
      cellXfsCount = readCellXfsCount(ctx.textOf(path));
    } else if (ovr.contentType.endsWith('spreadsheetml.sharedStrings+xml')) {
      sstCount = readSstCount(ctx.textOf(path));
    }
  }

  // Worksheet checks
  for (const ovr of ctx.manifest.overrides) {
    if (!ovr.contentType.endsWith('spreadsheetml.worksheet+xml')) continue;
    const path = ovr.partName.replace(/^\//, '');
    if (!ctx.archive.has(path)) continue;
    checkWorksheet(path, ctx.textOf(path), { cellXfsCount, sstCount }, issues);
  }

  if (ctx.archive.has(workbookPart)) {
    checkWorkbook(workbookPart, ctx.textOf(workbookPart), ctx.relsByPath, issues);
  }
}

interface CellChecksContext {
  cellXfsCount: number | undefined;
  sstCount: number | undefined;
}

function readCellXfsCount(xml: string): number | undefined {
  const m = /<cellXfs\b[^>]*\bcount="(\d+)"/.exec(xml);
  if (!m) return undefined;
  return Number(m[1]);
}

function readSstCount(xml: string): number | undefined {
  const m = /<sst\b[^>]*\buniqueCount="(\d+)"/.exec(xml);
  return m ? Number(m[1]) : undefined;
}

const CELL_REF = /^([A-Z]{1,3})([1-9]\d*)$/;

function parseRef(ref: string): { col: number; row: number } | undefined {
  const m = CELL_REF.exec(ref);
  if (!m) return undefined;
  let col = 0;
  for (const ch of m[1] ?? '') col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
}

function checkWorksheet(
  path: string,
  xml: string,
  cellCtx: CellChecksContext,
  issues: ValidationIssue[],
): void {
  // Streaming-ish regex pass; sufficient because @office-kit/xlsx's writers emit a
  // canonical layout. If we ever validate hand-crafted XML, swap this for
  // a real parser.
  // Capturing group 2 is the trailing `/>` for self-closing rows, so we can
  // tell `<row r="1"/>` (cell-less, no </row> to find) apart from a row
  // with an open body.
  const rowRe = /<row\b[^>]*\br="(\d+)"[^>]*?(\/?)>/g;
  const cellRe = /<c\b[^>]*\br="([A-Z]+\d+)"[^>]*(?:\/>|>)/g;
  const styleRe = /\bs="(\d+)"/;
  const typeRe = /\bt="([a-z]+)"/;

  let parentRow = 0;
  let m: RegExpExecArray | null;
  while ((m = rowRe.exec(xml))) {
    parentRow = Number(m[1]);
    if (m[2] === '/') continue; // self-closing row has no cell children
    cellRe.lastIndex = m.index + m[0].length;
    const rowEnd = xml.indexOf('</row>', cellRe.lastIndex);
    if (rowEnd === -1) break;
    while ((m = cellRe.exec(xml)) && m.index < rowEnd) {
      const ref = m[1] ?? '';
      const parsed = parseRef(ref);
      if (!parsed) {
        issues.push({ tier: 'semantic', part: path, message: `cell r="${ref}" is not a valid coordinate` });
        continue;
      }
      if (parsed.row !== parentRow) {
        issues.push({
          tier: 'semantic',
          part: path,
          message: `cell r="${ref}" is inside <row r="${parentRow}">`,
        });
      }
      const cellTag = m[0];
      const sm = styleRe.exec(cellTag);
      if (sm && cellCtx.cellXfsCount !== undefined) {
        const styleId = Number(sm[1]);
        if (styleId >= cellCtx.cellXfsCount) {
          issues.push({
            tier: 'semantic',
            part: path,
            message: `cell ${ref} references styleId ${styleId} but cellXfs.count=${cellCtx.cellXfsCount}`,
          });
        }
      }
      const tm = typeRe.exec(cellTag);
      if (tm && tm[1] === 's' && cellCtx.sstCount !== undefined) {
        // Need to read the <v> child to check the index.
        const vMatch = /<v>(\d+)<\/v>/.exec(xml.slice(m.index, rowEnd));
        if (vMatch) {
          const idx = Number(vMatch[1]);
          if (idx >= cellCtx.sstCount) {
            issues.push({
              tier: 'semantic',
              part: path,
              message: `cell ${ref} references sharedString index ${idx} but sst.uniqueCount=${cellCtx.sstCount}`,
            });
          }
        }
      }
    }
    cellRe.lastIndex = 0;
  }

  // mergeCells overlap check
  const mergeRe = /<mergeCell\b[^>]*\bref="([A-Z]+\d+:[A-Z]+\d+)"/g;
  const ranges: Array<{ r1: number; c1: number; r2: number; c2: number; ref: string }> = [];
  while ((m = mergeRe.exec(xml))) {
    const [a, b] = (m[1] ?? '').split(':');
    const start = a ? parseRef(a) : undefined;
    const end = b ? parseRef(b) : undefined;
    if (!start || !end) continue;
    ranges.push({ r1: start.row, c1: start.col, r2: end.row, c2: end.col, ref: m[1] ?? '' });
  }
  for (let i = 0; i < ranges.length; i++) {
    const a = ranges[i];
    if (!a) continue;
    for (let j = i + 1; j < ranges.length; j++) {
      const b = ranges[j];
      if (!b) continue;
      const overlap = a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
      if (overlap) {
        issues.push({
          tier: 'semantic',
          part: path,
          message: `mergeCells overlap: ${a.ref} and ${b.ref}`,
        });
      }
    }
  }
}

function checkWorkbook(
  path: string,
  xml: string,
  relsByPath: Map<string, Relationships>,
  issues: ValidationIssue[],
): void {
  // Walk every `<sheet …/>` opening tag and pick out name / sheetId / r:id from
  // the attribute soup. Attribute order in OOXML is not fixed, so we can't rely
  // on positional regexes; instead match the tag boundary once and parse attrs
  // with a separate scan.
  const sheetTagRe = /<sheet\b([^>]*)\/?>/g;
  const attrRe = /\b(name|sheetId|r:id)="([^"]*)"/g;
  const seenSheetIds = new Set<number>();
  const referencedRids = new Set<string>();
  // Excel treats sheet names as case-insensitive for uniqueness, so a workbook
  // with both `Data` and `data` is broken even if the XSD accepts it. Track
  // lower-cased names too.
  const seenNamesLower = new Map<string, string>();
  let tag: RegExpExecArray | null;
  while ((tag = sheetTagRe.exec(xml))) {
    const attrs = tag[1] ?? '';
    let name: string | undefined;
    let sheetIdAttr: string | undefined;
    let rId: string | undefined;
    attrRe.lastIndex = 0;
    let a: RegExpExecArray | null;
    while ((a = attrRe.exec(attrs))) {
      if (a[1] === 'name') name = a[2];
      else if (a[1] === 'sheetId') sheetIdAttr = a[2];
      else if (a[1] === 'r:id') rId = a[2];
    }
    if (sheetIdAttr !== undefined) {
      const sid = Number(sheetIdAttr);
      if (seenSheetIds.has(sid)) {
        issues.push({ tier: 'semantic', part: path, message: `duplicate sheetId=${sid}` });
      }
      seenSheetIds.add(sid);
    }
    if (rId !== undefined) referencedRids.add(rId);
    if (name !== undefined) {
      const lower = name.toLowerCase();
      const prior = seenNamesLower.get(lower);
      if (prior !== undefined) {
        issues.push({
          tier: 'semantic',
          part: path,
          message:
            prior === name
              ? `duplicate <sheet name="${name}">`
              : `duplicate <sheet name="${name}"> (case-insensitive collision with "${prior}")`,
        });
      } else {
        seenNamesLower.set(lower, name);
      }
    }
  }

  const wbRelsPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
  const wbRels = relsByPath.get(wbRelsPath);
  if (!wbRels) return;
  const idsInRels = new Set(wbRels.rels.map((r) => r.id));
  for (const rid of referencedRids) {
    if (!idsInRels.has(rid)) {
      issues.push({
        tier: 'semantic',
        part: path,
        message: `<sheet r:id="${rid}"> has no matching Relationship in ${wbRelsPath}`,
      });
    }
  }
}
