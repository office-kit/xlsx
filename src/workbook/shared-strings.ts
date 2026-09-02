// Shared-strings table read/write.
//
// Excel pulls every plain-string cell out of the sheet bodies and into
// `xl/sharedStrings.xml` so duplicates compress well. The format is a flat list
// of `<si>` entries, each holding either a single `<t>` (plain text) or a
// sequence of `<r><rPr/>?<t/></r>` runs (rich text).
//
// Rich-text entries are kept as their full `RichText` runs so per-run fonts
// (bold / italic / colour / size / …) survive the round-trip. Plain strings
// dedup against literal text; rich text dedups against the structural shape of
// its runs, so a string that is bold in one cell and plain in another still
// gets two slots.

import type { RichText } from '../cell/rich-text';
import { type Color, colorToHex } from '../styles/colors';
import { escapeCellString, escapeXmlAttr, escapeXmlText, unescapeCellString } from '../utils/escape';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { stableStringify } from '../utils/stable-stringify';
import { qname, SHEET_MAIN_NS } from '../xml/namespaces';
import { parseXml } from '../xml/parser';
import { findChild, findChildren, type XmlNode } from '../xml/tree';

const SST_TAG = `{${SHEET_MAIN_NS}}sst`;
const SI_TAG = `{${SHEET_MAIN_NS}}si`;
const T_TAG = `{${SHEET_MAIN_NS}}t`;
const R_TAG = `{${SHEET_MAIN_NS}}r`;

/** A single SST entry: either a plain string or a rich-text run array. */
export type SharedStringEntry = string | { kind: 'rich-text'; runs: RichText };

/**
 * Mutable shared-strings accumulator + lookup table. The same shape is used
 * during read (just populate `entries`) and write (call `addSharedString` from
 * the worksheet writer; emit to bytes at the end).
 */
export interface SharedStringsTable {
  /** Insertion-ordered list of unique entries. */
  entries: SharedStringEntry[];
  /** Reverse lookup keyed by literal text — rich-text entries skip this map. */
  index: Map<string, number>;
  /** Reverse lookup for rich-text entries, keyed by the structure of their runs. */
  richIndex: Map<string, number>;
}

export function makeSharedStrings(): SharedStringsTable {
  return { entries: [], index: new Map(), richIndex: new Map() };
}

/**
 * Insert a string and return its index. Idempotent: calling with the same value
 * twice gives the same index. Empty strings are deduped just like everything
 * else.
 */
export function addSharedString(table: SharedStringsTable, value: string): number {
  const cached = table.index.get(value);
  if (cached !== undefined) return cached;
  const id = table.entries.length;
  table.entries.push(value);
  table.index.set(value, id);
  return id;
}

/**
 * Insert a rich-text entry and return its index. Deduped on the structure of
 * the runs, so the same formatted string used in 100 cells lands in one slot —
 * which is how Excel stores it.
 */
export function addSharedRichText(table: SharedStringsTable, runs: RichText): number {
  const key = stableStringify(runs);
  const cached = table.richIndex.get(key);
  if (cached !== undefined) return cached;
  const id = table.entries.length;
  table.entries.push({ kind: 'rich-text', runs });
  table.richIndex.set(key, id);
  return id;
}

/** Look up a shared-string index by its literal text. Returns `undefined` for unknown values. */
export function getSharedStringIndex(table: SharedStringsTable, value: string): number | undefined {
  return table.index.get(value);
}

/**
 * Read a shared-string by its 0-based index. Returns `undefined` for
 * out-of-range. Rich-text entries surface their concatenated plain text so
 * callers that want only the textual body don't need to know about the
 * discriminated union.
 */
export function getSharedStringAt(table: SharedStringsTable, index: number): string | undefined {
  const entry = table.entries[index];
  if (entry === undefined) return undefined;
  if (typeof entry === 'string') return entry;
  return entry.runs.map((r) => r.text).join('');
}

/** Number of unique entries in the SST. */
export function sharedStringCount(table: SharedStringsTable): number {
  return table.entries.length;
}

// ---- read ------------------------------------------------------------------

/** Concatenate every `<t>` text node found inside an arbitrary XmlNode tree. */
const collectText = (node: XmlNode): string => {
  // Most common case: a direct `<si><t>x</t></si>` — bypass the recursion.
  if (node.children.length === 1) {
    const only = node.children[0];
    if (only && only.name === T_TAG) return unescapeCellString(only.text ?? '');
  }
  let out = '';
  for (const child of node.children) {
    if (child.name === T_TAG) {
      out += child.text ?? '';
    } else if (child.name === R_TAG) {
      const t = findChild(child, T_TAG);
      if (t?.text) out += t.text;
    }
  }
  return unescapeCellString(out);
};

/**
 * Parse a `xl/sharedStrings.xml` payload. Returns the table directly (rather
 * than just the array) so the worksheet writer can keep appending to it without
 * rebuilding the index.
 *
 * Rich-text runs are preserved as their full per-run formatting so
 * round-tripping a file with rich text doesn't drop the styling.
 */
export function parseSharedStringsXml(bytes: Uint8Array | string): SharedStringsTable {
  const root = parseXml(bytes);
  if (root.name !== SST_TAG) {
    throw new OpenXmlSchemaError(`parseSharedStringsXml: root is "${root.name}", expected sst`);
  }
  const table = makeSharedStrings();
  for (const si of findChildren(root, SI_TAG)) {
    const entry = parseRichString(si);
    // Don't dedup — Excel preserves duplicate `<si>` entries by index, and the
    // worksheet `t="s"` references depend on slot, not on text equality.
    const id = table.entries.length;
    table.entries.push(entry);
    if (typeof entry === 'string') {
      if (!table.index.has(entry)) table.index.set(entry, id);
    } else {
      const key = stableStringify(entry.runs);
      if (!table.richIndex.has(key)) table.richIndex.set(key, id);
    }
  }
  return table;
}

/**
 * Parse a `CT_Rst` body — the shape shared by `<si>` in sharedStrings.xml and
 * `<is>` in an inline-string cell. Returns rich text when the body is built
 * from `<r>` runs, and the plain concatenated text otherwise.
 */
export const parseRichString = (si: XmlNode): SharedStringEntry => {
  // Rich-text si has one or more <r> children. Plain si has a single <t>.
  const runEls = findChildren(si, R_TAG);
  if (runEls.length > 0) {
    const runs: Array<{ text: string; font?: import('../cell/rich-text').InlineFont }> = [];
    for (const rEl of runEls) {
      const tEl = findChild(rEl, T_TAG);
      const text = unescapeCellString(tEl?.text ?? '');
      const rPrEl = findChild(rEl, qname(SHEET_MAIN_NS, 'rPr'));
      const font = rPrEl ? parseRunPr(rPrEl) : undefined;
      runs.push(font !== undefined ? { text, font } : { text });
    }
    return { kind: 'rich-text', runs: Object.freeze(runs) };
  }
  return collectText(si);
};

const parseRunPr = (rPr: XmlNode): import('../cell/rich-text').InlineFont | undefined => {
  type InlineFontMutable = {
    -readonly [K in keyof import('../cell/rich-text').InlineFont]: import('../cell/rich-text').InlineFont[K];
  };
  const f: InlineFontMutable = {};
  for (const child of rPr.children) {
    const local = child.name.replace(/^\{[^}]+\}/, '');
    const valAttr = child.attrs['val'];
    switch (local) {
      case 'rFont':
      case 'name':
        if (valAttr !== undefined) f.name = valAttr;
        break;
      case 'sz':
        if (valAttr !== undefined) f.sz = Number.parseFloat(valAttr);
        break;
      case 'b':
        f.b = valAttr === undefined ? true : valAttr !== '0' && valAttr !== 'false';
        break;
      case 'i':
        f.i = valAttr === undefined ? true : valAttr !== '0' && valAttr !== 'false';
        break;
      case 'u': {
        const v = (valAttr ?? 'single') as import('../cell/rich-text').InlineUnderline;
        f.u = v;
        break;
      }
      case 'strike':
        f.strike = valAttr === undefined ? true : valAttr !== '0' && valAttr !== 'false';
        break;
      case 'vertAlign':
        if (valAttr !== undefined) f.vertAlign = valAttr as import('../cell/rich-text').InlineVertAlign;
        break;
      case 'family':
        if (valAttr !== undefined) f.family = Number.parseInt(valAttr, 10);
        break;
      case 'charset':
        if (valAttr !== undefined) f.charset = Number.parseInt(valAttr, 10);
        break;
      case 'scheme':
        if (valAttr !== undefined) f.scheme = valAttr as 'major' | 'minor';
        break;
      case 'color': {
        const c: { rgb?: string; theme?: number; indexed?: number; tint?: number; auto?: boolean } = {};
        if (child.attrs['rgb'] !== undefined) c.rgb = child.attrs['rgb'];
        if (child.attrs['theme'] !== undefined) c.theme = Number.parseInt(child.attrs['theme'], 10);
        if (child.attrs['indexed'] !== undefined) c.indexed = Number.parseInt(child.attrs['indexed'], 10);
        if (child.attrs['tint'] !== undefined) c.tint = Number.parseFloat(child.attrs['tint']);
        if (child.attrs['auto'] !== undefined) c.auto = child.attrs['auto'] === '1' || child.attrs['auto'] === 'true';
        f.color = c as Color;
        break;
      }
    }
  }
  return Object.keys(f).length === 0 ? undefined : Object.freeze(f as import('../cell/rich-text').InlineFont);
};

// ---- write -----------------------------------------------------------------

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Serialise a SharedStringsTable to its OOXML bytes. The `count` attribute
 * tracks total references (we don't know that here so we report the same as
 * `uniqueCount` — readers tolerate the discrepancy and Excel ignores `count` in
 * practice). `uniqueCount` always matches `entries.length`.
 */
export function sharedStringsToBytes(table: SharedStringsTable): Uint8Array {
  return new TextEncoder().encode(serializeSharedStrings(table));
}

export function serializeSharedStrings(table: SharedStringsTable): string {
  const total = table.entries.length;
  const parts: string[] = [XML_HEADER, `<sst xmlns="${SHEET_MAIN_NS}" count="${total}" uniqueCount="${total}">`];
  for (const value of table.entries) {
    parts.push(serializeSi(value));
  }
  parts.push('</sst>');
  return parts.join('');
}

const serializeSi = (value: SharedStringEntry): string => {
  if (typeof value === 'string') {
    // Whitespace at either end needs xml:space="preserve" so Excel doesn't
    // collapse it. Mirrors openpyxl's emitter.
    const preserve = value.length > 0 && (value[0] === ' ' || value[value.length - 1] === ' ' || /[\t\n]/.test(value));
    const tAttr = preserve ? ' xml:space="preserve"' : '';
    return `<si><t${tAttr}>${escapeXmlText(escapeCellString(value))}</t></si>`;
  }
  return `<si>${serializeRichTextRuns(value.runs)}</si>`;
};

/** Serialise a sequence of `<r>…</r>` runs into an `<si>` / `<is>` body. */
function serializeRichTextRuns(runs: import('../cell/rich-text').RichText): string {
  const parts: string[] = [];
  for (const run of runs) {
    parts.push('<r>');
    if (run.font) parts.push(serializeInlineFont(run.font));
    const text = run.text;
    const preserve = text.length > 0 && (text[0] === ' ' || text[text.length - 1] === ' ' || /[\t\n]/.test(text));
    const tAttr = preserve ? ' xml:space="preserve"' : '';
    parts.push(`<t${tAttr}>${escapeXmlText(escapeCellString(text))}</t>`);
    parts.push('</r>');
  }
  return parts.join('');
}

const serializeInlineFont = (f: import('../cell/rich-text').InlineFont): string => {
  // Element order per ECMA-376 §17.4.4.10 (CT_RPrElt). Excel's parser is
  // sensitive to ordering — out-of-order children make the run silently fall
  // back to the cell's font.
  const parts: string[] = ['<rPr>'];
  if (f.name !== undefined) parts.push(`<rFont val="${escapeXmlAttr(f.name)}"/>`);
  if (f.charset !== undefined) parts.push(`<charset val="${f.charset}"/>`);
  if (f.family !== undefined) parts.push(`<family val="${f.family}"/>`);
  if (f.b) parts.push('<b/>');
  if (f.i) parts.push('<i/>');
  if (f.strike) parts.push('<strike/>');
  if (f.outline) parts.push('<outline/>');
  if (f.shadow) parts.push('<shadow/>');
  if (f.condense) parts.push('<condense/>');
  if (f.extend) parts.push('<extend/>');
  if (f.color) parts.push(serializeRunColor(f.color));
  if (f.sz !== undefined) parts.push(`<sz val="${f.sz}"/>`);
  if (f.u) parts.push(`<u val="${f.u}"/>`);
  if (f.vertAlign) parts.push(`<vertAlign val="${f.vertAlign}"/>`);
  if (f.scheme) parts.push(`<scheme val="${f.scheme}"/>`);
  parts.push('</rPr>');
  return parts.join('');
};

const serializeRunColor = (c: Color): string => {
  const attrs: string[] = [];
  if (c.rgb !== undefined) attrs.push(`rgb="${escapeXmlAttr(c.rgb)}"`);
  else if (c.theme !== undefined) attrs.push(`theme="${c.theme}"`);
  else if (c.indexed !== undefined) attrs.push(`indexed="${c.indexed}"`);
  else if (c.auto !== undefined) attrs.push(`auto="${c.auto ? '1' : '0'}"`);
  if (c.tint !== undefined) attrs.push(`tint="${c.tint}"`);
  // Fallback to colorToHex if nothing was set above (shouldn't happen, but
  // safe).
  if (attrs.length === 0) {
    const hex = colorToHex(c);
    if (hex !== undefined) attrs.push(`rgb="${hex}"`);
  }
  return `<color${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}/>`;
};
