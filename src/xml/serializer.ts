// DOM-style XML serializer. Inverse of src/xml/parser.ts.
//
// XmlNode trees use Clark-notation names (`{ns}local`); the serializer
// rebuilds prefix mappings and emits a single UTF-8 byte payload.
//
// Algorithm:
//   1. walk the tree once, collecting every namespace URI used by
//      elements or attributes;
//   2. assign each URI a prefix — DEFAULT_PREFIXES first, then
//      auto-generated `ns{N}` — and decide which (if any) becomes the
//      default (xmlns="…") namespace by reusing the root element's NS;
//   3. emit the XML declaration, root open tag with all xmlns / xmlns:*
//      declarations, then a recursive children walk that escapes text
//      and attribute values.
//
// Output style matches openpyxl/Excel: no whitespace between elements,
// `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` on its own
// line, attribute values quoted with `"`.

import { escapeXmlAttr, escapeXmlText } from '../utils/escape';
import { DEFAULT_PREFIXES, parseQName, XML_NS } from './namespaces';
import type { XmlNode } from './tree';

export interface SerializeOptions {
  /** Emit `<?xml … ?>` declaration. Defaults to true. */
  xmlDeclaration?: boolean;
  /** `standalone` attribute on the declaration. Defaults to 'yes'. */
  standalone?: 'yes' | 'no' | 'omit';
  /**
   * Namespace URI → prefix already bound by an ancestor, for a fragment being
   * spliced into a larger document. Those bindings are reused and not
   * redeclared on the fragment's own root. `''` means the ancestor's default
   * namespace — attributes never inherit a default binding, so only use it for
   * a namespace whose attributes are unqualified (every SpreadsheetML part).
   */
  inScopePrefixes?: Readonly<Record<string, string>>;
}

/**
 * Serialize an {@link XmlNode} tree into a UTF-8 byte payload. The
 * inverse of {@link parseXml}: `parseXml(serializeXml(n))` yields a
 * tree structurally equivalent to `n` modulo attribute insertion order
 * (which both directions preserve on Node 18+ / V8).
 */
export function serializeXml(root: XmlNode, opts: SerializeOptions = {}): Uint8Array {
  const { xmlDeclaration = true, standalone = 'yes', inScopePrefixes } = opts;

  const allocation = allocatePrefixes(root, inScopePrefixes);
  const out: string[] = [];

  if (xmlDeclaration) {
    out.push('<?xml version="1.0" encoding="UTF-8"');
    if (standalone !== 'omit') out.push(` standalone="${standalone}"`);
    out.push('?>\n');
  }

  emit(out, root, allocation, /* isRoot */ true);

  return new TextEncoder().encode(out.join(''));
}

// ---- prefix allocation ------------------------------------------------------

interface Allocation {
  /** ns URI → prefix ('' = default namespace) */
  prefixOf: Map<string, string>;
  /** prefix declarations to emit on the root element, in deterministic order */
  declarations: Array<{ prefix: string; ns: string }>;
}

const allocatePrefixes = (root: XmlNode, inScope?: Readonly<Record<string, string>>): Allocation => {
  const used = new Set<string>();
  collectNamespaces(root, used, /* attrsToo */ true);
  used.delete(''); // empty NS = unprefixed names; not emitted as xmlns

  const prefixOf = new Map<string, string>();
  // Bindings an ancestor already declared: reuse them, and leave them out of
  // this fragment's own declarations.
  const inherited = new Set<string>();
  for (const [ns, prefix] of Object.entries(inScope ?? {})) {
    if (ns === '' || prefixOf.has(ns)) continue;
    prefixOf.set(ns, prefix);
    inherited.add(ns);
  }

  // Decide the default namespace. Prefer the root element's NS if its
  // canonical prefix in DEFAULT_PREFIXES is empty (i.e. designed to live
  // as a default), so the bulk of the root subtree stays unprefixed — unless
  // an ancestor already bound that namespace, in which case its binding wins.
  const rootNs = parseQName(root.name).ns;
  let defaultNs = '';
  if (rootNs && DEFAULT_PREFIXES[rootNs] === '' && !prefixOf.has(rootNs)) {
    defaultNs = rootNs;
    prefixOf.set(defaultNs, '');
  }
  // `xml` and `xmlns` are reserved by the XMLNS spec; xml ↔ XML_NS is
  // predefined and must NOT be redeclared via xmlns:xml="…".
  prefixOf.set(XML_NS, 'xml');

  // Pass 1: respect DEFAULT_PREFIXES for namespaces that have a
  // canonical short prefix.
  let auto = 0;
  const usedPrefixes = new Set<string>(['', 'xml', 'xmlns', ...prefixOf.values()]);
  const ordered = Array.from(used).sort();
  for (const ns of ordered) {
    if (prefixOf.has(ns)) continue;
    const preferred = DEFAULT_PREFIXES[ns];
    if (preferred !== undefined && preferred !== '' && !usedPrefixes.has(preferred)) {
      prefixOf.set(ns, preferred);
      usedPrefixes.add(preferred);
    }
  }
  // Pass 2: auto-allocate for the remainder.
  for (const ns of ordered) {
    if (prefixOf.has(ns)) continue;
    let prefix = `ns${auto++}`;
    while (usedPrefixes.has(prefix)) prefix = `ns${auto++}`;
    prefixOf.set(ns, prefix);
    usedPrefixes.add(prefix);
  }

  // Emission order: default first, then other prefixes alphabetically by ns
  // URI for determinism. XML_NS is reserved and never declared.
  const declarations: Array<{ prefix: string; ns: string }> = [];
  if (defaultNs !== '') declarations.push({ prefix: '', ns: defaultNs });
  const inheritedDefault = [...inherited].find((ns) => prefixOf.get(ns) === '');
  for (const ns of ordered) {
    if (ns === defaultNs) continue;
    if (ns === inheritedDefault) continue;
    if (ns === XML_NS) continue;
    if (inherited.has(ns)) continue;
    const prefix = prefixOf.get(ns);
    if (prefix === undefined) continue;
    declarations.push({ prefix, ns });
  }

  return { prefixOf, declarations };
};

const collectNamespaces = (node: XmlNode, into: Set<string>, attrsToo: boolean): void => {
  into.add(parseQName(node.name).ns);
  if (attrsToo) {
    for (const attrName of Object.keys(node.attrs)) {
      into.add(parseQName(attrName).ns);
    }
  }
  for (const c of node.children) collectNamespaces(c, into, attrsToo);
};

// ---- emission ---------------------------------------------------------------

const buildElementPrefix = (name: string, prefixOf: Map<string, string>): string => {
  const { ns, local } = parseQName(name);
  if (ns === '') return local;
  const prefix = prefixOf.get(ns);
  if (prefix === undefined || prefix === '') return local;
  return `${prefix}:${local}`;
};

const buildAttrPrefix = (name: string, prefixOf: Map<string, string>): string => {
  const { ns, local } = parseQName(name);
  if (ns === '') return local;
  // Attributes never inherit the default namespace; if the namespace is
  // mapped to the default prefix, they still need an explicit prefix.
  const prefix = prefixOf.get(ns);
  if (prefix === undefined || prefix === '') {
    // Attribute lives in a namespace that we marked as default-only;
    // fall through to no prefix (the namespace must already be declared
    // by ancestor walk, but for default-namespace attrs that's invalid
    // XML — we keep it bare anyway since OOXML never produces this).
    return local;
  }
  return `${prefix}:${local}`;
};

const emit = (out: string[], node: XmlNode, allocation: Allocation, isRoot: boolean): void => {
  const tag = buildElementPrefix(node.name, allocation.prefixOf);
  out.push('<', tag);

  if (isRoot) {
    for (const { prefix, ns } of allocation.declarations) {
      out.push(prefix === '' ? ` xmlns="${escapeXmlAttr(ns)}"` : ` xmlns:${prefix}="${escapeXmlAttr(ns)}"`);
    }
  }

  for (const [name, value] of Object.entries(node.attrs)) {
    const attrName = buildAttrPrefix(name, allocation.prefixOf);
    out.push(' ', attrName, '="', escapeXmlAttr(value), '"');
  }

  const text = node.text;
  const hasText = text !== undefined && text !== '';
  const hasChildren = node.children.length > 0;
  if (!hasText && !hasChildren) {
    out.push('/>');
    return;
  }

  out.push('>');
  if (hasText) out.push(escapeXmlText(text));
  for (const c of node.children) emit(out, c, allocation, /* isRoot */ false);
  out.push('</', tag, '>');
};
