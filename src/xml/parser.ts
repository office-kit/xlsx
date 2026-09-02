// DOM-style XML parser. fast-xml-parser does the lexing, then we walk its
// preserveOrder tree to:
//   1. resolve `prefix:local` element + attribute names to Clark notation
//      (`{ns}local`) using a namespace-declaration stack;
//   2. fold text segments into XmlNode.text for text-only elements;
//   3. drop XML declarations and processing instructions.
//
// DOCTYPE / external entity declarations are rejected outright via a byte-level
// prescan before the parser ever sees the input — fast-xml-parser does not
// expand external entities, but we still want the offending document to fail
// loudly.

import { XMLParser } from 'fast-xml-parser';
import { OpenXmlSchemaError } from '../utils/exceptions';
import { qname } from './namespaces';
import { el, type XmlNode } from './tree';

// ---- DOCTYPE / DTD prescan --------------------------------------------------

const decoder = new TextDecoder('utf-8', { fatal: false });

const decodeForPrescan = (input: Uint8Array | string): string => {
  if (typeof input === 'string') return input;
  return decoder.decode(input);
};

const checkForDoctype = (text: string): void => {
  // Strip XML declaration so any subsequent `<!DOCTYPE` is the real thing. The
  // declaration is always the first non-BOM token in well-formed XML.
  const stripped = text.replace(/^﻿/, '');
  if (/<!DOCTYPE\b/.test(stripped)) {
    throw new OpenXmlSchemaError('DTD declarations are not permitted in OOXML payloads');
  }
  if (/<!ENTITY\b/.test(stripped)) {
    throw new OpenXmlSchemaError('Entity declarations are not permitted in OOXML payloads');
  }
};

// ---- fast-xml-parser configuration ------------------------------------------

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  attributesGroupName: ':@',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  // OOXML needs the standard XML entities (&amp; / &lt; / &gt; / &quot; /
  // &apos;) expanded; HTML / numeric entities outside that set are not used in
  // SpreadsheetML payloads.
  processEntities: true,
  htmlEntities: false,
});

// ---- preserveOrder shape ----------------------------------------------------

type FxpAttrs = Record<string, string>;
type FxpEntry = { ':@'?: FxpAttrs } & { [tagOrText: string]: FxpEntry[] | string | FxpAttrs | undefined };
type FxpTree = FxpEntry[];

const ATTR_KEY = ':@';
const TEXT_KEY = '#text';

// ---- public API -------------------------------------------------------------

export interface ParsedDocument {
  root: XmlNode;
  /**
   * The root element's own `xmlns` declarations, in source order (`prefix` is
   * `''` for the default one). Clark notation carries the namespace but not
   * the prefix, so these are only needed by parts that must be written back
   * with the prefixes the producer chose — `xl/workbook.xml`, whose
   * `mc:Ignorable` names them by prefix.
   */
  rootNamespaces: Array<{ prefix: string; ns: string }>;
}

/**
 * Parse a UTF-8 XML payload into an {@link XmlNode} tree. Element and attribute
 * names are returned in Clark notation. Throws {@link OpenXmlSchemaError} on
 * DTD/entity declarations or on multi-root documents.
 *
 * Shorthand for {@link parseXmlDocument} when the root's namespace prefixes
 * don't matter — which is everywhere except the handful of parts that carry an
 * `mc:Ignorable`.
 */
export function parseXml(input: Uint8Array | string): XmlNode {
  return parseXmlDocument(input).root;
}

/** {@link parseXml} plus the root element's namespace declarations. */
export function parseXmlDocument(input: Uint8Array | string): ParsedDocument {
  const text = decodeForPrescan(input);
  checkForDoctype(text);

  let raw: FxpTree;
  try {
    raw = parser.parse(text) as FxpTree;
  } catch (cause) {
    throw new OpenXmlSchemaError('parseXml: failed to parse XML payload', { cause });
  }

  // Skip XML declaration, processing instructions and any leading whitespace
  // text nodes.
  const roots: FxpEntry[] = [];
  for (const entry of raw) {
    const tag = elementTag(entry);
    if (tag === undefined) continue; // text-only entry
    if (isProcessingInstruction(tag)) continue; // <?xml …?>, other PIs
    roots.push(entry);
  }
  if (roots.length === 0) {
    throw new OpenXmlSchemaError('parseXml: document has no root element');
  }
  if (roots.length > 1) {
    throw new OpenXmlSchemaError(`parseXml: document has ${roots.length} root elements; expected exactly one`);
  }

  const initial: NamespaceStack = { default: '', byPrefix: {} };
  const [root] = roots;
  if (root === undefined) {
    throw new OpenXmlSchemaError('parseXml: no root element');
  }
  return { root: convertElement(root, initial), rootNamespaces: declarationsOf(root[ATTR_KEY] as FxpAttrs | undefined) };
}

const declarationsOf = (attrs: FxpAttrs | undefined): Array<{ prefix: string; ns: string }> => {
  const out: Array<{ prefix: string; ns: string }> = [];
  for (const [k, v] of Object.entries(attrs ?? {})) {
    if (k === 'xmlns') out.push({ prefix: '', ns: v });
    else if (k.startsWith('xmlns:')) out.push({ prefix: k.slice('xmlns:'.length), ns: v });
  }
  return out;
};

// ---- conversion -------------------------------------------------------------

interface NamespaceStack {
  /** Default namespace URI (xmlns="…"); empty string means no default. */
  readonly default: string;
  /** Map of prefix → namespace URI declared in this scope or any ancestor. */
  readonly byPrefix: Readonly<Record<string, string>>;
}

const elementTag = (entry: FxpEntry): string | undefined => {
  for (const k of Object.keys(entry)) {
    if (k === ATTR_KEY) continue;
    return k;
  }
  return undefined;
};

const isProcessingInstruction = (tag: string): boolean => tag.startsWith('?');

const splitPrefixed = (qname0: string): { prefix: string; local: string } => {
  const idx = qname0.indexOf(':');
  if (idx < 0) return { prefix: '', local: qname0 };
  return { prefix: qname0.slice(0, idx), local: qname0.slice(idx + 1) };
};

const extendStack = (parent: NamespaceStack, attrs: FxpAttrs | undefined): NamespaceStack => {
  if (attrs === undefined) return parent;
  let nextDefault = parent.default;
  let nextByPrefix: Record<string, string> | undefined;
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'xmlns') {
      nextDefault = v;
      continue;
    }
    if (k.startsWith('xmlns:')) {
      const prefix = k.slice('xmlns:'.length);
      nextByPrefix ??= { ...parent.byPrefix };
      nextByPrefix[prefix] = v;
    }
  }
  if (nextDefault === parent.default && nextByPrefix === undefined) return parent;
  return {
    default: nextDefault,
    byPrefix: nextByPrefix ?? parent.byPrefix,
  };
};

const resolveElementName = (raw: string, stack: NamespaceStack): string => {
  const { prefix, local } = splitPrefixed(raw);
  if (prefix === '') return qname(stack.default, local);
  const ns = stack.byPrefix[prefix];
  if (ns === undefined) {
    throw new OpenXmlSchemaError(`parseXml: undeclared namespace prefix "${prefix}" on element <${raw}>`);
  }
  return qname(ns, local);
};

const resolveAttrName = (raw: string, stack: NamespaceStack): string => {
  const { prefix, local } = splitPrefixed(raw);
  // Unprefixed attributes do NOT inherit the default namespace (XMLNS spec).
  if (prefix === '') return local;
  if (prefix === 'xml') return qname('http://www.w3.org/XML/1998/namespace', local);
  const ns = stack.byPrefix[prefix];
  if (ns === undefined) {
    throw new OpenXmlSchemaError(`parseXml: undeclared namespace prefix "${prefix}" on attribute "${raw}"`);
  }
  return qname(ns, local);
};

const filterAttrs = (rawAttrs: FxpAttrs | undefined, stack: NamespaceStack): { resolved: Record<string, string> } => {
  const resolved: Record<string, string> = {};
  if (rawAttrs === undefined) return { resolved };
  for (const [k, v] of Object.entries(rawAttrs)) {
    // xmlns / xmlns:* declarations: dropped from the XmlNode attribute table.
    // The serializer rebuilds them from the Clark-notation namespaces it walks,
    // so round-tripping does not require preserving the declarations.
    if (k === 'xmlns' || k.startsWith('xmlns:')) continue;
    resolved[resolveAttrName(k, stack)] = v;
  }
  return { resolved };
};

const isWhitespaceOnly = (s: string): boolean => /^\s*$/.test(s);

const convertElement = (entry: FxpEntry, parentStack: NamespaceStack): XmlNode => {
  const rawTag = elementTag(entry);
  if (rawTag === undefined) {
    throw new OpenXmlSchemaError('parseXml: encountered an entry with no element tag');
  }
  if (isProcessingInstruction(rawTag)) {
    throw new OpenXmlSchemaError(`parseXml: processing instructions are not supported (saw "<${rawTag}>")`);
  }

  const rawAttrs = entry[ATTR_KEY] as FxpAttrs | undefined;
  const stack = extendStack(parentStack, rawAttrs);

  const { resolved } = filterAttrs(rawAttrs, stack);
  const node = el(resolveElementName(rawTag, stack), resolved);

  const childEntries = entry[rawTag] as FxpEntry[] | undefined;
  if (childEntries === undefined || childEntries.length === 0) return node;

  const textParts: string[] = [];
  for (const child of childEntries) {
    if (Object.hasOwn(child, TEXT_KEY)) {
      const t = child[TEXT_KEY];
      if (typeof t === 'string') textParts.push(t);
      continue;
    }
    if (textParts.length > 0 && node.children.length === 0) {
      // text accumulated *before* any child element — keep collecting.
    } else if (textParts.length > 0 && node.children.length > 0) {
      const acc = textParts.join('');
      if (!isWhitespaceOnly(acc)) {
        throw new OpenXmlSchemaError(`parseXml: mixed content not supported (text between elements under <${rawTag}>)`);
      }
      // whitespace-only inter-element text: drop.
      textParts.length = 0;
    }
    node.children.push(convertElement(child, stack));
  }
  // Trailing text after the last child element.
  if (textParts.length > 0) {
    const acc = textParts.join('');
    if (node.children.length === 0) {
      // text-only element (the common case): keep as the element's text.
      node.text = acc;
    } else if (!isWhitespaceOnly(acc)) {
      throw new OpenXmlSchemaError(`parseXml: mixed content not supported (trailing text under <${rawTag}>)`);
    }
  }
  return node;
};
