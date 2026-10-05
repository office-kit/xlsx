// DOM-style XML parser. lex.ts does the lexing, then we walk its tree to:
//   1. resolve `prefix:local` element + attribute names to Clark notation
//      (`{ns}local`) using a namespace-declaration stack;
//   2. fold text segments into XmlNode.text for text-only elements;
//   3. drop XML declarations and processing instructions.
//
// DOCTYPE / external entity declarations are rejected outright via a byte-level
// prescan before the lexer ever sees the input. The lexer does not expand
// entities, but we still want the offending document to fail loudly.

import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { ATTR_KEY, CDATA_KEY, lexXml, type RawAttrs, type RawEntry, TEXT_KEY } from './lex.js';
import { MARKUP_COMPAT_NS, qname } from './namespaces.js';
import { el, type XmlNode } from './tree.js';

// ---- DOCTYPE / DTD prescan --------------------------------------------------

const decoder = new TextDecoder('utf-8', { fatal: false });

const decodeForPrescan = (input: Uint8Array | string): string => {
  if (typeof input === 'string') return input;
  return decoder.decode(input);
};

/**
 * Refuse a payload carrying a DTD. Exported because the worksheet reader lexes
 * `<sheetData>` itself and so has to run this over the span `parseXml` no
 * longer sees.
 */
export const rejectDtdDeclarations = (text: string): void => {
  // Drop a leading byte-order mark so the scan below starts at the markup. The
  // XML declaration is left in place: neither pattern can match inside one, and
  // `<!DOCTYPE` / `<!ENTITY` anywhere in the payload is refused whether or not
  // it sits where a DTD legally could.
  const stripped = text.replace(/^﻿/, '');
  if (/<!DOCTYPE\b/.test(stripped)) {
    throw new OpenXmlSchemaError('DTD declarations are not permitted in OOXML payloads');
  }
  if (/<!ENTITY\b/.test(stripped)) {
    throw new OpenXmlSchemaError('Entity declarations are not permitted in OOXML payloads');
  }
};

// ---- entities ---------------------------------------------------------------

// The lexer leaves references as written; they are decoded here in one pass,
// so `&amp;#65;` is the literal text "&#65;", not "A". CDATA is never decoded.
const XML_ENTITY_RE = /&(amp|lt|gt|quot|apos|#(?:[0-9]+|x[0-9A-Fa-f]+));/g;

/** Decode the five predefined XML entities and decimal/hex character refs once. */
const decodeXmlEntities = (value: string): string =>
  value.replace(XML_ENTITY_RE, (reference, body: string) => {
    switch (body) {
      case 'amp':
        return '&';
      case 'lt':
        return '<';
      case 'gt':
        return '>';
      case 'quot':
        return '"';
      case 'apos':
        return "'";
      default: {
        const radix = body[1] === 'x' ? 16 : 10;
        const digits = radix === 16 ? body.slice(2) : body.slice(1);
        const codePoint = Number.parseInt(digits, radix);
        if (!isXmlCodePoint(codePoint)) {
          throw new OpenXmlSchemaError(`parseXml: invalid XML character reference "${reference}"`);
        }
        return String.fromCodePoint(codePoint);
      }
    }
  });

const isXmlCodePoint = (codePoint: number): boolean =>
  codePoint === 0x09 ||
  codePoint === 0x0a ||
  codePoint === 0x0d ||
  (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
  (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
  (codePoint >= 0x10000 && codePoint <= 0x10ffff);

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
  rejectDtdDeclarations(text);

  let raw: RawEntry[];
  try {
    raw = lexXml(text);
  } catch (cause) {
    throw new OpenXmlSchemaError('parseXml: failed to parse XML payload', { cause });
  }

  // Skip XML declaration, processing instructions and any leading whitespace
  // text nodes.
  const roots: RawEntry[] = [];
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
  if (Object.hasOwn(root, TEXT_KEY)) {
    // `x</q>`: text before a stray closing tag is the only top-level entry.
    // Converting it as an element used to recurse until a raw RangeError.
    throw new OpenXmlSchemaError('parseXml: failed to parse XML payload', {
      cause: new OpenXmlSchemaError('parseXml: the only top-level entry is text, not an element'),
    });
  }
  return { root: convertElement(root, initial), rootNamespaces: declarationsOf(root[ATTR_KEY] as RawAttrs | undefined) };
}

const declarationsOf = (attrs: RawAttrs | undefined): Array<{ prefix: string; ns: string }> => {
  const out: Array<{ prefix: string; ns: string }> = [];
  for (const [k, v] of Object.entries(attrs ?? {})) {
    if (k === 'xmlns') out.push({ prefix: '', ns: decodeXmlEntities(v) });
    else if (k.startsWith('xmlns:')) out.push({ prefix: k.slice('xmlns:'.length), ns: decodeXmlEntities(v) });
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

const elementTag = (entry: RawEntry): string | undefined => {
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

const extendStack = (parent: NamespaceStack, attrs: RawAttrs | undefined): NamespaceStack => {
  if (attrs === undefined) return parent;
  let nextDefault = parent.default;
  let nextByPrefix: Record<string, string> | undefined;
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'xmlns') {
      nextDefault = decodeXmlEntities(v);
      continue;
    }
    if (k.startsWith('xmlns:')) {
      const prefix = k.slice('xmlns:'.length);
      nextByPrefix ??= { ...parent.byPrefix };
      nextByPrefix[prefix] = decodeXmlEntities(v);
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

const filterAttrs = (rawAttrs: RawAttrs | undefined, stack: NamespaceStack): { resolved: Record<string, string> } => {
  const resolved: Record<string, string> = {};
  if (rawAttrs === undefined) return { resolved };
  for (const [k, v] of Object.entries(rawAttrs)) {
    // xmlns / xmlns:* declarations: dropped from the XmlNode attribute table.
    // The serializer rebuilds them from the Clark-notation namespaces it walks,
    // so round-tripping does not require preserving the declarations.
    if (k === 'xmlns' || k.startsWith('xmlns:')) continue;
    resolved[resolveAttrName(k, stack)] = decodeXmlEntities(v);
  }
  keepMcPrefixDeclarations(resolved, stack);
  return { resolved };
};

// Markup-compatibility attributes name namespaces by prefix inside their
// value (`<mc:Choice Requires="cx1">`, `mc:Ignorable="x14ac"`), which Clark
// notation cannot carry. Keep the declarations of those prefixes as literal
// `xmlns:*` attributes so the serializer writes them back; without them Excel
// cannot resolve the prefix and asks to repair the file.
const MC_PREFIX_LIST_ATTRS = ['Requires', `{${MARKUP_COMPAT_NS}}Ignorable`, `{${MARKUP_COMPAT_NS}}MustUnderstand`, `{${MARKUP_COMPAT_NS}}ProcessContent`];

const keepMcPrefixDeclarations = (resolved: Record<string, string>, stack: NamespaceStack): void => {
  for (const attr of MC_PREFIX_LIST_ATTRS) {
    const value = resolved[attr];
    if (value === undefined) continue;
    // ProcessContent lists qualified names (`w:p`); the others bare prefixes.
    for (const token of value.split(/\s+/)) {
      const prefix = token.split(':')[0];
      if (!prefix) continue;
      const ns = stack.byPrefix[prefix];
      if (ns !== undefined) resolved[`xmlns:${prefix}`] = ns;
    }
  }
};

/**
 * Whether a text run is ignorable between child elements. Exported for the same
 * reason as {@link rejectDtdDeclarations}: the worksheet reader applies this
 * rule to the `<is>` subtrees it builds itself.
 */
export const isWhitespaceOnly = (s: string): boolean => /^\s*$/.test(s);

const convertElement = (entry: RawEntry, parentStack: NamespaceStack): XmlNode => {
  const rawTag = elementTag(entry);
  if (rawTag === undefined) {
    throw new OpenXmlSchemaError('parseXml: encountered an entry with no element tag');
  }
  if (isProcessingInstruction(rawTag)) {
    throw new OpenXmlSchemaError(`parseXml: processing instructions are not supported (saw "<${rawTag}>")`);
  }

  const rawAttrs = entry[ATTR_KEY] as RawAttrs | undefined;
  const stack = extendStack(parentStack, rawAttrs);

  const { resolved } = filterAttrs(rawAttrs, stack);
  const node = el(resolveElementName(rawTag, stack), resolved);

  const childEntries = entry[rawTag] as RawEntry[] | undefined;
  if (childEntries === undefined || childEntries.length === 0) return node;

  const textParts: string[] = [];
  for (const child of childEntries) {
    if (Object.hasOwn(child, TEXT_KEY)) {
      const t = child[TEXT_KEY];
      if (typeof t === 'string') textParts.push(decodeXmlEntities(t));
      continue;
    }
    if (Object.hasOwn(child, CDATA_KEY)) {
      const section = child[CDATA_KEY] as RawEntry[];
      for (const part of section) {
        const t = part[TEXT_KEY];
        if (typeof t === 'string') textParts.push(t);
      }
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
