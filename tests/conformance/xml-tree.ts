// Independent, namespace-aware test oracle. Do not import production XML parsers.
import { SaxesParser, type SaxesAttributeNS } from 'saxes';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
export const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
export interface XmlElement {
  name: string;
  local: string;
  uri: string;
  attributes: SaxesAttributeNS[];
  namespaces: Record<string, string>;
  children: Array<XmlElement | string>;
}
export function parseDocument(xml: string): XmlElement {
  const parser = new SaxesParser({ xmlns: true });
  const stack: XmlElement[] = [];
  let root: XmlElement | undefined;
  parser.on('error', cause => { throw new OpenXmlSchemaError(cause.message, { cause }); });
  parser.on('doctype', () => { throw new OpenXmlSchemaError('DOCTYPE is forbidden in an OPC XML part'); });
  parser.on('opentag', tag => {
    const parent = stack.at(-1);
    const element: XmlElement = { name: tag.name, local: tag.local, uri: tag.uri,
      attributes: Object.values(tag.attributes), namespaces: { ...parent?.namespaces, ...tag.ns }, children: [] };
    if (parent) parent.children.push(element); else root = element;
    stack.push(element);
  });
  const append = (text: string) => { stack.at(-1)?.children.push(text); };
  parser.on('text', append);
  parser.on('cdata', append);
  parser.on('closetag', () => { stack.pop(); });
  parser.write(xml).close();
  if (!root) throw new OpenXmlSchemaError('Missing document element');
  return root;
}
export function attribute(node: XmlElement, local: string, uri = ''): string | undefined {
  return node.attributes.find(a => a.local === local && a.uri === uri)?.value;
}
export function children(node: XmlElement, local: string, uri = node.uri): XmlElement[] {
  return node.children.filter((c): c is XmlElement => typeof c !== 'string' && c.local === local && c.uri === uri);
}
export function* descendants(node: XmlElement): Generator<XmlElement> {
  yield node;
  for (const child of node.children) if (typeof child !== 'string') yield* descendants(child);
}
export function textContent(node: XmlElement): string {
  return node.children.map(c => typeof c === 'string' ? c : textContent(c)).join('');
}
function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r/g, '&#13;');
}
export function serializeDocument(node: XmlElement, parent: Record<string, string> = {}): string {
  const attrs: string[] = [];
  const quote = (s: string) => escapeText(s).replace(/"/g, '&quot;').replace(/\n/g, '&#10;').replace(/\t/g, '&#9;');
  for (const [prefix, uri] of Object.entries(node.namespaces)) {
    if (prefix === 'xml' || prefix === 'xmlns' || parent[prefix] === uri) continue;
    attrs.push(`${prefix ? `xmlns:${prefix}` : 'xmlns'}="${quote(uri)}"`);
  }
  for (const a of node.attributes) if (a.uri !== XMLNS_NS) attrs.push(`${a.name}="${quote(a.value)}"`);
  const start = `<${node.name}${attrs.length ? ` ${attrs.join(' ')}` : ''}`;
  if (!node.children.length) return `${start}/>`;
  return `${start}>${node.children.map(c => typeof c === 'string' ? escapeText(c) : serializeDocument(c, node.namespaces)).join('')}</${node.name}>`;
}
