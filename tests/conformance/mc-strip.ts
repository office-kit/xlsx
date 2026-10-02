// ECMA-376 Part 3 consumer profile for the vendored Transitional schemas.
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { XML_NS, attribute, parseDocument, serializeDocument, type XmlElement } from './xml-tree.js';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const SUPPORTED = new Set([
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  ...['main', 'chart', 'spreadsheetDrawing', 'chartDrawing', 'diagram', 'picture'].map(n => `http://schemas.openxmlformats.org/drawingml/2006/${n}`),
  'http://purl.oclc.org/ooxml/spreadsheetml/main',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  ...['main', 'chart', 'spreadsheetDrawing', 'chartDrawing', 'diagram', 'picture'].map(n => `http://purl.oclc.org/ooxml/drawingml/${n}`),
  XML_NS,
]);
interface Scope { ignorable: Set<string>; processContent: Set<string> }
const words = (s: string | undefined): string[] => s?.trim().split(/\s+/).filter(Boolean) ?? [];
function namespace(node: XmlElement, prefix: string): string {
  const uri = node.namespaces[prefix];
  if (!uri) throw new OpenXmlSchemaError(`MC references undeclared prefix ${prefix}`);
  return uri;
}
function extend(node: XmlElement, parent: Scope): Scope {
  const scope = { ignorable: new Set(parent.ignorable), processContent: new Set(parent.processContent) };
  for (const p of words(attribute(node, 'Ignorable', MC))) scope.ignorable.add(namespace(node, p));
  for (const p of words(attribute(node, 'MustUnderstand', MC))) {
    if (!SUPPORTED.has(namespace(node, p))) throw new OpenXmlSchemaError(`Unsupported mc:MustUnderstand namespace ${p}`);
  }
  for (const q of words(attribute(node, 'ProcessContent', MC))) {
    const [prefix, local] = q.split(':');
    if (!prefix || !local) throw new OpenXmlSchemaError(`Invalid ProcessContent QName ${q}`);
    const uri = namespace(node, prefix);
    if (!scope.ignorable.has(uri)) throw new OpenXmlSchemaError('ProcessContent must name an ignorable namespace');
    scope.processContent.add(`${uri}|${local}`);
  }
  return scope;
}
function process(node: XmlElement, parent: Scope): Array<XmlElement | string> {
  const space = attribute(node, 'space', XML_NS);
  if (space !== undefined && space !== 'default' && space !== 'preserve') throw new OpenXmlSchemaError('Invalid xml:space');
  const scope = extend(node, parent);
  const contents = (n: XmlElement, s: Scope) => n.children.flatMap(c => typeof c === 'string' ? [c] : process(c, s));
  if (node.uri === MC && node.local === 'AlternateContent') {
    let fallback: XmlElement | undefined;
    let selected: XmlElement | undefined;
    let choices = 0;
    for (const c of node.children) {
      if (typeof c === 'string') {
        if (c.trim()) throw new OpenXmlSchemaError('Text in AlternateContent');
        continue;
      }
      if (c.uri !== MC) throw new OpenXmlSchemaError('Invalid AlternateContent child');
      if (c.local === 'Choice') {
        if (fallback) throw new OpenXmlSchemaError('Choice after Fallback');
        choices++;
        const requires = words(attribute(c, 'Requires'));
        if (!requires.length) throw new OpenXmlSchemaError('Choice requires a non-empty Requires');
        const understood = requires.map(p => SUPPORTED.has(namespace(c, p)));
        if (!selected && understood.every(Boolean)) selected = c;
      } else if (c.local === 'Fallback') {
        if (fallback) throw new OpenXmlSchemaError('Duplicate Fallback');
        fallback = c;
      }
      else throw new OpenXmlSchemaError('Invalid AlternateContent child');
    }
    if (!choices) throw new OpenXmlSchemaError('AlternateContent requires a Choice');
    const branch = selected ?? fallback;
    return branch ? contents(branch, extend(branch, scope)) : [];
  }
  if (scope.ignorable.has(node.uri) && !SUPPORTED.has(node.uri)) {
    return scope.processContent.has(`${node.uri}|${node.local}`) ? contents(node, scope) : [];
  }
  if (node.uri === MC) throw new OpenXmlSchemaError('MC element outside AlternateContent');
  return [{ ...node, children: contents(node, scope), attributes: node.attributes.filter(a =>
    a.uri !== MC && !(a.uri === XML_NS && a.local === 'space') &&
    !(scope.ignorable.has(a.uri) && !SUPPORTED.has(a.uri))) }];
}
/** xml:space alone is removed for the vendored XSD's string-type limitation. */
export function stripIgnorableMarkup(xml: string): string {
  const result = process(parseDocument(xml), { ignorable: new Set(), processContent: new Set() });
  const root = result[0];
  if (result.length !== 1 || !root || typeof root === 'string') throw new OpenXmlSchemaError('MC removed the document element');
  return serializeDocument(root);
}
