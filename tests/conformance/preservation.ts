// Inspect serialized OPC packages, without the production model/parser.
import { posix } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { attribute, children, descendants, parseDocument, type XmlElement } from './xml-tree.js';
import { required } from './required.js';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
export interface PreservationContract {
  parts: string[];
  relationshipTypes: string[];
  relationshipSources?: string[];
  fragments?: Array<{ part: string; local: string; uri: string }>;
}
function canonical(node: XmlElement): unknown {
  return { uri: node.uri, local: node.local,
    attributes: node.attributes.filter(a => a.uri !== 'http://www.w3.org/2000/xmlns/').map(a => [a.uri, a.local, a.value]).sort(),
    children: node.children.filter(c => typeof c !== 'string' || c.trim() !== '').map(c => typeof c === 'string' ? c : canonical(c)) };
}
export function preservationProjection(bytes: Uint8Array, contract: PreservationContract): unknown {
  const parts = unzipSync(bytes);
  const contentTypes = parseDocument(strFromU8(required(parts['[Content_Types].xml'])));
  const types = new Map(children(contentTypes, 'Override', CT).map(n => [attribute(n, 'PartName'), attribute(n, 'ContentType')]));
  const defaults = new Map(children(contentTypes, 'Default', CT).map(n => [attribute(n, 'Extension'), attribute(n, 'ContentType')]));
  const relationships: string[][] = [];
  const bindings: string[][] = [];
  for (const [path, data] of Object.entries(parts)) {
    if (!path.endsWith('.rels')) continue;
    const source = path === '_rels/.rels' ? '' : path.replace(/(^|\/)_rels\//, '$1').slice(0, -5);
    if (contract.relationshipSources && !contract.relationshipSources.includes(source)) continue;
    const byId = new Map<string, string[]>();
    for (const n of children(parseDocument(strFromU8(data)), 'Relationship', REL)) {
      const type = required(attribute(n, 'Type'));
      const mode = attribute(n, 'TargetMode') ?? 'Internal';
      const target = required(attribute(n, 'Target'));
      const resolved = mode === 'External' ? target : target.startsWith('/') ? posix.normalize(target).slice(1) : posix.normalize(posix.join(posix.dirname(source), target));
      const edge = [source, type, mode, resolved];
      if (contract.relationshipTypes.includes(type.split('/').at(-1) ?? '')) {
        if (mode === 'Internal' && !parts[resolved]) throw new Error(`Missing relationship target ${resolved}`);
        relationships.push(edge);
        byId.set(required(attribute(n, 'Id')), edge);
      }
    }
    if (source && parts[source] && source.endsWith('.xml')) {
      for (const node of descendants(parseDocument(strFromU8(required(parts[source]))))) {
        for (const attr of node.attributes) {
          if (attr.uri !== R) continue;
          const edge = byId.get(attr.value);
          if (edge) bindings.push([node.uri, node.local, attr.local, ...edge]);
        }
      }
    }
  }
  return {
    parts: contract.parts.map(path => {
      const data = parts[path];
      if (!data) throw new Error(`Missing preserved part ${path}`);
      const type = types.get('/' + path) ?? defaults.get(path.split('.').at(-1));
      if (!type) throw new Error(`Missing content type ${path}`);
      return { path, type, data: Array.from(data) };
    }),
    relationships: relationships.sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
    bindings: bindings.sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
    fragments: (contract.fragments ?? []).map(f => ({ part: f.part, local: f.local, uri: f.uri,
      nodes: [...descendants(parseDocument(strFromU8(required(parts[f.part]))))].filter(n => n.local === f.local && n.uri === f.uri).map(canonical) })),
  };
}
