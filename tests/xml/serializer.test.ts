import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { CHART_NS, PKG_REL_NS, REL_NS, SHEET_MAIN_NS, XML_NS } from '../../src/xml/namespaces.js';
import { parseXml } from '../../src/xml/parser.js';
import { serializeXml } from '../../src/xml/serializer.js';
import { el } from '../../src/xml/tree.js';
import { openZip } from '../../src/zip/reader.js';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');
const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

describe('serializeXml — minimal cases', () => {
  it('emits a self-closing element when there is no text or children', () => {
    expect(decode(serializeXml(el('r')))).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<r/>');
  });

  it('emits a text-only element', () => {
    expect(decode(serializeXml(el('t', {}, [], 'hello')))).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<t>hello</t>',
    );
  });

  it('emits attributes in insertion order', () => {
    const node = el('row', { r: '1', ht: '12.5', customHeight: 'true' });
    expect(decode(serializeXml(node))).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<row r="1" ht="12.5" customHeight="true"/>',
    );
  });

  it('escapes text content', () => {
    expect(decode(serializeXml(el('t', {}, [], 'a & b < c > d')))).toContain('<t>a &amp; b &lt; c &gt; d</t>');
  });

  it('escapes attribute values, leaving whitespace literal', () => {
    // The canonical escapeXmlAttr keeps `\r` / `\n` / `\t` literal so the
    // parser (which doesn't decode `&#9;` / `&#10;` / `&#13;`) round-trips
    // them faithfully. See src/utils/escape.ts for the rationale.
    const node = el('a', { v: 'x"y\nz\tw' });
    expect(decode(serializeXml(node))).toContain('<a v="x&quot;y\nz\tw"/>');
  });

  it('lets the caller suppress the XML declaration', () => {
    expect(decode(serializeXml(el('r'), { xmlDeclaration: false }))).toBe('<r/>');
  });

  it('omits the standalone attribute when requested', () => {
    expect(decode(serializeXml(el('r'), { standalone: 'omit' }))).toBe('<?xml version="1.0" encoding="UTF-8"?>\n<r/>');
  });
});

describe('serializeXml — namespaces', () => {
  it('emits the root namespace as the default when DEFAULT_PREFIXES says it should be default', () => {
    const node = el(`{${SHEET_MAIN_NS}}workbook`);
    const out = decode(serializeXml(node));
    expect(out).toContain(`<workbook xmlns="${SHEET_MAIN_NS}"/>`);
  });

  it('emits prefixed declarations for non-default namespaces', () => {
    const node = el(`{${SHEET_MAIN_NS}}workbook`, {}, [el(`{${SHEET_MAIN_NS}}sheet`, { [`{${REL_NS}}id`]: 'rId1' })]);
    const out = decode(serializeXml(node));
    expect(out).toContain(`xmlns="${SHEET_MAIN_NS}"`);
    expect(out).toContain(`xmlns:r="${REL_NS}"`);
    expect(out).toContain('<sheet r:id="rId1"/>');
  });

  it('keeps no-namespace attributes unprefixed', () => {
    const node = el(`{${SHEET_MAIN_NS}}sheet`, { name: 'Sheet1' });
    expect(decode(serializeXml(node))).toContain('<sheet xmlns=');
    expect(decode(serializeXml(node))).toContain('name="Sheet1"');
  });

  it('uses the canonical xml: binding for {http://www.w3.org/XML/1998/namespace}space', () => {
    const node = el(`{${SHEET_MAIN_NS}}t`, { [`{${XML_NS}}space`]: 'preserve' }, [], '  hi  ');
    const out = decode(serializeXml(node));
    expect(out).toContain('xml:space="preserve"');
  });

  it('auto-allocates ns0/ns1 for unrecognised namespaces', () => {
    const someNs = 'urn:example:my-extension';
    const node = el(`{${someNs}}thing`, {});
    const out = decode(serializeXml(node));
    expect(out).toMatch(/<ns0:thing xmlns:ns0="urn:example:my-extension"\/>/);
  });

  it('does not emit the default namespace when the root has no namespace', () => {
    const node = el('plain');
    expect(decode(serializeXml(node))).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<plain/>');
  });
});

describe('serializeXml — round-trip with parseXml', () => {
  it('parse → serialize → parse gives an equivalent tree (synthetic)', () => {
    const original = `<workbook xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}"><fileVersion appName="xl"/><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    const a = parseXml(original);
    const bytes = serializeXml(a);
    const b = parseXml(bytes);
    expect(b).toEqual(a);
  });

  it('round-trips xl/workbook.xml from openpyxl genuine/empty.xlsx', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const a = parseXml(zip.read('xl/workbook.xml'));
    const bytes = serializeXml(a);
    const b = parseXml(bytes);
    expect(b).toEqual(a);
  });

  it('round-trips [Content_Types].xml', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const a = parseXml(zip.read('[Content_Types].xml'));
    const bytes = serializeXml(a);
    const b = parseXml(bytes);
    expect(b).toEqual(a);
  });

  it('round-trips _rels/.rels (PKG_REL_NS as default namespace)', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const a = parseXml(zip.read('_rels/.rels'));
    expect(a.name).toBe(`{${PKG_REL_NS}}Relationships`);
    const bytes = serializeXml(a);
    const b = parseXml(bytes);
    expect(b).toEqual(a);
  });

  it('round-trips a node containing chart and drawing namespaces', () => {
    const node = el(`{${SHEET_MAIN_NS}}root`, {}, [
      el(`{${CHART_NS}}barChart`, {}, []),
      el(`{${REL_NS}}link`, { id: 'x' }),
    ]);
    const a = parseXml(serializeXml(node));
    expect(a).toEqual(node);
  });
});

// `Requires` / `mc:Ignorable` name namespaces by prefix inside their value.
// Excel asks to repair a drawing whose `<mc:Choice Requires="cx1">` lost the
// `xmlns:cx1` declaration on a re-save.
describe('markup-compatibility prefixes', () => {
  const CX1 = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex';
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

  it('keeps the declaration of a prefix named by Requires', () => {
    const src = `<root xmlns:mc="${MC}"><mc:AlternateContent><mc:Choice xmlns:cx1="${CX1}" Requires="cx1"><a/></mc:Choice></mc:AlternateContent></root>`;
    const out = new TextDecoder().decode(serializeXml(parseXml(src), { xmlDeclaration: false }));
    expect(out).toContain(`<mc:Choice Requires="cx1" xmlns:cx1="${CX1}">`);
  });

  it('declares a prefix named by mc:Ignorable only once on the root', () => {
    const X14AC = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac';
    const src = `<r xmlns:mc="${MC}" xmlns:x14ac="${X14AC}" mc:Ignorable="x14ac"><c x14ac:dyDescent="1"/></r>`;
    const out = new TextDecoder().decode(serializeXml(parseXml(src), { xmlDeclaration: false }));
    expect(out.match(/xmlns:x14ac=/g)).toHaveLength(1);
    expect(parseXml(out).children[0]?.attrs[`{${X14AC}}dyDescent`]).toBe('1');
  });
});
