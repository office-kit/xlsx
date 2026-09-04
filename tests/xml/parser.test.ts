import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { OpenXmlSchemaError } from '../../src/utils/exceptions';
import { REL_NS, SHEET_MAIN_NS, XML_NS } from '../../src/xml/namespaces';
import { parseXml } from '../../src/xml/parser';
import { findChild, findChildren } from '../../src/xml/tree';
import { openZip } from '../../src/zip/reader';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

describe('parseXml — minimal cases', () => {
  it('parses a single root element with no attrs or children', () => {
    const root = parseXml('<r/>');
    expect(root.name).toBe('r');
    expect(root.attrs).toEqual({});
    expect(root.children).toEqual([]);
    expect(root.text).toBeUndefined();
  });

  it('parses a text-only element', () => {
    const root = parseXml('<t>hello world</t>');
    expect(root.name).toBe('t');
    expect(root.text).toBe('hello world');
    expect(root.children).toEqual([]);
  });

  it('preserves whitespace inside a text-only element when xml:space="preserve" is in effect (no trim)', () => {
    const root = parseXml('<t xml:space="preserve">  abc  </t>');
    expect(root.text).toBe('  abc  ');
    // xml:space attribute is namespaced under the XML namespace.
    expect(root.attrs[`{${XML_NS}}space`]).toBe('preserve');
  });

  it('parses attributes', () => {
    const root = parseXml('<row r="1" ht="12.5" customHeight="1"/>');
    expect(root.attrs).toEqual({ r: '1', ht: '12.5', customHeight: '1' });
  });

  it('expands the standard XML entities in text', () => {
    const root = parseXml('<t>&amp; &lt; &gt; &apos; &quot;</t>');
    expect(root.text).toBe('& < > \' "');
  });

  it('expands decimal and hexadecimal character references in text and attributes', () => {
    const root = parseXml('<t label="&#65; &#x1F600;">&#20219;&#x52A1;</t>');
    expect(root.attrs['label']).toBe('A 😀');
    expect(root.text).toBe('任务');
  });

  it('decodes entities once rather than recursively', () => {
    const root = parseXml('<t label="&amp;#65;">&amp;#x41;</t>');
    expect(root.attrs['label']).toBe('&#65;');
    expect(root.text).toBe('&#x41;');
  });

  it('skips the XML declaration', () => {
    const root = parseXml('<?xml version="1.0" encoding="UTF-8"?><r/>');
    expect(root.name).toBe('r');
  });

  it('accepts a Uint8Array directly', () => {
    const bytes = new TextEncoder().encode('<r a="1"/>');
    const root = parseXml(bytes);
    expect(root.attrs['a']).toBe('1');
  });
});

describe('parseXml — namespaces', () => {
  it('rewrites the default namespace into Clark notation', () => {
    const xml = `<workbook xmlns="${SHEET_MAIN_NS}"><sheets/></workbook>`;
    const root = parseXml(xml);
    expect(root.name).toBe(`{${SHEET_MAIN_NS}}workbook`);
    expect(root.children[0]?.name).toBe(`{${SHEET_MAIN_NS}}sheets`);
  });

  it('rewrites prefixed elements and attributes', () => {
    const xml = `<workbook xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}"><sheet r:id="rId1"/></workbook>`;
    const root = parseXml(xml);
    const sheet = root.children[0];
    expect(sheet?.name).toBe(`{${SHEET_MAIN_NS}}sheet`);
    expect(sheet?.attrs[`{${REL_NS}}id`]).toBe('rId1');
  });

  it('does not apply the default namespace to unprefixed attributes', () => {
    const xml = `<workbook xmlns="${SHEET_MAIN_NS}"><sheet name="Sheet1"/></workbook>`;
    const sheet = parseXml(xml).children[0];
    // Per XMLNS rules unprefixed attribute names belong to no namespace.
    expect(sheet?.attrs['name']).toBe('Sheet1');
    expect(sheet?.attrs[`{${SHEET_MAIN_NS}}name`]).toBeUndefined();
  });

  it('drops xmlns / xmlns:* declarations from the attribute table', () => {
    const xml = `<r xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}" foo="bar"/>`;
    const root = parseXml(xml);
    expect(Object.keys(root.attrs).sort()).toEqual(['foo']);
  });

  it('throws on an undeclared element prefix', () => {
    expect(() => parseXml('<r:foo/>')).toThrowError(OpenXmlSchemaError);
  });

  it('throws on an undeclared attribute prefix', () => {
    expect(() => parseXml('<foo r:bar="x"/>')).toThrowError(OpenXmlSchemaError);
  });
});

describe('parseXml — security & validation', () => {
  it('rejects DOCTYPE declarations with OpenXmlSchemaError', () => {
    const xml = '<!DOCTYPE foo SYSTEM "http://example/foo.dtd"><foo/>';
    expect(() => parseXml(xml)).toThrowError(OpenXmlSchemaError);
  });

  it('rejects ENTITY declarations even without DOCTYPE wrapping (defence in depth)', () => {
    const xml = '<!ENTITY foo "bar"><r/>';
    expect(() => parseXml(xml)).toThrowError(OpenXmlSchemaError);
  });

  it('rejects mixed content (text between sibling elements)', () => {
    const xml = '<r>text<a/></r>';
    expect(() => parseXml(xml)).toThrowError(OpenXmlSchemaError);
  });

  it('throws when there is no root element', () => {
    expect(() => parseXml('<?xml version="1.0"?>')).toThrowError(OpenXmlSchemaError);
  });
});

describe('parseXml — openpyxl genuine/empty.xlsx workbook part', () => {
  it('parses xl/workbook.xml with the expected sheet metadata', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const root = parseXml(zip.read('xl/workbook.xml'));
    expect(root.name).toBe(`{${SHEET_MAIN_NS}}workbook`);

    const sheets = findChild(root, `{${SHEET_MAIN_NS}}sheets`);
    if (!sheets) throw new Error('expected <sheets> child to be present');
    const sheetEls = findChildren(sheets, `{${SHEET_MAIN_NS}}sheet`);
    expect(sheetEls).toHaveLength(3);

    expect(sheetEls.map((s) => s.attrs['name'])).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    expect(sheetEls.map((s) => s.attrs['sheetId'])).toEqual(['1', '2', '3']);
    expect(sheetEls.map((s) => s.attrs[`{${REL_NS}}id`])).toEqual(['rId1', 'rId2', 'rId3']);
  });

  it('parses [Content_Types].xml from the same fixture', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const root = parseXml(zip.read('[Content_Types].xml'));
    // Loose check: the root is a Types element under the package content-types namespace.
    expect(root.name.endsWith('}Types')).toBe(true);
    expect(root.children.length).toBeGreaterThan(0);
  });
});
