// src/xml/lex.ts replaced fast-xml-parser as parseXml's lexer and must lex
// exactly as it did: the same tree, or a failure on both. fast-xml-parser stays
// a dev dependency for this comparison only.

import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { XMLParser } from 'fast-xml-parser';
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { lexXml } from '../../src/xml/lex.js';
import { parseXmlDocument, rejectDtdDeclarations } from '../../src/xml/parser.js';

// The options parser.ts passed before the swap.
const fastXmlParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  attributesGroupName: ':@',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: false,
  htmlEntities: false,
  cdataPropName: '#cdata',
});
const decoder = new TextDecoder('utf-8', { fatal: false });

type Outcome = { tree: unknown } | { error: unknown };
const attempt = (lex: () => unknown): Outcome => {
  try {
    return { tree: lex() };
  } catch (error) {
    return { error };
  }
};

/** Whether parseXml refuses the payload before lexing, so neither lexer sees it. */
const refusedByPrescan = (text: string): boolean => {
  try {
    rejectDtdDeclarations(text);
    return false;
  } catch (error) {
    if (error instanceof OpenXmlSchemaError) return true;
    throw error;
  }
};

const expectSameLex = (input: string | Uint8Array): void => {
  const text = typeof input === 'string' ? input : decoder.decode(input);
  if (refusedByPrescan(text)) return;
  const expected = attempt(() => fastXmlParser.parse(text));
  const actual = attempt(() => lexXml(text));
  if ('tree' in expected) expect(actual).toStrictEqual(expected);
  else {
    expect(actual).toHaveProperty('error');
    if ('error' in actual) expect(actual.error).toBeInstanceOf(OpenXmlSchemaError);
  }
};

const N = 'xmlns="urn:a"';
const deep = (n: number): string => `<r ${N}>${'<c>'.repeat(n)}${'</c>'.repeat(n)}</r>`;

const cases: Record<string, string> = {
  'attr TAB': `<r ${N} v="a\tb"/>`,
  'attr LF': `<r ${N} v="a\nb"/>`,
  'attr CR': `<r ${N} v="a\rb"/>`,
  'attr CRLF': `<r ${N} v="a\r\nb"/>`,
  'attr char refs': `<r ${N} v="a&#10;b&#x9;c"/>`,
  'text CR': `<r ${N}>a\rb</r>`,
  'text CRLF': `<r ${N}>a\r\nb</r>`,
  'text LF': `<r ${N}>a\nb</r>`,
  '5 entities': `<r ${N} v="&amp;&lt;&gt;&quot;&apos;">&amp;&lt;&gt;&quot;&apos;</r>`,
  'numeric refs': `<r ${N}>&#65;&#x42;&#x1F600;</r>`,
  'single pass': `<r ${N}>&amp;#65;</r>`,
  'bare & text': `<r ${N}>a & b</r>`,
  'bare & attr': `<r ${N} v="a & b"/>`,
  'unknown entity text': `<r ${N}>a&nbsp;b</r>`,
  'unknown entity attr': `<r ${N} v="a&nbsp;b"/>`,
  'invalid numeric &#0;': `<r ${N}>&#0;</r>`,
  'invalid numeric &#xFFFE;': `<r ${N}>&#xFFFE;</r>`,
  'CDATA literal': `<r ${N}><![CDATA[&amp; <x>]]></r>`,
  'CDATA CR': `<r ${N}><![CDATA[a\r\nb]]></r>`,
  'CDATA + text': `<r ${N}>x<![CDATA[y]]>z</r>`,
  'xml decl + PI': `<?xml version="1.0"?><?pi data?><r ${N}/>`,
  'comment inside': `<r ${N}><!-- c --><c/></r>`,
  'comment in text': `<r ${N}>a<!-- c -->b</r>`,
  'PI inside': `<r ${N}><?pi x?><c/></r>`,
  'mixed content': `<r ${N}>t<c/></r>`,
  'trailing text': `<r ${N}><c/>t</r>`,
  'ws between': `<r ${N}>\n  <c/>\n</r>`,
  'prefixes + clark': `<a:r xmlns:a="urn:a" xmlns:b="urn:b" xmlns="urn:d" b:x="1" y="2" xml:space="preserve"><c/><b:c/></a:r>`,
  'undeclared prefix': `<a:r xmlns="urn:d"/>`,
  'undeclared attr prefix': `<r ${N} z:q="1"/>`,
  'duplicate attrs': `<r ${N} v="1" v="2"/>`,
  '< in attr': `<r ${N} v="a<b"/>`,
  '> in attr': `<r ${N} v="a>b"/>`,
  'mismatched close': `<r ${N}><c></d></r>`,
  'unclosed': `<r ${N}><c>`,
  'two roots': `<r ${N}/><s ${N}/>`,
  'no root': `<?xml version="1.0"?>`,
  'DOCTYPE': `<!DOCTYPE r><r/>`,
  'ENTITY': `<r ${N}><!ENTITY x "y"></r>`,
  BOM: `﻿<r ${N}/>`,
  'empty text el': `<r ${N}><c></c></r>`,
  'single-quoted attr LF': `<r ${N} v='a\nb'/>`,
  'ctrl char in text': `<r ${N}>a\u0001b</r>`,
  'lt in text': `<r ${N}>a < b</r>`,
  ']]> in text': `<r ${N}>a ]]> b</r>`,
  'unquoted attr': `<r ${N} v=1/>`,
  'attr no value': `<r ${N} v/>`,
  'bad name': `<1r ${N}/>`,
  'text before root': `x<r ${N}/>`,
  'text after root': `<r ${N}/>x`,
  'nested mismatch': `<r ${N}><a><b></a></b></r>`,
  'close without open': `<r ${N}></x></r>`,
  'attr dup prefixed': `<r ${N} xmlns:p="urn:p" p:v="1" p:v="2"/>`,
  'xmlns redefine nested': `<p:r xmlns:p="urn:1"><p:c xmlns:p="urn:2"/></p:r>`,
  'default ns unset': `<r xmlns="urn:a"><c xmlns=""/></r>`,
  'three roots': `<r ${N}/><s ${N}/><t ${N}/>`,
  'two roots, 2nd bad prefix': `<r ${N}/><z:s/>`,
  'both prefixes undeclared': `<a:r z:q="1"/>`,
  'undeclared nested element': `<r ${N}><q:c/></r>`,
  ']]> in attr': `<r ${N} v="a]]>b"/>`,
  ']]> next to entity': `<r ${N}>]]&gt; ]]> &amp;</r>`,
  'CDATA then ]]> text': `<r ${N}><![CDATA[x]]>]]></r>`,
  'second root with child': `<r ${N}/><s ${N}><c/></s>`,
  'M1 13-char dec ref': `<r ${N}>&#0000000065;</r>`,
  'M1 14-char hex ref': `<r ${N}>&#x0000000041;</r>`,
  'M1 long ref in attr': `<r ${N} v="&#0000000065;"/>`,
  'M1 long ref in ns URI': `<p:r xmlns:p="urn:&#0000000065;"/>`,
  'M1 12-char boundary': `<r ${N}>&#000000065;</r>`,
  'M2 bad ref in CDATA': `<r ${N}><![CDATA[&#0;]]></r>`,
  'M2 bad ref in comment': `<r ${N}><!-- &#xFFFE; --><c/></r>`,
  'M2 bad ref in PI': `<?pi &#0;?><r ${N}/>`,
  'quote inside other quote': `<r ${N} v="it's\n"/>`,
};

const namespaceCases: Record<string, string> = {
  // rootNamespaces: order, default first or last, redeclaration on a child,
  // xmlns="" on the root, an entity inside a namespace URI.
  'rootNS default first': `<p:r xmlns="urn:d" xmlns:p="urn:p" xmlns:q="urn:q"/>`,
  'rootNS default last': `<p:r xmlns:q="urn:q" xmlns:p="urn:p" xmlns="urn:d"/>`,
  'rootNS child redeclares': `<p:r xmlns:p="urn:p"><p:c xmlns:p="urn:x" xmlns:z="urn:z"/></p:r>`,
  'rootNS empty default': `<r xmlns=""/>`,
  'rootNS entity in URI': `<p:r xmlns:p="urn:a&amp;b"/>`,
  'rootNS mc:Ignorable workbook': `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><sheets/></workbook>`,
  // Which error wins when an input has several.
  'priority: bad ref in 1st of 2 roots': `<r ${N}>&#0;</r><s ${N}/>`,
  'priority: bad prefix on 1st of 2 roots': `<a:r/><s ${N}/>`,
  'priority: bad attr prefix and bad ref': `<a:r z:v="&#0;"/>`,
};

const quirkCases: Record<string, string> = {
  'depth 100': deep(99),
  'depth 101': deep(100),
  'depth 102': deep(101),
  'tag __proto__': `<r ${N}><__proto__/></r>`,
  'tag constructor open': `<r ${N}><constructor></constructor></r>`,
  'tag toString': `<r ${N}><toString/></r>`,
  'tag toString spaced': `<r ${N}><toString /></r>`,
  'attr __proto__': `<r ${N} __proto__="1"/>`,
  'attr valueOf': `<r ${N} valueOf="1"/>`,
  'attr __proto__ no value': `<r ${N} __proto__/>`,
  'tag #text': `<r ${N}><#text/></r>`,
  'tag :@': `<r ${N}><:@/></r>`,
  'close __proto__': `<r ${N}><c></__proto__></r>`,
  'PI attr __proto__': `<?pi __proto__="1"?><r ${N}/>`,
  'PI name only __proto__': `<?__proto__?><r ${N}/>`,
  'top text before close': `<r ${N}/>x</q>`,
  'top text then close only': `x</q>`,
  'open after over-close': `<r ${N}></r></q><s ${N}/>`,
  'CDATA after over-close': `<r ${N}></r></q><![CDATA[x]]>`,
  'CDATA at top level': `<![CDATA[x]]><r ${N}/>`,
  '<!D not doctype': `<r ${N}><!Dx></r>`,
  '<!x tag': `<r ${N}><!x/></r>`,
  'unterminated comment': `<r ${N}><!-- x`,
  'unterminated CDATA': `<r ${N}><![CDATA[x`,
  'unterminated PI': `<r ${N}><?pi x`,
  'unterminated tag': `<r ${N}><c`,
  'unterminated close': `<r ${N}></r`,
  'TAB in tag': `<r\t${N}\tv="a\tb"/>`,
  'attr same as name': `<r r/>`,
  'numeric attr names': `<r ${N} b="1" 2="x" a="2" 1="y"/>`,
  'duplicate xmlns': `<r xmlns="urn:a" xmlns="urn:b"/>`,
  'duplicate prefix': `<p:r xmlns:p="urn:1" xmlns:p="urn:2"/>`,
  'close before root': `</x><r ${N}/>`,
  'unclosed root only': `<r ${N}>`,
  'close wrong name': `<r ${N}><c></r>`,
  'lone high surrogate text': `<r ${N}>a\uD800b</r>`,
  'lone low surrogate attr': `<r ${N} v="\uDFFF"/>`,
  'noncharacter FFFE text': `<r ${N}>a\uFFFEb</r>`,
  'astral pair': `<r ${N}>\u{1F600}</r>`,
  'ref surrogate D800': `<r ${N}>&#xD800;</r>`,
  'ref DFFF': `<r ${N}>&#xDFFF;</r>`,
  'ref 55296': `<r ${N}>&#55296;</r>`,
  'ref 110000': `<r ${N}>&#x110000;</r>`,
  'ref 1114112': `<r ${N}>&#1114112;</r>`,
  'ref 10FFFF': `<r ${N}>&#x10FFFF;</r>`,
  'ref E000': `<r ${N}>&#xE000;</r>`,
  'ref FFFD': `<r ${N}>&#xFFFD;</r>`,
  'ref 20': `<r ${N}>&#x20;</r>`,
  'ref uppercase X': `<r ${N}>&#X41;</r>`,
  'comment double hyphen': `<r ${N}><!-- a -- b --><c/></r>`,
  'xml-stylesheet PI': `<?xml version="1.0"?><?xml-stylesheet href="a"?><r ${N}/>`,
  'two CDATA': `<r ${N}><![CDATA[a]]><![CDATA[b]]></r>`,
  'lt in text then quote': `<r ${N}>a < b's</r>`,
  'lt in text then dquote': `<r ${N}>a < "b</r>`,
  'gt in text': `<r ${N}>a > b</r>`,
  'quote in text': `<r ${N}>it's "x"</r>`,
  'empty doc': ``,
  'whitespace doc': ` \n `,
  'self-close no space slash attr': `<r ${N} v="1"/>`,
  'slash inside value': `<r ${N} v="a/"/>`,
  'value ending slash unquoted': `<r ${N} v=a/>`,
  'nested same names': `<r ${N}><r><r/></r></r>`,
  'Object.prototype name isPrototypeOf': `<r ${N}><isPrototypeOf/></r>`,
  'Object.prototype attr propertyIsEnumerable': `<r ${N} propertyIsEnumerable="1"/>`,
  'Object.prototype name toLocaleString': `<r ${N}><toLocaleString/></r>`,
  'Object.prototype name hasOwnProperty': `<r ${N}><hasOwnProperty/></r>`,
  'lowercase doctype': `<!doctype r><r ${N}/>`,
  'IGNORE section': `<![IGNORE[x]]><r ${N}/>`,
  'non-CDATA <![ inside': `<r ${N}><![x]]></r>`,
  'PI unbalanced quote': `<?pi it's?><r ${N}/>`,
  'tag unbalanced quote closed later': `<r ${N}><c v="a>b'></c>"/></r>`,
  'slash then space': `<r ${N}><c/ ></r>`,
  'comment <!-->': `<r ${N}><!--><c/></r>`,
  'comment <!--->': `<r ${N}><!---><c/></r>`,
  'cr only line endings': `<r ${N}>\r<c/>\r</r>`,
};

const byteCases: Record<string, Uint8Array> = {
  'bytes invalid utf-8': new Uint8Array([0x3c, 0x72, 0x20, 0x78, 0x6d, 0x6c, 0x6e, 0x73, 0x3d, 0x22, 0x75, 0x22, 0x3e, 0xff, 0x3c, 0x2f, 0x72, 0x3e]),
  'bytes BOM': new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(`<r ${N}/>`)]),
};

describe('lexXml lexes as fast-xml-parser did', () => {
  it.each(Object.entries({ ...cases, ...namespaceCases, ...quirkCases, ...byteCases }))('%s', (_, input) => {
    expectSameLex(input);
  });

  it('on every XML part of every fixture workbook', () => {
    const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
    // Not zip archives on purpose: an empty file and an OLE (encrypted) workbook.
    const notZip = new Set(['reference/openpyxl/openpyxl/reader/tests/data/null_file.xlsx', 'tests/fixtures/cfb/encrypted.xlsx']);
    let parts = 0;
    for (const dir of ['reference/openpyxl', 'tests/fixtures']) {
      for (const entry of readdirSync(join(root, dir), { recursive: true, encoding: 'utf8' })) {
        if (!/\.(xlsx|xlsm|xltx|xltm)$/i.test(entry)) continue;
        const file = relative(root, join(root, dir, entry));
        const bytes = readFileSync(join(root, file));
        if (notZip.has(file)) {
          expect(() => unzipSync(bytes)).toThrow();
          continue;
        }
        for (const [name, part] of Object.entries(unzipSync(bytes))) {
          if (!/\.(xml|rels|vml)$/i.test(name)) continue;
          parts++;
          expectSameLex(part);
        }
      }
    }
    // Guards against the walk silently finding nothing (e.g. no submodule).
    expect(parts).toBeGreaterThan(1500);
  });
});

describe('parseXmlDocument with text as the only top-level entry', () => {
  // fast-xml-parser lexed these to a single text entry, and converting it as an
  // element recursed until a raw RangeError escaped. This is the one intended
  // difference from the old parser: a schema error now, the lex is unchanged.
  it.each(['x</q>', ' x </q>', 'x</q></q>'])('%j fails with OpenXmlSchemaError', (input) => {
    const outcome = attempt(() => parseXmlDocument(input));
    expect(outcome).toHaveProperty('error');
    if (!('error' in outcome)) return;
    expect(outcome.error).toBeInstanceOf(OpenXmlSchemaError);
    expect(outcome.error).toHaveProperty('message', 'parseXml: failed to parse XML payload');
    expect(outcome.error).toHaveProperty('cause', expect.any(OpenXmlSchemaError));
  });

  it('still counts that text as a root next to an element', () => {
    expect(() => parseXmlDocument(`<r ${N}/>x</q>`)).toThrow('parseXml: document has 2 root elements; expected exactly one');
  });
});
