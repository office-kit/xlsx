// The lexing half of parseXml. It replaced fast-xml-parser (preserveOrder mode
// with the options parser.ts used) and reproduces that output on purpose,
// quirks included, so the swap changed no parse result. tests/xml/lex.test.ts
// compares the two on hand-written cases and on every fixture part. Namespace
// resolution, entity decoding and text folding stay in parser.ts.
//
// The behaviour kept, which the comparison pins:
// - CR and CRLF become LF before anything else.
// - A tag runs to the first `>` outside a quoted value. Its name is the text up
//   to the first whitespace; a trailing `/` makes it self-closing.
// - Attributes are `name = "value"` / `'value'` pairs. A name without a quoted
//   value is dropped. A repeated name keeps the first position and last value.
// - A closing tag closes the innermost open element whatever its name.
// - Elements still open at the end are kept as they are.
// - Text outside every element is dropped, except text right before a closing
//   tag at the top level, which becomes a top-level entry.
// - Comments are dropped; PIs become `?name` entries; CDATA is literal.
// - Deeper than 100 open elements, the names `__proto__` / `constructor` /
//   `prototype`, and the reserved names `#text` / `#cdata` / `:@` fail. Names
//   in a fixed list of Object.prototype members get a `__` prefix.

import { OpenXmlSchemaError } from '../utils/exceptions.js';

export const TEXT_KEY = '#text';
export const CDATA_KEY = '#cdata';
export const ATTR_KEY = ':@';

export type RawAttrs = Record<string, string>;
/** `{ [tag]: children, ':@'?: attrs }`, `{ '#text': text }` or `{ '#cdata': [{ '#text': text }] }`. */
export type RawEntry = { ':@'?: RawAttrs } & { [tagOrText: string]: RawEntry[] | string | RawAttrs | undefined };

type LexNode =
  | { kind: 'element'; tag: string; attrs: RawAttrs | undefined; children: LexNode[] }
  | { kind: 'text'; text: string }
  | { kind: 'cdata'; text: string };
type LexElement = Extract<LexNode, { kind: 'element' }>;

const MAX_DEPTH = 100;
const FORBIDDEN_NAMES = new Set(['__proto__', 'constructor', 'prototype']);
const RENAMED_NAMES = new Set([
  'hasOwnProperty',
  'toString',
  'valueOf',
  '__defineGetter__',
  '__defineSetter__',
  '__lookupGetter__',
  '__lookupSetter__',
]);
const RESERVED_TAGS = new Set([TEXT_KEY, CDATA_KEY, ATTR_KEY]);
const ATTR_RE = /([^\s=]+)\s*(=\s*(['"])([\s\S]*?)\3)?/g;
// fast-xml-parser's name for the document node. An element spelled `<!xml>`
// is matched by it too, as it was.
const DOCUMENT_TAG = '!xml';
const LT = 0x3c;
const TAB = 0x09;
const QUOTE = 0x22;
const APOSTROPHE = 0x27;

const lexError = (detail: string): OpenXmlSchemaError => new OpenXmlSchemaError(`parseXml: ${detail}`);

const safeName = (name: string): string => {
  if (FORBIDDEN_NAMES.has(name)) throw lexError(`reserved name "${name}"`);
  return RENAMED_NAMES.has(name) ? `__${name}` : name;
};

/** Index of the last character of the first `close` at or after `from`. */
const closing = (s: string, close: string, from: number, what: string): number => {
  const at = s.indexOf(close, from);
  if (at < 0) throw lexError(`${what} is not closed`);
  return at + close.length - 1;
};

/** Tag body from `from` up to the first `close` outside quotes; a TAB outside quotes reads as a space. */
const tagBody = (s: string, from: number, close: string): { body: string; end: number } | undefined => {
  let quote = 0;
  let body = '';
  for (let i = from; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (quote !== 0) {
      if (c === quote) quote = 0;
      body += s[i];
    } else if (c === QUOTE || c === APOSTROPHE) {
      quote = c;
      body += s[i];
    } else if (s.startsWith(close, i)) {
      return { body, end: i };
    } else body += c === TAB ? ' ' : s[i];
  }
  return undefined;
};

const splitTag = (body: string): { name: string; rest: string } => {
  const ws = body.search(/\s/);
  return ws < 0 ? { name: body, rest: body } : { name: body.slice(0, ws), rest: body.slice(ws + 1).trimStart() };
};

const readAttrs = (rest: string): RawAttrs | undefined => {
  let attrs: RawAttrs | undefined;
  for (const m of rest.matchAll(ATTR_RE)) {
    const [, rawName = '', , , value] = m;
    if (rawName.length === 0) continue;
    const name = safeName(rawName);
    if (value === undefined) continue;
    (attrs ??= {})[name] = value;
  }
  return attrs;
};

const emit = (n: LexNode): RawEntry => {
  switch (n.kind) {
    case 'text':
      return { [TEXT_KEY]: n.text };
    case 'cdata':
      return { [CDATA_KEY]: [{ [TEXT_KEY]: n.text }] };
    case 'element': {
      const entry: RawEntry = {};
      entry[n.tag] = n.children.map(emit);
      if (n.attrs !== undefined) entry[ATTR_KEY] = n.attrs;
      return entry;
    }
  }
};

/** Lex an XML payload into top-level entries. Throws {@link OpenXmlSchemaError} on input it cannot lex. */
export function lexXml(input: string): RawEntry[] {
  const s = input.replace(/\r\n?/g, '\n');
  const top: LexElement = { kind: 'element', tag: DOCUMENT_TAG, attrs: undefined, children: [] };
  let node: LexElement | undefined = top;
  const stack: Array<LexElement | undefined> = [];
  let text = '';
  const flush = (parent: LexElement): void => {
    if (text !== '') parent.children.push({ kind: 'text', text });
    text = '';
  };
  const need = (n: LexElement | undefined): LexElement => {
    if (n === undefined) throw lexError('markup after the document was closed');
    return n;
  };
  for (let i = 0; i < s.length; i++) {
    if (s.charCodeAt(i) !== LT) {
      const next = s.indexOf('<', i);
      const stop = next < 0 ? s.length : next;
      text += s.slice(i, stop);
      i = stop - 1;
      continue;
    }
    if (s[i + 1] === '/') {
      const end = closing(s, '>', i, 'closing tag');
      safeName(s.slice(i + 2, end).trim());
      if (node !== undefined) flush(node);
      text = '';
      node = stack.pop();
      i = end;
    } else if (s[i + 1] === '?') {
      const t = tagBody(s, i + 1, '?>');
      if (t === undefined) throw lexError('processing instruction is not closed');
      const { name, rest } = splitTag(t.body);
      if (text !== '') flush(need(node));
      // Read even when unused: a reserved name anywhere in the body fails.
      const attrs = readAttrs(rest);
      need(node).children.push({
        kind: 'element',
        tag: name,
        attrs: name !== rest ? attrs : undefined,
        children: [{ kind: 'text', text: '' }],
      });
      i = t.end + 1;
    } else if (s.startsWith('!--', i + 1)) {
      i = closing(s, '-->', i + 4, 'comment');
    } else if (s.startsWith('!D', i + 1)) {
      // A real DOCTYPE was refused before lexing; anything else here is not a tag.
      throw lexError('unexpected "<!D"');
    } else if (s.startsWith('![', i + 1)) {
      const end = closing(s, ']]>', i, 'CDATA section') - 2;
      if (text !== '') flush(need(node));
      // `substring`, not `slice`: for a short non-CDATA `<![x]]>` the start is
      // past the end, and fast-xml-parser's swapped-range result is kept.
      need(node).children.push({ kind: 'cdata', text: s.substring(i + 9, end) });
      i = end + 2;
    } else {
      const t = tagBody(s, i + 1, '>');
      if (t === undefined) throw lexError('tag is not closed');
      let { name, rest } = splitTag(t.body);
      name = safeName(name);
      if (RESERVED_TAGS.has(name)) throw lexError(`reserved tag name "${name}"`);
      if (node !== undefined && text !== '' && node.tag !== DOCUMENT_TAG) flush(node);
      let selfClosing = false;
      if (rest.endsWith('/')) {
        selfClosing = true;
        if (name.endsWith('/')) {
          name = name.slice(0, -1);
          rest = name;
        } else rest = rest.slice(0, -1);
      }
      const attrs = name !== rest ? readAttrs(rest) : undefined;
      if (selfClosing) name = safeName(name);
      const child: LexElement = { kind: 'element', tag: name, attrs, children: [] };
      if (!selfClosing && stack.length > MAX_DEPTH) throw lexError(`elements nested deeper than ${MAX_DEPTH}`);
      need(node).children.push(child);
      if (!selfClosing) {
        stack.push(node);
        node = child;
      }
      text = '';
      i = t.end;
    }
  }
  return top.children.map(emit);
}
