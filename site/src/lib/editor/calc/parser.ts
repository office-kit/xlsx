// Pratt parser over the token stream. Operator precedence follows Excel, which
// differs from most languages in two places: unary minus binds tighter than
// `^` (`-2^2` is 4) and `^` is left-associative (`2^3^2` is 64).

import type { AstNode, BinaryOp, StructuredSpec, StructuredSpecial } from './ast.ts';
import { CalcParseError, type Token, tokenize } from './lexer.ts';
import { type CalcScalar, ERRORS } from './types.ts';

const BINARY_POWER: ReadonlyMap<string, number> = new Map<BinaryOp, number>([
  ['=', 10],
  ['<>', 10],
  ['<', 10],
  ['>', 10],
  ['<=', 10],
  ['>=', 10],
  ['&', 20],
  ['+', 30],
  ['-', 30],
  ['*', 40],
  ['/', 40],
  ['^', 50],
  [',', 80],
  [' ', 90],
  [':', 100],
]);

const isBinaryOp = (op: string): op is BinaryOp => BINARY_POWER.has(op);
const PREFIX_SIGN_POWER = 55;
const PREFIX_AT_POWER = 85;
const PERCENT_POWER = 70;
const SPILL_POWER = 110;

const STORAGE_PREFIXES = /^(?:_XLFN\.|_XLWS\.|_XLL\.)+/;

/** Upper-case a function name and drop the prefixes files store future functions under (`_xlfn._xlws.SORT`). */
export function normalizeFunctionName(raw: string): string {
  return raw.toUpperCase().replace(STORAGE_PREFIXES, '');
}

/** LAMBDA / LET parameters are stored as `_xlpm.x`. */
export function normalizeLocalName(raw: string): string {
  return raw.toUpperCase().startsWith('_XLPM.') ? raw.slice(6) : raw;
}

const SPECIALS: ReadonlyMap<string, StructuredSpecial> = new Map([
  ['#ALL', '#All'],
  ['#DATA', '#Data'],
  ['#HEADERS', '#Headers'],
  ['#TOTALS', '#Totals'],
  ['#THIS ROW', '#This Row'],
]);

// `'` escapes the next character inside a structured reference column name.
const unescapeColumn = (s: string): string => s.replace(/'(.)/g, '$1');

/** Parse the inside of `Table1[...]`. */
export function parseStructuredBody(body: string, position: number): StructuredSpec {
  const trimmed = body.trim();
  if (trimmed === '') return { specials: [] };
  if (trimmed.startsWith('@')) {
    const rest = trimmed.slice(1).trim();
    if (rest === '') return { specials: ['#This Row'] };
    if (!rest.startsWith('[')) return { specials: ['#This Row'], col1: unescapeColumn(rest) };
    const inner = parseStructuredBody(rest, position);
    return { ...inner, specials: ['#This Row'] };
  }
  if (!trimmed.startsWith('[')) {
    const special = SPECIALS.get(trimmed.toUpperCase());
    return special !== undefined ? { specials: [special] } : { specials: [], col1: unescapeColumn(trimmed) };
  }
  // A comma-separated list of `[item]` groups; columns may form a `[A]:[B]` span.
  const specials: StructuredSpecial[] = [];
  const columns: string[] = [];
  let i = 0;
  let pendingSpan = false;
  while (i < trimmed.length) {
    const ch = trimmed[i];
    if (ch === ' ' || ch === ',') {
      i++;
      continue;
    }
    if (ch === ':') {
      pendingSpan = true;
      i++;
      continue;
    }
    if (ch !== '[') throw new CalcParseError('Invalid structured reference', position);
    let j = i + 1;
    let item = '';
    while (j < trimmed.length && trimmed[j] !== ']') {
      if (trimmed[j] === "'") j++;
      item += trimmed[j] ?? '';
      j++;
    }
    if (j >= trimmed.length) throw new CalcParseError('Invalid structured reference', position);
    const special = SPECIALS.get(item.trim().toUpperCase());
    if (special !== undefined) {
      specials.push(special);
    } else {
      if (columns.length > 0 && !pendingSpan) throw new CalcParseError('Invalid structured reference', position);
      columns.push(item);
    }
    pendingSpan = false;
    i = j + 1;
  }
  const [col1, col2] = columns;
  return {
    specials,
    ...(col1 !== undefined ? { col1 } : {}),
    ...(col2 !== undefined ? { col2 } : {}),
  };
}

class Parser {
  private pos = 0;
  private readonly tokens: readonly Token[];
  private readonly length: number;

  constructor(tokens: readonly Token[], length: number) {
    this.tokens = tokens;
    this.length = length;
  }

  parse(): AstNode {
    if (this.tokens.length === 0) throw new CalcParseError('Empty formula', 0);
    const node = this.expression(0, false);
    const rest = this.peek();
    if (rest !== undefined) throw new CalcParseError('Unexpected token', rest.start);
    return node;
  }

  private peek(): Token | undefined {
    return this.tokens[this.pos];
  }

  private take(): Token {
    const t = this.tokens[this.pos];
    if (t === undefined) throw new CalcParseError('Unexpected end of formula', this.length);
    this.pos++;
    return t;
  }

  private expect(kind: Token['kind'], what: string): Token {
    const t = this.peek();
    if (t === undefined || t.kind !== kind) throw new CalcParseError(`Expected ${what}`, t?.start ?? this.length);
    this.pos++;
    return t;
  }

  private expression(minPower: number, union: boolean): AstNode {
    let left = this.prefix(union);
    for (;;) {
      const t = this.peek();
      if (t === undefined) break;
      let op: string | undefined;
      if (t.kind === 'op') {
        op = t.op;
      } else if (t.kind === 'comma' && union) {
        op = ',';
      } else if (t.spaceBefore && isOperandStart(t)) {
        // Whitespace between two operands is the intersection operator.
        op = ' ';
      }
      if (op === undefined) break;

      if (op === '%') {
        if (PERCENT_POWER < minPower) break;
        this.pos++;
        left = { type: 'postfix', op: '%', operand: left };
        continue;
      }
      if (op === '#') {
        if (SPILL_POWER < minPower) break;
        this.pos++;
        left = { type: 'postfix', op: '#', operand: left };
        continue;
      }
      if (!isBinaryOp(op)) throw new CalcParseError(`Unexpected operator '${op}'`, t.start);
      const power = BINARY_POWER.get(op) ?? 0;
      if (power < minPower) break;
      if (op !== ' ') this.pos++;
      const right = this.expression(power + 1, union);
      left = { type: 'binary', op, left, right };
    }
    return left;
  }

  private prefix(union: boolean): AstNode {
    const t = this.take();
    switch (t.kind) {
      case 'number':
        return { type: 'number', value: t.value };
      case 'string':
        return { type: 'string', value: t.value };
      case 'bool':
        return { type: 'bool', value: t.value };
      case 'error':
        return { type: 'error', error: t.error };
      case 'ref':
        return t.area === undefined ? { type: 'error', error: ERRORS.REF } : { type: 'ref', prefix: t.prefix, area: t.area };
      case 'name':
        return { type: 'name', name: normalizeLocalName(t.name), prefix: t.prefix };
      case 'struct':
        return { type: 'structured', table: t.table, spec: parseStructuredBody(t.body, t.start) };
      case 'func': {
        this.expect('lparen', "'('");
        const call: AstNode = { type: 'call', name: normalizeFunctionName(t.name), args: this.argumentList() };
        return this.invocations(call);
      }
      case 'lparen': {
        const inner = this.expression(0, true);
        this.expect('rparen', "')'");
        return this.invocations(inner);
      }
      case 'lbrace':
        return this.arrayConstant();
      case 'op':
        if (t.op === '-' || t.op === '+') {
          return { type: 'unary', op: t.op, operand: this.expression(PREFIX_SIGN_POWER, union) };
        }
        if (t.op === '@') return { type: 'unary', op: '@', operand: this.expression(PREFIX_AT_POWER, union) };
        throw new CalcParseError(`Unexpected operator '${t.op}'`, t.start);
      default:
        throw new CalcParseError('Unexpected token', t.start);
    }
  }

  /** `f(...)(...)`: a LAMBDA-valued expression called immediately. */
  private invocations(callee: AstNode): AstNode {
    let node = callee;
    for (;;) {
      const t = this.peek();
      if (t === undefined || t.kind !== 'lparen' || t.spaceBefore) return node;
      this.pos++;
      node = { type: 'invoke', callee: node, args: this.argumentList() };
    }
  }

  /** Arguments after the opening `(`, consuming the closing `)`. */
  private argumentList(): AstNode[] {
    const args: AstNode[] = [];
    if (this.peek()?.kind === 'rparen') {
      this.pos++;
      return args;
    }
    for (;;) {
      const t = this.peek();
      if (t === undefined) throw new CalcParseError("Expected ')'", this.length);
      if (t.kind === 'comma' || t.kind === 'rparen') {
        args.push({ type: 'missing' });
      } else {
        args.push(this.expression(0, false));
      }
      const sep = this.take();
      if (sep.kind === 'rparen') return args;
      if (sep.kind !== 'comma') throw new CalcParseError("Expected ',' or ')'", sep.start);
    }
  }

  private arrayConstant(): AstNode {
    const rows: CalcScalar[][] = [[]];
    for (;;) {
      const current = rows[rows.length - 1] ?? [];
      current.push(this.arrayElement());
      const sep = this.take();
      if (sep.kind === 'rbrace') break;
      if (sep.kind === 'semicolon') {
        rows.push([]);
        continue;
      }
      if (sep.kind !== 'comma') throw new CalcParseError('Invalid array constant', sep.start);
    }
    const cols = rows[0]?.length ?? 0;
    if (rows.some((r) => r.length !== cols)) throw new CalcParseError('Array constant rows differ in length', this.length);
    return { type: 'array', rows: rows.length, cols, data: rows.flat() };
  }

  private arrayElement(): CalcScalar {
    const t = this.take();
    if (t.kind === 'op' && (t.op === '-' || t.op === '+')) {
      const num = this.take();
      if (num.kind !== 'number') throw new CalcParseError('Invalid array constant', num.start);
      return t.op === '-' ? -num.value : num.value;
    }
    switch (t.kind) {
      case 'number':
      case 'string':
      case 'bool':
        return t.value;
      case 'error':
        return t.error;
      default:
        throw new CalcParseError('Invalid array constant', t.start);
    }
  }
}

const isOperandStart = (t: Token): boolean =>
  t.kind === 'ref' ||
  t.kind === 'name' ||
  t.kind === 'func' ||
  t.kind === 'struct' ||
  t.kind === 'lparen' ||
  t.kind === 'number' ||
  t.kind === 'string' ||
  t.kind === 'bool' ||
  t.kind === 'error' ||
  t.kind === 'lbrace';

/**
 * Parse formula text (without the leading `=`; one is tolerated) into a tree.
 * Throws {@link CalcParseError} carrying the offending offset on bad syntax,
 * which is what the editor uses to refuse input the way Excel does.
 */
export function parseFormula(text: string): AstNode {
  const body = text.startsWith('=') ? text.slice(1) : text;
  const offset = text.length - body.length;
  try {
    return new Parser(tokenize(body), body.length).parse();
  } catch (e) {
    if (e instanceof CalcParseError && offset > 0) throw new CalcParseError(e.message, e.position + offset);
    throw e;
  }
}
