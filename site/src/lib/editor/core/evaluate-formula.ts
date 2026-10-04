// Formulas ▸ Evaluate Formula: reduce a formula one step at a time, the way
// Excel's dialog does. Each step replaces the leftmost innermost part that can
// be worked out (a reference, a name, or an operation whose operands are all
// values) with its value; the text is printed back from the tree with that
// next part marked.

import { parseFormula, type AstNode, type CalcScalar } from '../calc/index.ts';
import { renderArea, renderPrefix } from '../calc/address.ts';
import type { CalcArray } from '../calc/types.ts';

type Literal = Extract<AstNode, { type: 'number' | 'string' | 'bool' | 'error' | 'array' | 'missing' }>;

const isLiteral = (n: AstNode): n is Literal =>
  n.type === 'number' || n.type === 'string' || n.type === 'bool' || n.type === 'error' || n.type === 'array' || n.type === 'missing';

function children(n: AstNode): readonly AstNode[] {
  switch (n.type) {
    case 'unary':
    case 'postfix':
      return [n.operand];
    case 'binary':
      return [n.left, n.right];
    case 'call':
      return n.args;
    case 'invoke':
      return [n.callee, ...n.args];
    default:
      return [];
  }
}

/** Above this many cells a range isn't spelled out as an array; its function is evaluated whole instead. */
const MAX_SHOWN_CELLS = 1000;

function isBigRef(n: AstNode): boolean {
  if (n.type !== 'ref') return false;
  const a = n.area;
  return (Math.abs(a.r2 - a.r1) + 1) * (Math.abs(a.c2 - a.c1) + 1) > MAX_SHOWN_CELLS;
}

/** The next node Excel would evaluate: leftmost, innermost, not yet a value. */
export function nextStep(n: AstNode): AstNode | undefined {
  if (isLiteral(n) || isBigRef(n)) return undefined;
  for (const c of children(n)) {
    const inner = nextStep(c);
    if (inner) return inner;
  }
  return n;
}

const PRECEDENCE: Readonly<Record<string, number>> = {
  ':': 9, ' ': 8, ',': 7, '^': 5, '*': 4, '/': 4, '+': 3, '-': 3, '&': 2,
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
};

function scalarText(v: CalcScalar): string {
  if (v === null) return '';
  if (typeof v === 'number') return String(v);
  if (typeof v === 'string') return `"${v.replace(/"/g, '""')}"`;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return v.code;
}

/** Print `n`; `mark` gets the span of `target` in the output. */
export function printNode(n: AstNode, target?: AstNode, mark?: { start: number; end: number }): string {
  let out = '';
  const emit = (node: AstNode, parentPrec: number): void => {
    const start = out.length;
    write(node, parentPrec);
    if (node === target && mark) {
      mark.start = start;
      mark.end = out.length;
    }
  };
  const write = (node: AstNode, parentPrec: number): void => {
    switch (node.type) {
      case 'number':
        out += String(node.value);
        return;
      case 'string':
        out += scalarText(node.value);
        return;
      case 'bool':
        out += node.value ? 'TRUE' : 'FALSE';
        return;
      case 'error':
        out += node.error.code;
        return;
      case 'array': {
        const rows: string[] = [];
        for (let r = 0; r < node.rows; r++) rows.push(node.data.slice(r * node.cols, (r + 1) * node.cols).map(scalarText).join(','));
        out += `{${rows.join(';')}}`;
        return;
      }
      case 'ref':
        out += (node.prefix ? renderPrefix(node.prefix) : '') + renderArea(node.area);
        return;
      case 'name':
        out += (node.prefix ? renderPrefix(node.prefix) : '') + node.name;
        return;
      case 'structured': {
        const parts = [...node.spec.specials, ...(node.spec.col1 ? [node.spec.col2 ? `${node.spec.col1}]:[${node.spec.col2}` : node.spec.col1] : [])];
        out += `${node.table ?? ''}[${parts.map((p) => `[${p}]`).join(',')}]`;
        return;
      }
      case 'unary':
        out += node.op;
        emit(node.operand, 6);
        return;
      case 'postfix':
        emit(node.operand, 6);
        out += node.op;
        return;
      case 'binary': {
        const prec = PRECEDENCE[node.op] ?? 0;
        const paren = prec < parentPrec;
        if (paren) out += '(';
        emit(node.left, prec);
        out += node.op;
        // Same-precedence operators are left-associative, so a right operand at that level needs parentheses.
        emit(node.right, prec + 1);
        if (paren) out += ')';
        return;
      }
      case 'call':
        out += `${node.name}(`;
        node.args.forEach((a, i) => {
          if (i > 0) out += ',';
          emit(a, 0);
        });
        out += ')';
        return;
      case 'invoke':
        emit(node.callee, 10);
        out += '(';
        node.args.forEach((a, i) => {
          if (i > 0) out += ',';
          emit(a, 0);
        });
        out += ')';
        return;
      case 'missing':
        break;
    }
  };
  emit(n, 0);
  return out;
}

function toLiteral(v: CalcScalar | CalcArray): AstNode {
  if (v === null) return { type: 'number', value: 0 };
  if (typeof v === 'number') return { type: 'number', value: v };
  if (typeof v === 'string') return { type: 'string', value: v };
  if (typeof v === 'boolean') return { type: 'bool', value: v };
  if (v.kind === 'error') return { type: 'error', error: v };
  return { type: 'array', rows: v.rows, cols: v.cols, data: v.data };
}

function replace(n: AstNode, target: AstNode, by: AstNode): AstNode {
  if (n === target) return by;
  switch (n.type) {
    case 'unary':
    case 'postfix':
      return { ...n, operand: replace(n.operand, target, by) };
    case 'binary':
      return { ...n, left: replace(n.left, target, by), right: replace(n.right, target, by) };
    case 'call':
      return { ...n, args: n.args.map((a) => replace(a, target, by)) };
    case 'invoke':
      return { ...n, callee: replace(n.callee, target, by), args: n.args.map((a) => replace(a, target, by)) };
    default:
      return n;
  }
}

export interface EvaluationState {
  readonly tree: AstNode;
  /** Printed formula and the span of the part the next step evaluates (empty when done). */
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly done: boolean;
}

function stateOf(tree: AstNode): EvaluationState {
  // A formula that is only a big range never reduces; it is done as it stands.
  const next = nextStep(tree);
  const mark = { start: 0, end: 0 };
  const text = printNode(tree, next, mark);
  return { tree, text, start: mark.start, end: mark.end, done: next === undefined };
}

export function startEvaluation(formula: string): EvaluationState {
  return stateOf(parseFormula(formula));
}

/** One "Evaluate" press. `evaluate` works out a printed sub-formula at the formula's own cell. */
export function stepEvaluation(state: EvaluationState, evaluate: (formula: string) => CalcScalar | CalcArray): EvaluationState {
  const next = nextStep(state.tree);
  if (!next) return state;
  return stateOf(replace(state.tree, next, toLiteral(evaluate(printNode(next)))));
}
