// Tree-walking evaluator with Excel's dynamic-array semantics: operators work
// element-wise over arrays and ranges (broadcasting a single row / column,
// padding mismatched sizes with #N/A), and functions with scalar parameters
// "lift" over array arguments. Recursion here is bounded by formula nesting
// depth, never by the length of a dependency chain — the engine orders cell
// evaluation itself.

import type { AstNode, RefArea, SheetPrefix } from './ast.ts';
import { approxAdd, compareScalars, round15, toNumber, toText } from './coerce.ts';
import type { ArgKind, EagerSpec, FnContext, FunctionSpec, Thunk } from './function-spec.ts';
import type { EvalHost } from './host.ts';
import { CalcParseError } from './lexer.ts';
import { normalizeLocalName, parseFormula } from './parser.ts';
import {
  type Area,
  type CalcArray,
  type CalcError,
  type CalcLambda,
  type CalcRef,
  type CalcScalar,
  type CalcValue,
  ERRORS,
  isArray,
  isError,
  isLambda,
  isRef,
  makeArray,
  MAX_COL,
  MAX_ROW,
  type Scope,
} from './types.ts';

/** Where a formula is being evaluated. */
export interface Frame {
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
  /**
   * Added to the relative parts of every reference. Non-zero for followers of
   * a shared formula, whose tree is the master's.
   */
  readonly dRow: number;
  readonly dCol: number;
  readonly scope: Scope | undefined;
  readonly depth: number;
}

// Names can refer to names and LAMBDAs can recurse; Excel gives up somewhere
// too. This keeps a runaway definition from blowing the JS stack.
const MAX_DEPTH = 256;

// Shared formulas whose relative references run off the grid wrap around, as
// Excel's own relative-reference encoding does.
const wrap = (v: number, max: number): number => ((((v - 1) % max) + max) % max) + 1;

type Scalarish = CalcScalar | CalcArray;

/** The normalised rectangle a written reference denotes once the shared-formula offset is applied. */
export function resolveRect(area: RefArea, dRow: number, dCol: number): { r1: number; c1: number; r2: number; c2: number } {
  const fixedRows = area.kind === 'cols';
  const fixedCols = area.kind === 'rows';
  const r1 = fixedRows || area.r1Abs ? area.r1 : wrap(area.r1 + dRow, MAX_ROW);
  const r2 = fixedRows || area.r2Abs ? area.r2 : wrap(area.r2 + dRow, MAX_ROW);
  const c1 = fixedCols || area.c1Abs ? area.c1 : wrap(area.c1 + dCol, MAX_COL);
  const c2 = fixedCols || area.c2Abs ? area.c2 : wrap(area.c2 + dCol, MAX_COL);
  return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
}

export class Evaluator {
  constructor(
    private readonly host: EvalHost,
    private readonly functions: ReadonlyMap<string, FunctionSpec>,
  ) {}

  /** Evaluate a whole formula: references are read, LAMBDAs without a call become #CALC!. */
  evaluateFormula(node: AstNode, frame: Frame): Scalarish {
    // Excel rounds a formula's last addition or subtraction to 0 when it only
    // cancels to floating-point noise: =0.1+0.2-0.3 is 0, =(0.1+0.2-0.3) is not.
    const v = node.type === 'binary' && (node.op === '+' || node.op === '-') && !node.parenthesized ? this.binary(node.op, node.left, node.right, frame, true) : this.evaluate(node, frame);
    if (isLambda(v)) return ERRORS.CALC;
    return this.deref(v);
  }

  evaluate(node: AstNode, frame: Frame): CalcValue {
    switch (node.type) {
      case 'number':
      case 'string':
      case 'bool':
        return node.value;
      case 'error':
        return node.error;
      case 'missing':
        return null;
      case 'array':
        return makeArray(node.rows, node.cols, node.data.slice());
      case 'ref':
        return this.reference(node.prefix, node.area, frame);
      case 'name':
        return this.name(node.name, node.prefix, frame);
      case 'structured': {
        const area = this.host.structuredArea(node.table, node.spec, frame.sheet, frame.row, frame.col);
        return isError(area) ? area : { kind: 'ref', areas: [area] };
      }
      case 'unary':
        return this.unary(node.op, this.evaluate(node.operand, frame), frame);
      case 'postfix': {
        const operand = this.evaluate(node.operand, frame);
        if (node.op === '%') return mapUnary(this.deref(operand), (x) => arith(x, 100, (a, b) => a / b));
        return this.spill(operand);
      }
      case 'binary':
        return this.binary(node.op, node.left, node.right, frame);
      case 'call':
        return this.call(node.name, node.args, frame);
      case 'invoke': {
        const callee = this.evaluate(node.callee, frame);
        if (!isLambda(callee)) return isError(callee) ? callee : ERRORS.VALUE;
        return this.invokeLambda(callee, node.args.map((a) => this.evaluate(a, frame)), frame);
      }
    }
  }

  // ---- references --------------------------------------------------------

  private reference(prefix: SheetPrefix | undefined, area: RefArea, frame: Frame): CalcValue {
    const sheets = this.sheetsFor(prefix, frame);
    if (isError(sheets)) return sheets;
    const rect = resolveRect(area, frame.dRow, frame.dCol);
    return { kind: 'ref', areas: sheets.map((sheet) => ({ sheet, r1: rect.r1, c1: rect.c1, r2: rect.r2, c2: rect.c2 })) };
  }

  private sheetsFor(prefix: SheetPrefix | undefined, frame: Frame): string[] | CalcError {
    if (prefix === undefined) return [frame.sheet];
    if (prefix.external !== undefined) return ERRORS.REF;
    if (prefix.sheet2 !== undefined) return this.host.sheetsBetween(prefix.sheet, prefix.sheet2) ?? ERRORS.REF;
    const sheet = this.host.sheetName(prefix.sheet);
    return sheet === undefined ? ERRORS.REF : [sheet];
  }

  private name(name: string, prefix: SheetPrefix | undefined, frame: Frame): CalcValue {
    const key = name.toUpperCase();
    if (prefix === undefined) {
      for (let s = frame.scope; s !== undefined; s = s.parent) {
        const v = s.vars.get(key);
        if (v !== undefined) return v;
      }
    }
    let scopeSheet = frame.sheet;
    if (prefix !== undefined) {
      const sheets = this.sheetsFor(prefix, frame);
      if (isError(sheets)) return sheets;
      scopeSheet = sheets[0] ?? frame.sheet;
    }
    const body = this.host.definedName(name, scopeSheet);
    if (body === undefined) return ERRORS.NAME;
    if (frame.depth > MAX_DEPTH) return ERRORS.REF;
    // Relative references inside a name are stored relative to A1 and follow
    // the cell that uses the name.
    return this.evaluate(body, {
      sheet: frame.sheet,
      row: frame.row,
      col: frame.col,
      dRow: frame.row - 1,
      dCol: frame.col - 1,
      scope: undefined,
      depth: frame.depth + 1,
    });
  }

  private spill(operand: CalcValue): CalcValue {
    if (isError(operand)) return operand;
    if (!isRef(operand)) return ERRORS.REF;
    const a = operand.areas[0];
    if (operand.areas.length !== 1 || a === undefined || a.r1 !== a.r2 || a.c1 !== a.c2) return ERRORS.REF;
    const area = this.host.spillArea(a.sheet, a.r1, a.c1);
    return area === undefined ? ERRORS.REF : { kind: 'ref', areas: [area] };
  }

  /** Read a reference; single cells become scalars, everything else an array. */
  deref(v: CalcValue): Scalarish {
    if (!isRef(v)) return isLambda(v) ? ERRORS.VALUE : v;
    if (v.areas.length !== 1) return ERRORS.VALUE;
    const a = v.areas[0];
    if (a === undefined) return ERRORS.REF;
    if (a.r1 === a.r2 && a.c1 === a.c2) return this.host.cellValue(a.sheet, a.r1, a.c1);
    return this.host.areaValues(a);
  }

  // ---- operators ---------------------------------------------------------

  private unary(op: '-' | '+' | '@', v: CalcValue, frame: Frame): CalcValue {
    if (op === '+') return v;
    if (op === '-') return mapUnary(this.deref(v), (x) => arith(x, -1, (a, b) => a * b));
    return this.implicitIntersection(v, frame);
  }

  /** `@`: the single cell of a range in the formula's row or column, as pre-dynamic-array Excel did. */
  implicitIntersection(v: CalcValue, frame: Frame): CalcValue {
    if (isArray(v)) return v.data[0] ?? null;
    if (!isRef(v)) return v;
    const a = v.areas[0];
    if (v.areas.length !== 1 || a === undefined) return ERRORS.VALUE;
    if (a.r1 === a.r2 && a.c1 === a.c2) return this.host.cellValue(a.sheet, a.r1, a.c1);
    const rowOk = frame.row >= a.r1 && frame.row <= a.r2;
    const colOk = frame.col >= a.c1 && frame.col <= a.c2;
    if (a.c1 === a.c2 && rowOk) return this.host.cellValue(a.sheet, frame.row, a.c1);
    if (a.r1 === a.r2 && colOk) return this.host.cellValue(a.sheet, a.r1, frame.col);
    if (rowOk && colOk && a.sheet === frame.sheet) return this.host.cellValue(a.sheet, frame.row, frame.col);
    return ERRORS.VALUE;
  }

  private binary(op: string, leftNode: AstNode, rightNode: AstNode, frame: Frame, final = false): CalcValue {
    const left = this.evaluate(leftNode, frame);
    const right = this.evaluate(rightNode, frame);
    switch (op) {
      case ':':
        return rangeOp(left, right);
      case ' ':
        return intersectOp(left, right);
      case ',':
        return unionOp(left, right);
      case '&':
        return mapBinary(this.deref(left), this.deref(right), concat);
      case '=':
      case '<>':
      case '<':
      case '>':
      case '<=':
      case '>=':
        return mapBinary(this.deref(left), this.deref(right), (a, b) => compareOp(op, a, b));
      default:
        return mapBinary(this.deref(left), this.deref(right), (a, b) => arithmeticOp(op, a, b, final));
    }
  }

  // ---- calls -------------------------------------------------------------

  private call(name: string, argNodes: readonly AstNode[], frame: Frame): CalcValue {
    const spec = this.functions.get(name);
    if (spec === undefined) {
      // Not a built-in: a LET / LAMBDA variable or a defined name holding a LAMBDA.
      // A call to a variable is stored with its `_xlpm.` prefix too (`_xlpm.f(4)`).
      const target = this.name(normalizeLocalName(name), undefined, frame);
      if (isLambda(target)) return this.invokeLambda(target, argNodes.map((a) => this.evaluate(a, frame)), frame);
      return ERRORS.NAME;
    }
    if (argNodes.length < spec.minArgs || argNodes.length > spec.maxArgs) return ERRORS.VALUE;
    const ctx = this.context(frame);
    if (spec.lazy === true) {
      const thunks: Thunk[] = argNodes.map((n) => () => this.evaluate(n, frame));
      return spec.impl(thunks, ctx, argNodes);
    }
    const raw = argNodes.map((n) => this.evaluate(n, frame));
    return this.callEager(spec, raw, ctx);
  }

  private callEager(spec: EagerSpec, raw: CalcValue[], ctx: FnContext): CalcValue {
    const kinds = argKinds(spec, raw.length);
    let lift: { rows: number; cols: number } | undefined;
    const prepared: CalcValue[] = raw.map((v, i) => {
      const kind = kinds[i] ?? 'scalar';
      if (kind === 'ref') return v;
      const d = isLambda(v) ? v : this.deref(v);
      if (kind === 'scalar' && isArray(d)) {
        lift = lift === undefined ? { rows: d.rows, cols: d.cols } : broadcastShape(lift, d);
      }
      return d;
    });
    if (lift === undefined) return this.invokeScalar(spec, prepared, kinds, ctx);
    const { rows, cols } = lift;
    const out: CalcScalar[] = new Array<CalcScalar>(rows * cols);
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const args = prepared.map((v, i) => ((kinds[i] ?? 'scalar') === 'scalar' && isArray(v) ? elementAt(v, r, c) : v));
        const result = this.invokeScalar(spec, args, kinds, ctx);
        out[r * cols + c] = topLeft(this.deref(result));
      }
    }
    return makeArray(rows, cols, out);
  }

  private invokeScalar(spec: EagerSpec, args: CalcValue[], kinds: readonly ArgKind[], ctx: FnContext): CalcValue {
    if (spec.acceptsErrors !== true) {
      for (let i = 0; i < args.length; i++) {
        const v = args[i];
        if ((kinds[i] ?? 'scalar') === 'scalar' && v !== undefined && isError(v)) return v;
      }
    }
    return spec.impl(args, ctx);
  }

  private invokeLambda(fn: CalcLambda, args: readonly CalcValue[], frame: Frame): CalcValue {
    if (args.length > fn.params.length) return ERRORS.VALUE;
    if (frame.depth > MAX_DEPTH) return ERRORS.NUM;
    const vars = new Map<string, CalcValue>();
    fn.params.forEach((p, i) => vars.set(p.toUpperCase(), args[i] ?? null));
    return this.evaluate(fn.body, { ...frame, scope: { vars, parent: fn.scope }, depth: frame.depth + 1 });
  }

  private context(frame: Frame): FnContext {
    return {
      host: this.host,
      sheet: frame.sheet,
      row: frame.row,
      col: frame.col,
      deref: (v) => this.deref(v),
      toArray: (v) => {
        const d = this.deref(v);
        return isArray(d) ? d : makeArray(1, 1, [d]);
      },
      callLambda: (fn, args) => this.invokeLambda(fn, args, frame),
      referenceFromText: (text, a1) => this.referenceFromText(text, a1, frame),
      evaluateWith: (node, vars) =>
        this.evaluate(node, { ...frame, scope: { vars: new Map(vars), parent: frame.scope }, depth: frame.depth + 1 }),
      makeLambda: (params, body) => ({ kind: 'lambda', params, body, scope: frame.scope }),
    };
  }

  private referenceFromText(text: string, a1: boolean, frame: Frame): CalcRef | undefined {
    const trimmed = text.trim();
    if (!a1) return r1c1Reference(trimmed, frame, (p) => this.sheetsFor(p, frame));
    let node: AstNode;
    try {
      node = parseFormula(trimmed);
    } catch (e) {
      // INDIRECT of text that is not a reference is #REF!, not a formula error.
      if (e instanceof CalcParseError) return undefined;
      throw e;
    }
    if (node.type !== 'ref' && node.type !== 'name' && node.type !== 'structured') return undefined;
    const v = this.evaluate(node, { ...frame, dRow: 0, dCol: 0 });
    return isRef(v) ? v : undefined;
  }
}

// ---- helpers ---------------------------------------------------------------

const argKinds = (spec: EagerSpec, count: number): ArgKind[] => {
  const lead = spec.args ?? [];
  const rest = spec.rest ?? (lead.length > 0 ? [lead[lead.length - 1] ?? 'scalar'] : ['scalar']);
  const kinds: ArgKind[] = [];
  for (let i = 0; i < count; i++) {
    kinds.push(i < lead.length ? (lead[i] ?? 'scalar') : (rest[(i - lead.length) % rest.length] ?? 'scalar'));
  }
  return kinds;
};

const broadcastShape = (a: { rows: number; cols: number }, b: { rows: number; cols: number }): { rows: number; cols: number } => ({
  rows: a.rows === 1 ? b.rows : b.rows === 1 ? a.rows : Math.max(a.rows, b.rows),
  cols: a.cols === 1 ? b.cols : b.cols === 1 ? a.cols : Math.max(a.cols, b.cols),
});

/** Element (r, c) of `a` under broadcasting; positions past a non-broadcast edge are #N/A. */
const elementAt = (a: CalcArray, r: number, c: number): CalcScalar => {
  const rr = a.rows === 1 ? 0 : r;
  const cc = a.cols === 1 ? 0 : c;
  if (rr >= a.rows || cc >= a.cols) return ERRORS.NA;
  return a.data[rr * a.cols + cc] ?? null;
};

export function topLeft(v: Scalarish): CalcScalar {
  return isArray(v) ? (v.data[0] ?? null) : v;
}

export function mapUnary(v: Scalarish, fn: (x: CalcScalar) => CalcScalar): Scalarish {
  if (!isArray(v)) return fn(v);
  return makeArray(v.rows, v.cols, v.data.map(fn));
}

export function mapBinary(a: Scalarish, b: Scalarish, fn: (x: CalcScalar, y: CalcScalar) => CalcScalar): Scalarish {
  if (!isArray(a) && !isArray(b)) return fn(a, b);
  const shapeA = isArray(a) ? a : { rows: 1, cols: 1 };
  const shapeB = isArray(b) ? b : { rows: 1, cols: 1 };
  const { rows, cols } = broadcastShape(shapeA, shapeB);
  const out: CalcScalar[] = new Array<CalcScalar>(rows * cols);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const x = isArray(a) ? elementAt(a, r, c) : a;
      const y = isArray(b) ? elementAt(b, r, c) : b;
      out[r * cols + c] = fn(x, y);
    }
  }
  return makeArray(rows, cols, out);
}

export function mapTernary(
  a: Scalarish,
  b: Scalarish,
  c: Scalarish,
  fn: (x: CalcScalar, y: CalcScalar, z: CalcScalar) => CalcScalar,
): Scalarish {
  const arrays = [a, b, c].filter(isArray);
  if (arrays.length === 0) return fn(topLeft(a), topLeft(b), topLeft(c));
  const { rows, cols } = arrays.reduce<{ rows: number; cols: number }>((s, x) => broadcastShape(s, x), { rows: 1, cols: 1 });
  const at = (v: Scalarish, r: number, col: number): CalcScalar => (isArray(v) ? elementAt(v, r, col) : v);
  const out: CalcScalar[] = new Array<CalcScalar>(rows * cols);
  for (let r = 0; r < rows; r++) for (let col = 0; col < cols; col++) out[r * cols + col] = fn(at(a, r, col), at(b, r, col), at(c, r, col));
  return makeArray(rows, cols, out);
}

const arith = (x: CalcScalar, y: number, fn: (a: number, b: number) => number): CalcScalar => {
  const n = toNumber(x);
  if (isError(n)) return n;
  return fn(n, y);
};

function arithmeticOp(op: string, x: CalcScalar, y: CalcScalar, final = false): CalcScalar {
  if (isError(x)) return x;
  if (isError(y)) return y;
  const a = toNumber(x);
  if (isError(a)) return a;
  const b = toNumber(y);
  if (isError(b)) return b;
  let r: number;
  switch (op) {
    case '+':
      r = final ? approxAdd(a, b) : a + b;
      break;
    case '-':
      r = final ? approxAdd(a, -b) : a - b;
      break;
    case '*':
      r = a * b;
      break;
    case '/':
      if (b === 0) return ERRORS.DIV0;
      r = a / b;
      break;
    case '^':
      return power(a, b);
    default:
      return ERRORS.VALUE;
  }
  return Number.isFinite(r) ? r : ERRORS.NUM;
}

/** `a ^ b` and POWER(a, b). */
export function power(a: number, b: number): number | CalcError {
  if (a === 0 && b === 0) return ERRORS.NUM;
  if (a === 0 && b < 0) return ERRORS.DIV0;
  let r = a ** b;
  // Excel takes odd roots of negatives, (-8)^(1/3) = -2, where IEEE pow gives NaN.
  if (a < 0 && !Number.isInteger(b)) {
    const root = round15(1 / b);
    r = Number.isInteger(root) && root % 2 !== 0 ? -((-a) ** b) : Number.NaN;
  }
  return Number.isFinite(r) ? r : ERRORS.NUM;
}

function concat(x: CalcScalar, y: CalcScalar): CalcScalar {
  const a = toText(x);
  if (isError(a)) return a;
  const b = toText(y);
  if (isError(b)) return b;
  return a + b;
}

function compareOp(op: string, x: CalcScalar, y: CalcScalar): CalcScalar {
  if (isError(x)) return x;
  if (isError(y)) return y;
  const c = compareScalars(x, y);
  switch (op) {
    case '=':
      return c === 0;
    case '<>':
      return c !== 0;
    case '<':
      return c < 0;
    case '>':
      return c > 0;
    case '<=':
      return c <= 0;
    default:
      return c >= 0;
  }
}

const singleArea = (v: CalcValue): Area | CalcError => {
  if (isError(v)) return v;
  if (!isRef(v) || v.areas.length !== 1) return ERRORS.VALUE;
  return v.areas[0] ?? ERRORS.REF;
};

/** `A1:B2`, `A1:INDEX(…)`: the bounding box of two references on one sheet. */
function rangeOp(left: CalcValue, right: CalcValue): CalcValue {
  const a = singleArea(left);
  if (isError(a)) return a;
  const b = singleArea(right);
  if (isError(b)) return b;
  if (a.sheet !== b.sheet) return ERRORS.VALUE;
  return {
    kind: 'ref',
    areas: [{ sheet: a.sheet, r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) }],
  };
}

function intersectOp(left: CalcValue, right: CalcValue): CalcValue {
  const a = singleArea(left);
  if (isError(a)) return a;
  const b = singleArea(right);
  if (isError(b)) return b;
  const r1 = Math.max(a.r1, b.r1);
  const r2 = Math.min(a.r2, b.r2);
  const c1 = Math.max(a.c1, b.c1);
  const c2 = Math.min(a.c2, b.c2);
  if (a.sheet !== b.sheet || r1 > r2 || c1 > c2) return ERRORS.NULL;
  return { kind: 'ref', areas: [{ sheet: a.sheet, r1, c1, r2, c2 }] };
}

function unionOp(left: CalcValue, right: CalcValue): CalcValue {
  if (isError(left)) return left;
  if (isError(right)) return right;
  if (!isRef(left) || !isRef(right)) return ERRORS.VALUE;
  return { kind: 'ref', areas: [...left.areas, ...right.areas] };
}

const R1C1 = /^(?:(.+)!)?R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?::R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?)?$/i;

/** INDIRECT(…, FALSE): `R2C3`, `R[-1]C`, `Sheet2!R1C1:R2C2`. */
function r1c1Reference(
  text: string,
  frame: Frame,
  sheetsFor: (p: SheetPrefix | undefined) => string[] | CalcError,
): CalcRef | undefined {
  const m = R1C1.exec(text);
  if (m === null) return undefined;
  const coord = (part: string | undefined, base: number): number =>
    part === undefined ? base : part.startsWith('[') ? base + Number(part.slice(1, -1)) : Number(part);
  let prefix: SheetPrefix | undefined;
  if (m[1] !== undefined) {
    const quoted = /^'(.*)'$/.exec(m[1]);
    prefix = { sheet: quoted !== null ? (quoted[1] ?? '').replace(/''/g, "'") : m[1] };
  }
  const sheets = sheetsFor(prefix);
  if (isError(sheets)) return undefined;
  const r1 = coord(m[2], frame.row);
  const c1 = coord(m[3], frame.col);
  const r2 = m[4] !== undefined || m[5] !== undefined ? coord(m[4], frame.row) : r1;
  const c2 = m[4] !== undefined || m[5] !== undefined ? coord(m[5], frame.col) : c1;
  if ([r1, r2].some((r) => r < 1 || r > MAX_ROW) || [c1, c2].some((c) => c < 1 || c > MAX_COL)) return undefined;
  return {
    kind: 'ref',
    areas: sheets.map((sheet) => ({ sheet, r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) })),
  };
}
