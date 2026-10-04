// Logical functions, plus LET / LAMBDA and the LAMBDA helpers (MAP, REDUCE…).

import { scalarsEqual, toBoolean } from '../coerce.ts';
import { mapBinary, mapTernary, topLeft } from '../evaluator.ts';
import type { FnContext, FunctionSpec, Thunk } from '../function-spec.ts';
import { normalizeLocalName } from '../parser.ts';
import {
  type CalcArray,
  type CalcScalar,
  type CalcValue,
  ERRORS,
  isArray,
  isError,
  isLambda,
  makeArray,
} from '../types.ts';
import { arrayOf, forEachValue, nums, scalar, trunc } from './helpers.ts';

const C = 'Logical' as const;

/** Combine AND / OR / XOR inputs: booleans and numbers count, text in ranges is skipped. */
const logicalValues = (args: readonly CalcValue[], ctx: FnContext): boolean[] | CalcScalar => {
  const out: boolean[] = [];
  let error: CalcScalar | undefined;
  forEachValue(args, ctx, (v, fromRange) => {
    if (isError(v)) error = v;
    else if (typeof v === 'boolean') out.push(v);
    else if (typeof v === 'number') out.push(v !== 0);
    else if (typeof v === 'string' && !fromRange) {
      const b = toBoolean(v);
      if (isError(b)) error = b;
      else out.push(b);
    }
    return error === undefined;
  });
  if (error !== undefined) return error;
  return out.length === 0 ? ERRORS.VALUE : out;
};

const logical = (name: string, description: string, fn: (values: boolean[]) => boolean): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(logical1, [logical2], ...)`,
  description,
  minArgs: 1,
  maxArgs: 255,
  args: ['ref'],
  impl: (args, ctx) => {
    const values = logicalValues(args, ctx);
    return Array.isArray(values) ? fn(values) : values;
  },
});

/** Evaluate a branch-selecting argument; arrays make the whole call element-wise. */
const condition = (thunk: Thunk | undefined, ctx: FnContext): CalcScalar | CalcArray => ctx.deref(thunk?.() ?? null);

const ifPick = (c: CalcScalar, x: CalcScalar, y: CalcScalar): CalcScalar => {
  if (isError(c)) return c;
  const b = toBoolean(c);
  if (isError(b)) return b;
  return b ? x : y;
};

const lambdaArg = (v: CalcValue | undefined): v is Extract<CalcValue, { kind: 'lambda' }> => v !== undefined && isLambda(v);

export const LOGICAL_FUNCTIONS: FunctionSpec[] = [
  {
    name: 'IF',
    category: C,
    syntax: 'IF(logical_test, [value_if_true], [value_if_false])',
    description: 'Checks whether a condition is met, and returns one value if TRUE and another if FALSE.',
    minArgs: 2,
    maxArgs: 3,
    lazy: true,
    impl: (args, ctx) => {
      const cond = condition(args[0], ctx);
      if (isArray(cond)) {
        const yes = ctx.deref(args[1]?.() ?? null);
        const no = args[2] === undefined ? false : ctx.deref(args[2]());
        return mapTernary(cond, yes, no, ifPick);
      }
      if (isError(cond)) return cond;
      const b = toBoolean(cond);
      if (isError(b)) return b;
      if (b) return args[1]?.() ?? null;
      return args[2] === undefined ? false : args[2]();
    },
  },
  {
    name: 'IFS',
    category: C,
    syntax: 'IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)',
    description: 'Checks whether one or more conditions are met and returns a value corresponding to the first TRUE condition.',
    minArgs: 2,
    maxArgs: 254,
    lazy: true,
    impl: (args, ctx) => {
      if (args.length % 2 !== 0) return ERRORS.VALUE;
      for (let i = 0; i < args.length; i += 2) {
        const c = topLeft(condition(args[i], ctx));
        if (isError(c)) return c;
        const b = toBoolean(c);
        if (isError(b)) return b;
        if (b) return args[i + 1]?.() ?? null;
      }
      return ERRORS.NA;
    },
  },
  ...(
    [
      ['IFERROR', 'Returns value_if_error if expression is an error and the value of the expression itself otherwise.', (v: CalcScalar) => isError(v)],
      ['IFNA', 'Returns the value you specify if the expression resolves to #N/A, otherwise returns the result of the expression.', (v: CalcScalar) => isError(v) && v.code === '#N/A'],
    ] as const
  ).map(
    ([name, description, caught]): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(value, ${name === 'IFNA' ? 'value_if_na' : 'value_if_error'})`,
      description,
      minArgs: 2,
      maxArgs: 2,
      lazy: true,
      impl: (args, ctx) => {
        const raw = args[0]?.() ?? null;
        const v = ctx.deref(raw);
        if (isArray(v)) {
          if (!v.data.some(caught)) return v;
          const alt = ctx.deref(args[1]?.() ?? null);
          return mapBinary(v, alt, (x, y) => (caught(x) ? (y ?? 0) : x));
        }
        if (caught(v)) return args[1]?.() ?? null;
        return raw;
      },
    }),
  ),
  logical('AND', 'Returns TRUE if all of its arguments are TRUE.', (v) => v.every(Boolean)),
  logical('OR', 'Returns TRUE if any argument is TRUE.', (v) => v.some(Boolean)),
  logical('XOR', 'Returns a logical exclusive OR of all arguments.', (v) => v.filter(Boolean).length % 2 === 1),
  {
    name: 'NOT',
    category: C,
    syntax: 'NOT(logical)',
    description: 'Reverses the logic of its argument.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args) => {
      const b = toBoolean(scalar(args[0]));
      return isError(b) ? b : !b;
    },
  },
  {
    name: 'TRUE',
    category: C,
    syntax: 'TRUE()',
    description: 'Returns the logical value TRUE.',
    minArgs: 0,
    maxArgs: 0,
    impl: () => true,
  },
  {
    name: 'FALSE',
    category: C,
    syntax: 'FALSE()',
    description: 'Returns the logical value FALSE.',
    minArgs: 0,
    maxArgs: 0,
    impl: () => false,
  },
  {
    name: 'SWITCH',
    category: C,
    syntax: 'SWITCH(expression, value1, result1, [default_or_value2, result2], ...)',
    description: 'Evaluates an expression against a list of values and returns the result corresponding to the first matching value.',
    minArgs: 3,
    maxArgs: 254,
    lazy: true,
    impl: (args, ctx) => {
      const target = topLeft(ctx.deref(args[0]?.() ?? null));
      if (isError(target)) return target;
      let i = 1;
      for (; i + 1 < args.length; i += 2) {
        const candidate = topLeft(ctx.deref(args[i]?.() ?? null));
        if (isError(candidate)) return candidate;
        if (scalarsEqual(target, candidate) && typeof target === typeof candidate) return args[i + 1]?.() ?? null;
      }
      return i < args.length ? (args[i]?.() ?? null) : ERRORS.NA;
    },
  },
  {
    name: 'LET',
    category: C,
    syntax: 'LET(name1, name_value1, calculation_or_name2, [name_value2, calculation_or_name3], ...)',
    description: 'Assigns names to calculation results to allow storing intermediate calculations inside a formula.',
    minArgs: 3,
    maxArgs: 253,
    lazy: true,
    impl: (_args, ctx, nodes) => {
      if (nodes.length % 2 !== 1) return ERRORS.VALUE;
      const vars = new Map<string, CalcValue>();
      for (let i = 0; i + 1 < nodes.length; i += 2) {
        const name = nodes[i];
        const value = nodes[i + 1];
        if (name?.type !== 'name' || name.prefix !== undefined || value === undefined) return ERRORS.NAME;
        vars.set(normalizeLocalName(name.name).toUpperCase(), ctx.evaluateWith(value, vars));
      }
      const body = nodes[nodes.length - 1];
      return body === undefined ? ERRORS.VALUE : ctx.evaluateWith(body, vars);
    },
  },
  {
    name: 'LAMBDA',
    category: C,
    syntax: 'LAMBDA([parameter1, parameter2, ...], calculation)',
    description: 'Creates a reusable function from parameters and a calculation.',
    minArgs: 1,
    maxArgs: 254,
    lazy: true,
    impl: (_args, ctx, nodes) => {
      const params: string[] = [];
      for (const p of nodes.slice(0, -1)) {
        if (p.type !== 'name' || p.prefix !== undefined) return ERRORS.VALUE;
        params.push(normalizeLocalName(p.name));
      }
      const body = nodes[nodes.length - 1];
      return body === undefined ? ERRORS.VALUE : ctx.makeLambda(params, body);
    },
  },
  {
    name: 'MAP',
    category: C,
    syntax: 'MAP(array1, [array2], ..., lambda)',
    description: 'Returns an array formed by mapping each value in the array(s) to a new value by applying a LAMBDA.',
    minArgs: 2,
    maxArgs: 254,
    args: ['value'],
    impl: (args, ctx) => {
      const fn = args[args.length - 1];
      if (!lambdaArg(fn)) return ERRORS.VALUE;
      const arrays = args.slice(0, -1).map((a) => ctx.toArray(a));
      const first = arrays[0];
      if (first === undefined) return ERRORS.VALUE;
      return arrayOf(first.rows, first.cols, (r, c) =>
        topLeft(ctx.deref(ctx.callLambda(fn, arrays.map((a) => (r < a.rows && c < a.cols ? (a.data[r * a.cols + c] ?? null) : ERRORS.NA))))),
      );
    },
  },
  {
    name: 'REDUCE',
    category: C,
    syntax: 'REDUCE([initial_value], array, lambda)',
    description: 'Reduces an array to an accumulated value by applying a LAMBDA to each value.',
    minArgs: 3,
    maxArgs: 3,
    args: ['value', 'value', 'value'],
    impl: (args, ctx) => {
      const fn = args[2];
      if (!lambdaArg(fn)) return ERRORS.VALUE;
      let acc: CalcValue = args[0] ?? null;
      for (const v of ctx.toArray(args[1] ?? null).data) acc = ctx.callLambda(fn, [acc, v]);
      return acc;
    },
  },
  {
    name: 'SCAN',
    category: C,
    syntax: 'SCAN([initial_value], array, lambda)',
    description: 'Scans an array by applying a LAMBDA to each value and returns an array that has each intermediate value.',
    minArgs: 3,
    maxArgs: 3,
    args: ['value', 'value', 'value'],
    impl: (args, ctx) => {
      const fn = args[2];
      if (!lambdaArg(fn)) return ERRORS.VALUE;
      const source = ctx.toArray(args[1] ?? null);
      let acc: CalcValue = args[0] ?? null;
      const out: CalcScalar[] = [];
      for (const v of source.data) {
        acc = ctx.callLambda(fn, [acc, v]);
        out.push(topLeft(ctx.deref(acc)));
      }
      return makeArray(source.rows, source.cols, out);
    },
  },
  ...(['BYROW', 'BYCOL'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array, lambda)`,
      description: name === 'BYROW' ? 'Applies a LAMBDA to each row and returns an array of the results.' : 'Applies a LAMBDA to each column and returns an array of the results.',
      minArgs: 2,
      maxArgs: 2,
      args: ['value', 'value'],
      impl: (args, ctx) => {
        const fn = args[1];
        if (!lambdaArg(fn)) return ERRORS.VALUE;
        const a = ctx.toArray(args[0] ?? null);
        const byRow = name === 'BYROW';
        const n = byRow ? a.rows : a.cols;
        const out: CalcScalar[] = [];
        for (let i = 0; i < n; i++) {
          const slice = byRow
            ? makeArray(1, a.cols, a.data.slice(i * a.cols, (i + 1) * a.cols))
            : makeArray(a.rows, 1, Array.from({ length: a.rows }, (_, r) => a.data[r * a.cols + i] ?? null));
          out.push(topLeft(ctx.deref(ctx.callLambda(fn, [slice]))));
        }
        return byRow ? makeArray(n, 1, out) : makeArray(1, n, out);
      },
    }),
  ),
  {
    name: 'MAKEARRAY',
    category: C,
    syntax: 'MAKEARRAY(rows, cols, lambda)',
    description: 'Returns a calculated array of a specified row and column size, by applying a LAMBDA.',
    minArgs: 3,
    maxArgs: 3,
    args: ['scalar', 'scalar', 'value'],
    impl: (args, ctx) => {
      const fn = args[2];
      if (!lambdaArg(fn)) return ERRORS.VALUE;
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const rows = trunc(v[0] ?? 0);
      const cols = trunc(v[1] ?? 0);
      if (rows < 1 || cols < 1) return ERRORS.VALUE;
      return arrayOf(rows, cols, (r, c) => topLeft(ctx.deref(ctx.callLambda(fn, [r + 1, c + 1]))));
    },
  },
];

