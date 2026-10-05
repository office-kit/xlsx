// Math & Trig functions.

import { round15 } from '../coerce.ts';
import { power } from '../evaluator.ts';
import type { FnContext, FunctionSpec } from '../function-spec.ts';
import { type CalcError, type CalcScalar, type CalcValue, ERRORS, isError, isRef, makeArray } from '../types.ts';
import { arrayOf, checked, collectNumbers, num, nums, optBool, str, sum, trunc, variance } from './helpers.ts';
import { combin, kth, median, modeSingle, percentileExc, percentileInc, stdev } from './stats.ts';

const C = 'Math & Trig' as const;

type RoundMode = 'half' | 'up' | 'down';

/**
 * Excel rounds the decimal value it displays, not the binary double: 2.675
 * is stored as 2.67499999…, yet ROUND(2.675, 2) is 2.68. Scaling and then
 * snapping to 15 significant digits recovers the decimal Excel sees.
 */
export function roundTo(x: number, digits: number, mode: RoundMode): number {
  const d = trunc(digits);
  const apply = (v: number): number => (mode === 'half' ? Math.floor(v + 0.5) : mode === 'up' ? Math.ceil(v) : Math.floor(v));
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(x);
  if (d >= 0) {
    if (d > 15 || ax * 10 ** d >= 1e15) return x;
    const f = 10 ** d;
    return (sign * apply(round15(ax * f))) / f;
  }
  const f = 10 ** -d;
  return sign * apply(round15(ax / f)) * f;
}

const unary = (name: string, description: string, fn: (x: number) => number | CalcError, syntax = `${name}(number)`): FunctionSpec => ({
  name,
  category: C,
  syntax,
  description,
  minArgs: 1,
  maxArgs: 1,
  impl: (args, ctx) => {
    const x = num(args[0], ctx);
    if (isError(x)) return x;
    const r = fn(x);
    return isError(r) ? r : checked(r);
  },
});

const rounding = (name: string, description: string, mode: RoundMode): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(number, num_digits)`,
  description,
  minArgs: 2,
  maxArgs: 2,
  impl: (args, ctx) => {
    const v = nums(args, ctx, [undefined, undefined]);
    if (isError(v)) return v;
    return roundTo(v[0] ?? 0, v[1] ?? 0, mode);
  },
});

const gcd2 = (a: number, b: number): number => {
  let x = a;
  let y = b;
  while (y !== 0) [x, y] = [y, x % y];
  return x;
};

const integerArgs = (args: CalcValue[], ctx: FnContext): number[] | CalcError => {
  const values = collectNumbers(args, ctx);
  if (isError(values)) return values;
  const out: number[] = [];
  for (const v of values) {
    if (v < 0) return ERRORS.NUM;
    out.push(Math.floor(v));
  }
  return out;
};

const factorial = (n: number): number => {
  let r = 1;
  for (let i = 2; i <= n; i++) r *= i;
  return r;
};

/** Values seen by SUBTOTAL / AGGREGATE, honouring their skip rules. */
const subtotalValues = (
  refs: readonly CalcValue[],
  ctx: FnContext,
  opts: { hidden: boolean; nested: boolean; errors: boolean },
): CalcScalar[] | CalcError => {
  const out: CalcScalar[] = [];
  for (const ref of refs) {
    if (isError(ref)) return ref;
    if (!isRef(ref)) {
      for (const v of ctx.toArray(ref).data) {
        if (isError(v) && opts.errors) continue;
        out.push(v);
      }
      continue;
    }
    for (const area of ref.areas) {
      const values = ctx.host.areaValues(area);
      for (let i = 0; i < values.data.length; i++) {
        const v = values.data[i] ?? null;
        if (v === null) continue;
        const row = area.r1 + Math.floor(i / values.cols);
        const col = area.c1 + (i % values.cols);
        if (opts.hidden && ctx.host.isRowHidden(area.sheet, row)) continue;
        if (opts.nested && ctx.host.isSubtotalCell(area.sheet, row, col)) continue;
        if (isError(v) && opts.errors) continue;
        out.push(v);
      }
    }
  }
  return out;
};

const aggregateBy = (fn: number, values: readonly CalcScalar[]): CalcValue => {
  const err = values.find(isError);
  if (err !== undefined && fn !== 3) return err;
  const numbers = values.filter((v): v is number => typeof v === 'number');
  const product = (): number => (numbers.length === 0 ? 0 : numbers.reduce((a, b) => a * b, 1));
  switch (fn) {
    case 1:
      return numbers.length === 0 ? ERRORS.DIV0 : sum(numbers) / numbers.length;
    case 2:
      return numbers.length;
    case 3:
      return values.filter((v) => v !== null).length;
    case 4:
      return numbers.length === 0 ? 0 : Math.max(...numbers);
    case 5:
      return numbers.length === 0 ? 0 : Math.min(...numbers);
    case 6:
      return product();
    case 7:
      return stdev(numbers, true);
    case 8:
      return stdev(numbers, false);
    case 9:
      return sum(numbers);
    case 10:
      return variance(numbers, true);
    case 11:
      return variance(numbers, false);
    case 12:
      return median(numbers);
    case 13:
      return modeSingle(numbers);
    default:
      return ERRORS.VALUE;
  }
};

const matrix = (v: CalcValue, ctx: FnContext): { rows: number; cols: number; m: number[] } | CalcError => {
  const a = ctx.toArray(v);
  const m: number[] = [];
  for (const x of a.data) {
    if (isError(x)) return x;
    if (typeof x !== 'number') return ERRORS.VALUE;
    m.push(x);
  }
  return { rows: a.rows, cols: a.cols, m };
};

/** Gaussian elimination with partial pivoting; returns the determinant and, optionally, the inverse. */
const eliminate = (n: number, m: readonly number[], invert: boolean): { det: number; inv: number[] } => {
  const a = m.slice();
  const inv: number[] = invert ? Array.from({ length: n * n }, (_, i) => (Math.floor(i / n) === i % n ? 1 : 0)) : [];
  let det = 1;
  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(a[r * n + col] ?? 0) > Math.abs(a[pivot * n + col] ?? 0)) pivot = r;
    const pv = a[pivot * n + col] ?? 0;
    if (pv === 0) return { det: 0, inv: [] };
    if (pivot !== col) {
      det = -det;
      for (let c = 0; c < n; c++) {
        [a[col * n + c], a[pivot * n + c]] = [a[pivot * n + c] ?? 0, a[col * n + c] ?? 0];
        if (invert) [inv[col * n + c], inv[pivot * n + c]] = [inv[pivot * n + c] ?? 0, inv[col * n + c] ?? 0];
      }
    }
    det *= pv;
    for (let c = 0; c < n; c++) {
      a[col * n + c] = (a[col * n + c] ?? 0) / pv;
      if (invert) inv[col * n + c] = (inv[col * n + c] ?? 0) / pv;
    }
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const f = a[r * n + col] ?? 0;
      if (f === 0) continue;
      for (let c = 0; c < n; c++) {
        a[r * n + c] = (a[r * n + c] ?? 0) - f * (a[col * n + c] ?? 0);
        if (invert) inv[r * n + c] = (inv[r * n + c] ?? 0) - f * (inv[col * n + c] ?? 0);
      }
    }
  }
  return { det, inv };
};

const ROMAN: ReadonlyArray<[number, string]> = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];

const pairwise = (name: string, description: string, fn: (x: number, y: number) => number): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(array_x, array_y)`,
  description,
  minArgs: 2,
  maxArgs: 2,
  args: ['value', 'value'],
  impl: (args, ctx) => {
    const a = ctx.toArray(args[0] ?? null);
    const b = ctx.toArray(args[1] ?? null);
    if (a.data.length !== b.data.length) return ERRORS.NA;
    let total = 0;
    for (let i = 0; i < a.data.length; i++) {
      const x = a.data[i] ?? null;
      const y = b.data[i] ?? null;
      if (isError(x)) return x;
      if (isError(y)) return y;
      if (typeof x === 'number' && typeof y === 'number') total += fn(x, y);
    }
    return total;
  },
});

export const MATH_FUNCTIONS: FunctionSpec[] = [
  {
    name: 'SUM',
    category: C,
    syntax: 'SUM(number1, [number2], ...)',
    description: 'Adds all the numbers in a range of cells.',
    minArgs: 1,
    maxArgs: 255,
    args: ['ref'],
    impl: (args, ctx) => {
      const v = collectNumbers(args, ctx);
      return isError(v) ? v : sum(v);
    },
  },
  {
    name: 'PRODUCT',
    category: C,
    syntax: 'PRODUCT(number1, [number2], ...)',
    description: 'Multiplies all the numbers given as arguments.',
    minArgs: 1,
    maxArgs: 255,
    args: ['ref'],
    impl: (args, ctx) => {
      const v = collectNumbers(args, ctx);
      if (isError(v)) return v;
      return v.length === 0 ? 0 : checked(v.reduce((a, b) => a * b, 1));
    },
  },
  {
    name: 'SUMSQ',
    category: C,
    syntax: 'SUMSQ(number1, [number2], ...)',
    description: 'Returns the sum of the squares of the arguments.',
    minArgs: 1,
    maxArgs: 255,
    args: ['ref'],
    impl: (args, ctx) => {
      const v = collectNumbers(args, ctx);
      return isError(v) ? v : sum(v.map((x) => x * x));
    },
  },
  {
    name: 'SUMPRODUCT',
    category: C,
    syntax: 'SUMPRODUCT(array1, [array2], [array3], ...)',
    description: 'Returns the sum of the products of corresponding array components.',
    minArgs: 1,
    maxArgs: 255,
    args: ['value'],
    impl: (args, ctx) => {
      const arrays = args.map((a) => ctx.toArray(a));
      const first = arrays[0];
      if (first === undefined) return ERRORS.VALUE;
      if (arrays.some((a) => a.rows !== first.rows || a.cols !== first.cols)) return ERRORS.VALUE;
      let total = 0;
      for (let i = 0; i < first.data.length; i++) {
        let p = 1;
        for (const a of arrays) {
          const v = a.data[i] ?? null;
          if (isError(v)) return v;
          p *= typeof v === 'number' ? v : 0;
        }
        total += p;
      }
      return total;
    },
  },
  pairwise('SUMX2MY2', 'Returns the sum of the difference of squares of corresponding values.', (x, y) => x * x - y * y),
  pairwise('SUMX2PY2', 'Returns the sum of the sum of squares of corresponding values.', (x, y) => x * x + y * y),
  pairwise('SUMXMY2', 'Returns the sum of squares of differences of corresponding values.', (x, y) => (x - y) ** 2),
  rounding('ROUND', 'Rounds a number to a specified number of digits.', 'half'),
  rounding('ROUNDUP', 'Rounds a number up, away from zero.', 'up'),
  rounding('ROUNDDOWN', 'Rounds a number down, toward zero.', 'down'),
  {
    name: 'MROUND',
    category: C,
    syntax: 'MROUND(number, multiple)',
    description: 'Returns a number rounded to the desired multiple.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const [n = 0, m = 0] = v;
      if (m === 0) return 0;
      if (n * m < 0) return ERRORS.NUM;
      return roundTo(n / m, 0, 'half') * m;
    },
  },
  {
    name: 'CEILING',
    category: C,
    syntax: 'CEILING(number, significance)',
    description: 'Rounds a number up to the nearest multiple of significance.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const [n = 0, s = 0] = v;
      if (s === 0) return 0;
      if (n > 0 && s < 0) return ERRORS.NUM;
      return Math.ceil(round15(n / s)) * s;
    },
  },
  ...(['CEILING.MATH', 'CEILING.PRECISE', 'ISO.CEILING'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: name === 'CEILING.MATH' ? 'CEILING.MATH(number, [significance], [mode])' : `${name}(number, [significance])`,
      description: 'Rounds a number up to the nearest integer or multiple of significance.',
      minArgs: 1,
      maxArgs: name === 'CEILING.MATH' ? 3 : 2,
      impl: (args, ctx) => {
        const v = nums(args, ctx, [undefined, 1, 0]);
        if (isError(v)) return v;
        const [n = 0, sRaw = 1, mode = 0] = v;
        const s = Math.abs(sRaw);
        if (s === 0) return 0;
        if (n < 0 && mode !== 0 && name === 'CEILING.MATH') return -Math.ceil(round15(-n / s)) * s;
        return Math.ceil(round15(n / s)) * s;
      },
    }),
  ),
  {
    name: 'FLOOR',
    category: C,
    syntax: 'FLOOR(number, significance)',
    description: 'Rounds a number down, toward zero, to the nearest multiple of significance.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const [n = 0, s = 0] = v;
      if (s === 0) return n === 0 ? 0 : ERRORS.DIV0;
      if (n > 0 && s < 0) return ERRORS.NUM;
      return Math.floor(round15(n / s)) * s;
    },
  },
  ...(['FLOOR.MATH', 'FLOOR.PRECISE'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: name === 'FLOOR.MATH' ? 'FLOOR.MATH(number, [significance], [mode])' : 'FLOOR.PRECISE(number, [significance])',
      description: 'Rounds a number down to the nearest integer or multiple of significance.',
      minArgs: 1,
      maxArgs: name === 'FLOOR.MATH' ? 3 : 2,
      impl: (args, ctx) => {
        const v = nums(args, ctx, [undefined, 1, 0]);
        if (isError(v)) return v;
        const [n = 0, sRaw = 1, mode = 0] = v;
        const s = Math.abs(sRaw);
        if (s === 0) return 0;
        if (n < 0 && mode !== 0 && name === 'FLOOR.MATH') return -Math.floor(round15(-n / s)) * s;
        return Math.floor(round15(n / s)) * s;
      },
    }),
  ),
  unary('INT', 'Rounds a number down to the nearest integer.', Math.floor),
  {
    name: 'TRUNC',
    category: C,
    syntax: 'TRUNC(number, [num_digits])',
    description: 'Truncates a number to an integer or to the given number of digits.',
    minArgs: 1,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, 0]);
      if (isError(v)) return v;
      return roundTo(v[0] ?? 0, v[1] ?? 0, 'down');
    },
  },
  unary('ABS', 'Returns the absolute value of a number.', Math.abs),
  unary('SIGN', 'Returns the sign of a number.', (x) => Math.sign(x)),
  {
    name: 'MOD',
    category: C,
    syntax: 'MOD(number, divisor)',
    description: 'Returns the remainder after a number is divided by a divisor.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const [n = 0, d = 0] = v;
      if (d === 0) return ERRORS.DIV0;
      const r = n - d * Math.floor(n / d);
      // Floating division can land a hair past a whole multiple; Excel reports 0 there.
      return Math.abs(r - d) < Math.abs(d) * 1e-15 ? 0 : r;
    },
  },
  {
    name: 'QUOTIENT',
    category: C,
    syntax: 'QUOTIENT(numerator, denominator)',
    description: 'Returns the integer portion of a division.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const [n = 0, d = 0] = v;
      return d === 0 ? ERRORS.DIV0 : trunc(n / d);
    },
  },
  {
    name: 'POWER',
    category: C,
    syntax: 'POWER(number, power)',
    description: 'Returns the result of a number raised to a power.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const [b = 0, e = 0] = v;
      return power(b, e);
    },
  },
  unary('SQRT', 'Returns a positive square root.', (x) => (x < 0 ? ERRORS.NUM : Math.sqrt(x))),
  unary('SQRTPI', 'Returns the square root of (number * pi).', (x) => (x < 0 ? ERRORS.NUM : Math.sqrt(x * Math.PI))),
  unary('EXP', 'Returns e raised to the power of a given number.', Math.exp),
  unary('LN', 'Returns the natural logarithm of a number.', (x) => (x <= 0 ? ERRORS.NUM : Math.log(x))),
  unary('LOG10', 'Returns the base-10 logarithm of a number.', (x) => (x <= 0 ? ERRORS.NUM : Math.log10(x))),
  {
    name: 'LOG',
    category: C,
    syntax: 'LOG(number, [base])',
    description: 'Returns the logarithm of a number to the base you specify.',
    minArgs: 1,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, 10]);
      if (isError(v)) return v;
      const [x = 0, b = 10] = v;
      if (x <= 0 || b <= 0) return ERRORS.NUM;
      if (b === 1) return ERRORS.DIV0;
      return b === 10 ? Math.log10(x) : Math.log(x) / Math.log(b);
    },
  },
  {
    name: 'PI',
    category: C,
    syntax: 'PI()',
    description: 'Returns the value of pi.',
    minArgs: 0,
    maxArgs: 0,
    impl: () => Math.PI,
  },
  {
    name: 'RAND',
    category: C,
    syntax: 'RAND()',
    description: 'Returns a random number between 0 and 1.',
    minArgs: 0,
    maxArgs: 0,
    volatile: true,
    impl: (_args, ctx) => ctx.host.random(),
  },
  {
    name: 'RANDBETWEEN',
    category: C,
    syntax: 'RANDBETWEEN(bottom, top)',
    description: 'Returns a random integer between the numbers you specify.',
    minArgs: 2,
    maxArgs: 2,
    volatile: true,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const lo = Math.ceil(v[0] ?? 0);
      const hi = Math.floor(v[1] ?? 0);
      if (lo > hi) return ERRORS.NUM;
      return lo + Math.floor(ctx.host.random() * (hi - lo + 1));
    },
  },
  {
    name: 'RANDARRAY',
    category: C,
    syntax: 'RANDARRAY([rows], [columns], [min], [max], [whole_number])',
    description: 'Returns an array of random numbers.',
    minArgs: 0,
    maxArgs: 5,
    volatile: true,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [1, 1, 0, 1]);
      if (isError(v)) return v;
      const whole = optBool(args[4], false);
      if (isError(whole)) return whole;
      const [rows = 1, cols = 1, lo = 0, hi = 1] = v;
      const r = trunc(rows);
      const c = trunc(cols);
      if (r < 1 || c < 1 || lo > hi) return r === 0 || c === 0 ? ERRORS.CALC : ERRORS.VALUE;
      return arrayOf(r, c, () => {
        const x = ctx.host.random();
        return whole ? Math.ceil(lo) + Math.floor(x * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + x * (hi - lo);
      });
    },
  },
  {
    name: 'SEQUENCE',
    category: C,
    syntax: 'SEQUENCE(rows, [columns], [start], [step])',
    description: 'Generates a list of sequential numbers in an array.',
    minArgs: 1,
    maxArgs: 4,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, 1, 1, 1]);
      if (isError(v)) return v;
      const [rows = 1, cols = 1, start = 1, step = 1] = v;
      const r = trunc(rows);
      const c = trunc(cols);
      if (r === 0 || c === 0) return ERRORS.CALC;
      if (r < 0 || c < 0) return ERRORS.VALUE;
      return arrayOf(r, c, (i, j) => start + (i * c + j) * step);
    },
  },
  {
    name: 'GCD',
    category: C,
    syntax: 'GCD(number1, [number2], ...)',
    description: 'Returns the greatest common divisor.',
    minArgs: 1,
    maxArgs: 255,
    args: ['ref'],
    impl: (args, ctx) => {
      const v = integerArgs(args, ctx);
      return isError(v) ? v : v.reduce(gcd2, 0);
    },
  },
  {
    name: 'LCM',
    category: C,
    syntax: 'LCM(number1, [number2], ...)',
    description: 'Returns the least common multiple.',
    minArgs: 1,
    maxArgs: 255,
    args: ['ref'],
    impl: (args, ctx) => {
      const v = integerArgs(args, ctx);
      if (isError(v)) return v;
      if (v.some((x) => x === 0)) return 0;
      return checked(v.reduce((a, b) => (a * b) / gcd2(a, b), 1));
    },
  },
  unary('FACT', 'Returns the factorial of a number.', (x) => (x < 0 ? ERRORS.NUM : factorial(Math.floor(x)))),
  unary('FACTDOUBLE', 'Returns the double factorial of a number.', (x) => {
    if (x < -1) return ERRORS.NUM;
    let r = 1;
    for (let i = Math.floor(x); i > 1; i -= 2) r *= i;
    return r;
  }),
  {
    name: 'COMBIN',
    category: C,
    syntax: 'COMBIN(number, number_chosen)',
    description: 'Returns the number of combinations for a given number of objects.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const n = trunc(v[0] ?? 0);
      const k = trunc(v[1] ?? 0);
      if (n < 0 || k < 0 || k > n) return ERRORS.NUM;
      return checked(combin(n, k));
    },
  },
  {
    name: 'COMBINA',
    category: C,
    syntax: 'COMBINA(number, number_chosen)',
    description: 'Returns the number of combinations with repetitions.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const n = trunc(v[0] ?? 0);
      const k = trunc(v[1] ?? 0);
      if (n < 0 || k < 0 || (n === 0 && k > 0)) return ERRORS.NUM;
      return checked(combin(n + k - 1, k));
    },
  },
  {
    name: 'MULTINOMIAL',
    category: C,
    syntax: 'MULTINOMIAL(number1, [number2], ...)',
    description: 'Returns the multinomial of a set of numbers.',
    minArgs: 1,
    maxArgs: 255,
    args: ['ref'],
    impl: (args, ctx) => {
      const v = integerArgs(args, ctx);
      if (isError(v)) return v;
      return checked(factorial(sum(v)) / v.reduce((a, b) => a * factorial(b), 1));
    },
  },
  unary('EVEN', 'Rounds a number up to the nearest even integer.', (x) => {
    const r = Math.ceil(Math.abs(x) / 2) * 2;
    return x < 0 ? -r : r;
  }),
  unary('ODD', 'Rounds a number up to the nearest odd integer.', (x) => {
    let r = Math.ceil(Math.abs(x));
    if (r % 2 === 0) r += 1;
    return x < 0 ? -r : r;
  }),
  unary('DEGREES', 'Converts radians to degrees.', (x) => (x * 180) / Math.PI, 'DEGREES(angle)'),
  unary('RADIANS', 'Converts degrees to radians.', (x) => (x * Math.PI) / 180, 'RADIANS(angle)'),
  unary('SIN', 'Returns the sine of the given angle.', Math.sin),
  unary('COS', 'Returns the cosine of a number.', Math.cos),
  unary('TAN', 'Returns the tangent of a number.', Math.tan),
  unary('ASIN', 'Returns the arcsine of a number.', (x) => (Math.abs(x) > 1 ? ERRORS.NUM : Math.asin(x))),
  unary('ACOS', 'Returns the arccosine of a number.', (x) => (Math.abs(x) > 1 ? ERRORS.NUM : Math.acos(x))),
  unary('ATAN', 'Returns the arctangent of a number.', Math.atan),
  unary('SINH', 'Returns the hyperbolic sine of a number.', Math.sinh),
  unary('COSH', 'Returns the hyperbolic cosine of a number.', Math.cosh),
  unary('TANH', 'Returns the hyperbolic tangent of a number.', Math.tanh),
  unary('ASINH', 'Returns the inverse hyperbolic sine of a number.', Math.asinh),
  unary('ACOSH', 'Returns the inverse hyperbolic cosine of a number.', (x) => (x < 1 ? ERRORS.NUM : Math.acosh(x))),
  unary('ATANH', 'Returns the inverse hyperbolic tangent of a number.', (x) => (Math.abs(x) >= 1 ? ERRORS.NUM : Math.atanh(x))),
  unary('COT', 'Returns the cotangent of an angle.', (x) => (x === 0 ? ERRORS.DIV0 : 1 / Math.tan(x))),
  unary('CSC', 'Returns the cosecant of an angle.', (x) => (x === 0 ? ERRORS.DIV0 : 1 / Math.sin(x))),
  unary('SEC', 'Returns the secant of an angle.', (x) => 1 / Math.cos(x)),
  unary('COTH', 'Returns the hyperbolic cotangent of a number.', (x) => (x === 0 ? ERRORS.DIV0 : 1 / Math.tanh(x))),
  unary('CSCH', 'Returns the hyperbolic cosecant of an angle.', (x) => (x === 0 ? ERRORS.DIV0 : 1 / Math.sinh(x))),
  unary('SECH', 'Returns the hyperbolic secant of an angle.', (x) => 1 / Math.cosh(x)),
  unary('ACOT', 'Returns the arccotangent of a number.', (x) => Math.PI / 2 - Math.atan(x)),
  {
    name: 'ATAN2',
    category: C,
    syntax: 'ATAN2(x_num, y_num)',
    description: 'Returns the arctangent from x- and y-coordinates.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const [x = 0, y = 0] = v;
      return x === 0 && y === 0 ? ERRORS.DIV0 : Math.atan2(y, x);
    },
  },
  {
    name: 'SUBTOTAL',
    category: C,
    syntax: 'SUBTOTAL(function_num, ref1, [ref2], ...)',
    description: 'Returns a subtotal in a list, ignoring other subtotals (and hidden rows for codes 101-111).',
    minArgs: 2,
    maxArgs: 255,
    args: ['scalar', 'ref'],
    impl: (args, ctx) => {
      const f = num(args[0], ctx);
      if (isError(f)) return f;
      const code = trunc(f);
      const base = code > 100 ? code - 100 : code;
      if (base < 1 || base > 11) return ERRORS.VALUE;
      const values = subtotalValues(args.slice(1), ctx, { hidden: code > 100, nested: true, errors: false });
      return isError(values) ? values : aggregateBy(base, values);
    },
  },
  {
    name: 'AGGREGATE',
    category: C,
    syntax: 'AGGREGATE(function_num, options, ref1, [ref2], ...)',
    description: 'Returns an aggregate in a list, optionally ignoring hidden rows and error values.',
    minArgs: 3,
    maxArgs: 255,
    args: ['scalar', 'scalar', 'ref', 'scalar'],
    rest: ['ref'],
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const f = trunc(v[0] ?? 0);
      const o = trunc(v[1] ?? 0);
      if (f < 1 || f > 19 || o < 0 || o > 7) return ERRORS.VALUE;
      const opts = { nested: o <= 3, hidden: o % 2 === 1, errors: o === 2 || o === 3 || o === 6 || o === 7 };
      if (f <= 13) {
        const values = subtotalValues(args.slice(2), ctx, opts);
        return isError(values) ? values : aggregateBy(f, values);
      }
      const values = subtotalValues([args[2] ?? null], ctx, opts);
      if (isError(values)) return values;
      const err = values.find(isError);
      if (err !== undefined) return err;
      const numbers = values.filter((x): x is number => typeof x === 'number');
      const k = num(args[3], ctx);
      if (isError(k)) return k;
      switch (f) {
        case 14:
          return kth(numbers, k, true);
        case 15:
          return kth(numbers, k, false);
        case 16:
          return percentileInc(numbers, k);
        case 17:
          return percentileInc(numbers, trunc(k) / 4);
        case 18:
          return percentileExc(numbers, k);
        default:
          return percentileExc(numbers, trunc(k) / 4);
      }
    },
  },
  {
    name: 'MMULT',
    category: C,
    syntax: 'MMULT(array1, array2)',
    description: 'Returns the matrix product of two arrays.',
    minArgs: 2,
    maxArgs: 2,
    args: ['value', 'value'],
    impl: (args, ctx) => {
      const a = matrix(args[0] ?? null, ctx);
      if (isError(a)) return a;
      const b = matrix(args[1] ?? null, ctx);
      if (isError(b)) return b;
      if (a.cols !== b.rows) return ERRORS.VALUE;
      return arrayOf(a.rows, b.cols, (r, c) => {
        let s = 0;
        for (let k = 0; k < a.cols; k++) s += (a.m[r * a.cols + k] ?? 0) * (b.m[k * b.cols + c] ?? 0);
        return s;
      });
    },
  },
  {
    name: 'MDETERM',
    category: C,
    syntax: 'MDETERM(array)',
    description: 'Returns the matrix determinant of an array.',
    minArgs: 1,
    maxArgs: 1,
    args: ['value'],
    impl: (args, ctx) => {
      const a = matrix(args[0] ?? null, ctx);
      if (isError(a)) return a;
      if (a.rows !== a.cols) return ERRORS.VALUE;
      return eliminate(a.rows, a.m, false).det;
    },
  },
  {
    name: 'MINVERSE',
    category: C,
    syntax: 'MINVERSE(array)',
    description: 'Returns the inverse matrix of an array.',
    minArgs: 1,
    maxArgs: 1,
    args: ['value'],
    impl: (args, ctx) => {
      const a = matrix(args[0] ?? null, ctx);
      if (isError(a)) return a;
      if (a.rows !== a.cols) return ERRORS.VALUE;
      const { det, inv } = eliminate(a.rows, a.m, true);
      if (det === 0) return ERRORS.NUM;
      return makeArray(a.rows, a.cols, inv);
    },
  },
  {
    name: 'MUNIT',
    category: C,
    syntax: 'MUNIT(dimension)',
    description: 'Returns the unit matrix for the specified dimension.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const n = num(args[0], ctx);
      if (isError(n)) return n;
      const d = trunc(n);
      if (d < 1) return ERRORS.VALUE;
      return arrayOf(d, d, (r, c) => (r === c ? 1 : 0));
    },
  },
  {
    name: 'ROMAN',
    category: C,
    syntax: 'ROMAN(number, [form])',
    description: 'Converts an arabic numeral to roman, as text.',
    minArgs: 1,
    maxArgs: 2,
    impl: (args, ctx) => {
      const n = num(args[0], ctx);
      if (isError(n)) return n;
      let v = trunc(n);
      if (v < 0 || v > 3999) return ERRORS.VALUE;
      let out = '';
      for (const [value, sym] of ROMAN) {
        while (v >= value) {
          out += sym;
          v -= value;
        }
      }
      return out;
    },
  },
  {
    name: 'ARABIC',
    category: C,
    syntax: 'ARABIC(text)',
    description: 'Converts a Roman numeral to an Arabic numeral.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args) => {
      const t = str(args[0]);
      if (isError(t)) return t;
      const s = t.trim().toUpperCase();
      const negative = s.startsWith('-');
      const body = negative ? s.slice(1) : s;
      const values: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
      let total = 0;
      for (let i = 0; i < body.length; i++) {
        const cur = values[body[i] ?? ''];
        const next = values[body[i + 1] ?? ''] ?? 0;
        if (cur === undefined) return ERRORS.VALUE;
        total += cur < next ? -cur : cur;
      }
      return negative ? -total : total;
    },
  },
  {
    name: 'BASE',
    category: C,
    syntax: 'BASE(number, radix, [min_length])',
    description: 'Converts a number into a text representation with the given radix (base).',
    minArgs: 2,
    maxArgs: 3,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined, 0]);
      if (isError(v)) return v;
      const n = trunc(v[0] ?? 0);
      const radix = trunc(v[1] ?? 0);
      const len = trunc(v[2] ?? 0);
      if (n < 0 || n >= 2 ** 53 || radix < 2 || radix > 36 || len < 0 || len > 255) return ERRORS.NUM;
      return n.toString(radix).toUpperCase().padStart(len, '0');
    },
  },
  {
    name: 'DECIMAL',
    category: C,
    syntax: 'DECIMAL(text, radix)',
    description: 'Converts a text representation of a number in a given base into a decimal number.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const t = str(args[0]);
      if (isError(t)) return t;
      const r = num(args[1], ctx);
      if (isError(r)) return r;
      const radix = trunc(r);
      if (radix < 2 || radix > 36) return ERRORS.NUM;
      let total = 0;
      for (const ch of t.trim().toUpperCase()) {
        const d = Number.parseInt(ch, 36);
        if (Number.isNaN(d) || d >= radix) return ERRORS.NUM;
        total = total * radix + d;
      }
      return total;
    },
  },
];
