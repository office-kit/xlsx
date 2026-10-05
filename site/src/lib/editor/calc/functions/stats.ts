// Statistical functions.

import { toNumber } from '../coerce.ts';
import type { FnContext, FunctionSpec } from '../function-spec.ts';
import { type CalcError, type CalcScalar, type CalcValue, ERRORS, isError, isRef } from '../types.ts';
import {
  arrayOf,
  checked,
  collectNumbers,
  column,
  flatten,
  forEachValue,
  num,
  nums,
  optBool,
  optNum,
  sum,
  trunc,
  variance,
} from './helpers.ts';

const C = 'Statistical' as const;

// ---- numeric cores (also used by SUBTOTAL / AGGREGATE) ---------------------

export function median(values: readonly number[]): number | CalcError {
  if (values.length === 0) return ERRORS.NUM;
  const s = [...values].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 === 1 ? (s[mid] ?? 0) : ((s[mid - 1] ?? 0) + (s[mid] ?? 0)) / 2;
}

export function modeSingle(values: readonly number[]): number | CalcError {
  const counts = new Map<number, number>();
  let best: number | undefined;
  let bestCount = 1;
  for (const v of values) {
    const c = (counts.get(v) ?? 0) + 1;
    counts.set(v, c);
    if (c > bestCount) {
      bestCount = c;
      best = v;
    }
  }
  // Ties go to the value that appeared first, which a later equal count must not displace.
  if (best === undefined) return ERRORS.NA;
  for (const v of values) if (counts.get(v) === bestCount) return v;
  return best;
}

export function percentileInc(values: readonly number[], k: number): number | CalcError {
  if (values.length === 0 || k < 0 || k > 1) return ERRORS.NUM;
  const s = [...values].sort((a, b) => a - b);
  const h = (s.length - 1) * k;
  const lo = Math.floor(h);
  const a = s[lo] ?? 0;
  const b = s[Math.min(lo + 1, s.length - 1)] ?? 0;
  return a + (h - lo) * (b - a);
}

export function percentileExc(values: readonly number[], k: number): number | CalcError {
  const n = values.length;
  if (n === 0 || k <= 0 || k >= 1) return ERRORS.NUM;
  const h = (n + 1) * k - 1;
  if (h < 0 || h > n - 1) return ERRORS.NUM;
  const s = [...values].sort((a, b) => a - b);
  const lo = Math.floor(h);
  const a = s[lo] ?? 0;
  const b = s[Math.min(lo + 1, n - 1)] ?? 0;
  return a + (h - lo) * (b - a);
}

export function kth(values: readonly number[], k: number, largest: boolean): number | CalcError {
  const idx = Math.ceil(k);
  if (idx < 1 || idx > values.length) return ERRORS.NUM;
  const s = [...values].sort((a, b) => (largest ? b - a : a - b));
  return s[idx - 1] ?? ERRORS.NUM;
}

export function stdev(values: readonly number[], sample: boolean): number | CalcError {
  const v = variance(values, sample);
  return isError(v) ? v : Math.sqrt(v);
}

// ---- helpers ---------------------------------------------------------------

const numbersOf = (args: readonly CalcValue[], ctx: FnContext, mode: 'plain' | 'a' = 'plain') =>
  collectNumbers(args, ctx, mode);

const aggregate =
  (fn: (values: number[]) => CalcValue, mode: 'plain' | 'a' = 'plain') =>
  (args: CalcValue[], ctx: FnContext): CalcValue => {
    const values = numbersOf(args, ctx, mode);
    return isError(values) ? values : fn(values);
  };

/** Paired numeric data for CORREL / SLOPE …: positions where both sides hold numbers. */
const pairs = (ys: CalcValue, xs: CalcValue, ctx: FnContext): { x: number[]; y: number[] } | CalcError => {
  const a = flatten(ys, ctx);
  const b = flatten(xs, ctx);
  if (a.length !== b.length) return ERRORS.NA;
  const x: number[] = [];
  const y: number[] = [];
  for (let i = 0; i < a.length; i++) {
    const p = a[i] ?? null;
    const q = b[i] ?? null;
    if (isError(p)) return p;
    if (isError(q)) return q;
    if (typeof p === 'number' && typeof q === 'number') {
      y.push(p);
      x.push(q);
    }
  }
  return { x, y };
};

const regression = (ys: CalcValue, xs: CalcValue, ctx: FnContext) => {
  const p = pairs(ys, xs, ctx);
  if (isError(p)) return p;
  const n = p.x.length;
  if (n < 1) return ERRORS.DIV0;
  const mx = sum(p.x) / n;
  const my = sum(p.y) / n;
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < n; i++) {
    const dx = (p.x[i] ?? 0) - mx;
    const dy = (p.y[i] ?? 0) - my;
    sxy += dx * dy;
    sxx += dx * dx;
    syy += dy * dy;
  }
  return { n, mx, my, sxy, sxx, syy };
};

const countValues = (args: readonly CalcValue[], ctx: FnContext, accept: (v: CalcScalar, fromRange: boolean) => boolean): number => {
  let n = 0;
  forEachValue(args, ctx, (v, fromRange) => {
    if (accept(v, fromRange)) n++;
  });
  return n;
};

export function normSDist(z: number): number {
  return 0.5 * erfc(-z / Math.SQRT2);
}

// Cody's rational approximations keep ~1e-15 relative accuracy across the
// range, which NORM.S.DIST needs to agree with Excel to 15 digits.
function erfc(x: number): number {
  // W. J. Cody, "Rational Chebyshev approximations for the error function" (1969).
  const ax = Math.abs(x);
  let result: number;
  if (ax < 0.5) {
    const t = x * x;
    const top = (((0.185777706184603153 * t + 3.16112374387056560) * t + 113.864154151050156) * t + 377.485237685302021) * t + 3209.37758913846947;
    const bot = (((t + 23.6012909523441209) * t + 244.024637934444173) * t + 1282.61652607737228) * t + 2844.23683343917062;
    return 1 - (x * top) / bot;
  }
  if (ax < 4) {
    const top =
      (((((((2.15311535474403846e-8 * ax + 0.564188496988670089) * ax + 8.88314979438837594) * ax + 66.1191906371416295) * ax + 298.635138197400131) * ax + 881.95222124176909) * ax + 1712.04761263407058) * ax + 2051.07837782607147) * ax + 1230.33935479799725;
    const bot =
      (((((((ax + 15.7449261107098347) * ax + 117.693950891312499) * ax + 537.181101862009858) * ax + 1621.38957456669019) * ax + 3290.79923573345963) * ax + 4362.61909014324716) * ax + 3439.36767414372164) * ax + 1230.33935480374942;
    result = Math.exp(-ax * ax) * (top / bot);
  } else {
    const z = 1 / (ax * ax);
    const top = ((((0.0163153871373020978 * z + 0.305326634961232344) * z + 0.360344899949804439) * z + 0.125781726111229246) * z + 0.0160837851487422766) * z + 6.58749161529837803e-4;
    const bot = ((((z + 2.56852019228982242) * z + 1.87295284992346725) * z + 0.527905102951428412) * z + 0.0605183413124413191) * z + 0.00233520497626869185;
    result = (Math.exp(-ax * ax) / ax) * (0.564189583547756287 - (z * top) / bot);
  }
  return x < 0 ? 2 - result : result;
}

/** Inverse standard normal CDF (Acklam's algorithm with one Halley refinement). */
export function normSInv(p: number): number {
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const at = (arr: number[], i: number): number => arr[i] ?? 0;
  const pLow = 0.02425;
  let x: number;
  if (p < pLow) {
    const q = Math.sqrt(-2 * Math.log(p));
    x = (((((at(c, 0) * q + at(c, 1)) * q + at(c, 2)) * q + at(c, 3)) * q + at(c, 4)) * q + at(c, 5)) / ((((at(d, 0) * q + at(d, 1)) * q + at(d, 2)) * q + at(d, 3)) * q + 1);
  } else if (p <= 1 - pLow) {
    const q = p - 0.5;
    const r = q * q;
    x = ((((((at(a, 0) * r + at(a, 1)) * r + at(a, 2)) * r + at(a, 3)) * r + at(a, 4)) * r + at(a, 5)) * q) / (((((at(b, 0) * r + at(b, 1)) * r + at(b, 2)) * r + at(b, 3)) * r + at(b, 4)) * r + 1);
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((at(c, 0) * q + at(c, 1)) * q + at(c, 2)) * q + at(c, 3)) * q + at(c, 4)) * q + at(c, 5)) / ((((at(d, 0) * q + at(d, 1)) * q + at(d, 2)) * q + at(d, 3)) * q + 1);
  }
  const e = normSDist(x) - p;
  const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2);
  return x - u / (1 + (x * u) / 2);
}

export function gammaLn(x: number): number {
  // Lanczos approximation (g = 7, n = 9).
  const g = [
    0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905,
    -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
  ];
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - gammaLn(1 - x);
  const xx = x - 1;
  let a = g[0] ?? 0;
  const t = xx + 7.5;
  for (let i = 1; i < 9; i++) a += (g[i] ?? 0) / (xx + i);
  return 0.5 * Math.log(2 * Math.PI) + (xx + 0.5) * Math.log(t) - t + Math.log(a);
}

export function gamma(x: number): number {
  if (Number.isInteger(x) && x > 0 && x < 171) {
    let r = 1;
    for (let i = 2; i < x; i++) r *= i;
    return r;
  }
  const sign = x < 0 && Math.floor(-x) % 2 === 0 ? -1 : 1;
  return sign * Math.exp(gammaLn(x));
}

const logFactorial = (n: number): number => gammaLn(n + 1);

export function combin(n: number, k: number): number {
  if (k < 0 || k > n) return 0;
  let r = 1;
  const kk = Math.min(k, n - k);
  for (let i = 1; i <= kk; i++) r = (r * (n - kk + i)) / i;
  return Math.round(r);
}

// ---- conditional / ranking helpers -----------------------------------------

const rank = (args: CalcValue[], ctx: FnContext, average: boolean): CalcValue => {
  const x = num(args[0], ctx);
  if (isError(x)) return x;
  const ref = args[1];
  if (ref === undefined || !isRef(ref)) return ERRORS.NA;
  const order = optNum(args[2], 0, ctx);
  if (isError(order)) return order;
  const values = collectNumbers([ref], ctx);
  if (isError(values)) return values;
  let better = 0;
  let ties = 0;
  for (const v of values) {
    if (v === x) ties++;
    else if (order === 0 ? v > x : v < x) better++;
  }
  if (ties === 0) return ERRORS.NA;
  return average ? better + (ties + 1) / 2 : better + 1;
};

const percentRank = (args: CalcValue[], ctx: FnContext, exclusive: boolean): CalcValue => {
  const values = collectNumbers([args[0] ?? null], ctx);
  if (isError(values)) return values;
  const x = num(args[1], ctx);
  if (isError(x)) return x;
  const sig = optNum(args[2], 3, ctx);
  if (isError(sig)) return sig;
  if (sig < 1 || values.length === 0) return ERRORS.NUM;
  const s = [...values].sort((a, b) => a - b);
  const n = s.length;
  const first = s[0] ?? 0;
  const last = s[n - 1] ?? 0;
  if (x < first || x > last) return ERRORS.NA;
  let pos: number | undefined;
  for (let i = 0; i < n; i++) {
    const v = s[i] ?? 0;
    if (v === x) {
      pos = i;
      break;
    }
    const next = s[i + 1];
    if (next !== undefined && v < x && x < next) {
      pos = i + (x - v) / (next - v);
      break;
    }
  }
  if (pos === undefined) return ERRORS.NA;
  const raw = exclusive ? (pos + 1) / (n + 1) : n === 1 ? 1 : pos / (n - 1);
  const factor = 10 ** Math.floor(sig);
  return Math.floor(raw * factor + 1e-9) / factor;
};

// ---- specs -----------------------------------------------------------------

const many = (name: string, syntax: string, description: string, impl: (args: CalcValue[], ctx: FnContext) => CalcValue): FunctionSpec => ({
  name,
  category: C,
  syntax,
  description,
  minArgs: 1,
  maxArgs: 255,
  args: ['ref'],
  impl,
});

const varianceOf = (sample: boolean, mode: 'plain' | 'a') => aggregate((v) => variance(v, sample), mode);
const stdevOf = (sample: boolean, mode: 'plain' | 'a') => aggregate((v) => stdev(v, sample), mode);

export const STATISTICAL_FUNCTIONS: FunctionSpec[] = [
  many('AVERAGE', 'AVERAGE(number1, [number2], ...)', 'Returns the arithmetic mean of its arguments.', aggregate((v) => (v.length === 0 ? ERRORS.DIV0 : sum(v) / v.length))),
  many('AVERAGEA', 'AVERAGEA(value1, [value2], ...)', 'Returns the mean of its arguments, counting text as 0 and TRUE as 1.', aggregate((v) => (v.length === 0 ? ERRORS.DIV0 : sum(v) / v.length), 'a')),
  many('COUNT', 'COUNT(value1, [value2], ...)', 'Counts how many numbers are in the list of arguments.', (args, ctx) =>
    countValues(args, ctx, (v, fromRange) => typeof v === 'number' || (!fromRange && (typeof v === 'boolean' || (typeof v === 'string' && !isError(toNumber(v, ctx.host.date1904)))))),
  ),
  many('COUNTA', 'COUNTA(value1, [value2], ...)', 'Counts how many values are not empty.', (args, ctx) => countValues(args, ctx, (v, fromRange) => v !== null || !fromRange)),
  {
    name: 'COUNTBLANK',
    category: C,
    syntax: 'COUNTBLANK(range)',
    description: 'Counts the empty cells in a range.',
    minArgs: 1,
    maxArgs: 1,
    args: ['ref'],
    impl: (args, ctx) => {
      const ref = args[0];
      if (ref === undefined || !isRef(ref)) return ERRORS.VALUE;
      let total = 0;
      for (const area of ref.areas) {
        const values = ctx.host.areaValues(area);
        const size = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1);
        let filled = 0;
        for (const v of values.data) if (v !== null && v !== '') filled++;
        total += size - filled;
      }
      return total;
    },
  },
  many('MAX', 'MAX(number1, [number2], ...)', 'Returns the largest value in a set of values.', aggregate((v) => (v.length === 0 ? 0 : v.reduce((a, b) => Math.max(a, b))))),
  many('MIN', 'MIN(number1, [number2], ...)', 'Returns the smallest value in a set of values.', aggregate((v) => (v.length === 0 ? 0 : v.reduce((a, b) => Math.min(a, b))))),
  many('MAXA', 'MAXA(value1, [value2], ...)', 'Returns the largest value, counting text as 0 and TRUE as 1.', aggregate((v) => (v.length === 0 ? 0 : v.reduce((a, b) => Math.max(a, b))), 'a')),
  many('MINA', 'MINA(value1, [value2], ...)', 'Returns the smallest value, counting text as 0 and TRUE as 1.', aggregate((v) => (v.length === 0 ? 0 : v.reduce((a, b) => Math.min(a, b))), 'a')),
  many('MEDIAN', 'MEDIAN(number1, [number2], ...)', 'Returns the median of the given numbers.', aggregate(median)),
  many('MODE', 'MODE(number1, [number2], ...)', 'Returns the most frequently occurring value.', aggregate(modeSingle)),
  many('MODE.SNGL', 'MODE.SNGL(number1, [number2], ...)', 'Returns the most frequently occurring value.', aggregate(modeSingle)),
  many('MODE.MULT', 'MODE.MULT(number1, [number2], ...)', 'Returns a vertical array of the most frequently occurring values.', aggregate((v) => {
    const counts = new Map<number, number>();
    for (const x of v) counts.set(x, (counts.get(x) ?? 0) + 1);
    const best = Math.max(0, ...counts.values());
    if (best < 2) return ERRORS.NA;
    return column([...counts].filter(([, c]) => c === best).map(([x]) => x));
  })),
  many('STDEV', 'STDEV(number1, [number2], ...)', 'Estimates standard deviation based on a sample.', stdevOf(true, 'plain')),
  many('STDEV.S', 'STDEV.S(number1, [number2], ...)', 'Estimates standard deviation based on a sample.', stdevOf(true, 'plain')),
  many('STDEV.P', 'STDEV.P(number1, [number2], ...)', 'Calculates standard deviation based on the entire population.', stdevOf(false, 'plain')),
  many('STDEVP', 'STDEVP(number1, [number2], ...)', 'Calculates standard deviation based on the entire population.', stdevOf(false, 'plain')),
  many('STDEVA', 'STDEVA(value1, [value2], ...)', 'Estimates standard deviation based on a sample, including text and logical values.', stdevOf(true, 'a')),
  many('STDEVPA', 'STDEVPA(value1, [value2], ...)', 'Calculates population standard deviation, including text and logical values.', stdevOf(false, 'a')),
  many('VAR', 'VAR(number1, [number2], ...)', 'Estimates variance based on a sample.', varianceOf(true, 'plain')),
  many('VAR.S', 'VAR.S(number1, [number2], ...)', 'Estimates variance based on a sample.', varianceOf(true, 'plain')),
  many('VAR.P', 'VAR.P(number1, [number2], ...)', 'Calculates variance based on the entire population.', varianceOf(false, 'plain')),
  many('VARP', 'VARP(number1, [number2], ...)', 'Calculates variance based on the entire population.', varianceOf(false, 'plain')),
  many('VARA', 'VARA(value1, [value2], ...)', 'Estimates variance based on a sample, including text and logical values.', varianceOf(true, 'a')),
  many('VARPA', 'VARPA(value1, [value2], ...)', 'Calculates population variance, including text and logical values.', varianceOf(false, 'a')),
  many('GEOMEAN', 'GEOMEAN(number1, [number2], ...)', 'Returns the geometric mean.', aggregate((v) => {
    if (v.length === 0 || v.some((x) => x <= 0)) return ERRORS.NUM;
    return Math.exp(sum(v.map(Math.log)) / v.length);
  })),
  many('HARMEAN', 'HARMEAN(number1, [number2], ...)', 'Returns the harmonic mean.', aggregate((v) => {
    if (v.length === 0 || v.some((x) => x <= 0)) return ERRORS.NUM;
    return v.length / sum(v.map((x) => 1 / x));
  })),
  many('AVEDEV', 'AVEDEV(number1, [number2], ...)', 'Returns the average of the absolute deviations from the mean.', aggregate((v) => {
    if (v.length === 0) return ERRORS.NUM;
    const m = sum(v) / v.length;
    return sum(v.map((x) => Math.abs(x - m))) / v.length;
  })),
  many('DEVSQ', 'DEVSQ(number1, [number2], ...)', 'Returns the sum of squares of deviations from the mean.', aggregate((v) => {
    if (v.length === 0) return ERRORS.NUM;
    const m = sum(v) / v.length;
    return sum(v.map((x) => (x - m) ** 2));
  })),
  many('SKEW', 'SKEW(number1, [number2], ...)', 'Returns the skewness of a distribution.', aggregate((v) => {
    const n = v.length;
    const sd = stdev(v, true);
    if (n < 3 || isError(sd) || sd === 0) return ERRORS.DIV0;
    const m = sum(v) / n;
    return (n / ((n - 1) * (n - 2))) * sum(v.map((x) => ((x - m) / sd) ** 3));
  })),
  many('KURT', 'KURT(number1, [number2], ...)', 'Returns the kurtosis of a data set.', aggregate((v) => {
    const n = v.length;
    const sd = stdev(v, true);
    if (n < 4 || isError(sd) || sd === 0) return ERRORS.DIV0;
    const m = sum(v) / n;
    const s4 = sum(v.map((x) => ((x - m) / sd) ** 4));
    return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * s4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
  })),
  {
    name: 'LARGE',
    category: C,
    syntax: 'LARGE(array, k)',
    description: 'Returns the k-th largest value in a data set.',
    minArgs: 2,
    maxArgs: 2,
    args: ['ref', 'scalar'],
    impl: (args, ctx) => {
      const values = collectNumbers([args[0] ?? null], ctx);
      if (isError(values)) return values;
      const k = num(args[1], ctx);
      return isError(k) ? k : kth(values, k, true);
    },
  },
  {
    name: 'SMALL',
    category: C,
    syntax: 'SMALL(array, k)',
    description: 'Returns the k-th smallest value in a data set.',
    minArgs: 2,
    maxArgs: 2,
    args: ['ref', 'scalar'],
    impl: (args, ctx) => {
      const values = collectNumbers([args[0] ?? null], ctx);
      if (isError(values)) return values;
      const k = num(args[1], ctx);
      return isError(k) ? k : kth(values, k, false);
    },
  },
  ...(['RANK', 'RANK.EQ'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(number, ref, [order])`,
      description: 'Returns the rank of a number in a list of numbers.',
      minArgs: 2,
      maxArgs: 3,
      args: ['scalar', 'ref', 'scalar'],
      impl: (args, ctx) => rank(args, ctx, false),
    }),
  ),
  {
    name: 'RANK.AVG',
    category: C,
    syntax: 'RANK.AVG(number, ref, [order])',
    description: 'Returns the rank of a number, averaging the rank of ties.',
    minArgs: 2,
    maxArgs: 3,
    args: ['scalar', 'ref', 'scalar'],
    impl: (args, ctx) => rank(args, ctx, true),
  },
  ...(
    [
      ['PERCENTILE', percentileInc, 'Returns the k-th percentile of values in a range.'],
      ['PERCENTILE.INC', percentileInc, 'Returns the k-th percentile of values in a range, k inclusive of 0 and 1.'],
      ['PERCENTILE.EXC', percentileExc, 'Returns the k-th percentile of values in a range, k exclusive of 0 and 1.'],
    ] as const
  ).map(
    ([name, fn, description]): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array, k)`,
      description,
      minArgs: 2,
      maxArgs: 2,
      args: ['ref', 'scalar'],
      impl: (args, ctx) => {
        const values = collectNumbers([args[0] ?? null], ctx);
        if (isError(values)) return values;
        const k = num(args[1], ctx);
        return isError(k) ? k : fn(values, k);
      },
    }),
  ),
  ...(
    [
      ['QUARTILE', false],
      ['QUARTILE.INC', false],
      ['QUARTILE.EXC', true],
    ] as const
  ).map(
    ([name, exclusive]): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array, quart)`,
      description: 'Returns the quartile of a data set.',
      minArgs: 2,
      maxArgs: 2,
      args: ['ref', 'scalar'],
      impl: (args, ctx) => {
        const values = collectNumbers([args[0] ?? null], ctx);
        if (isError(values)) return values;
        const q = num(args[1], ctx);
        if (isError(q)) return q;
        const quart = trunc(q);
        if (exclusive ? quart < 1 || quart > 3 : quart < 0 || quart > 4) return ERRORS.NUM;
        return exclusive ? percentileExc(values, quart / 4) : percentileInc(values, quart / 4);
      },
    }),
  ),
  ...(
    [
      ['PERCENTRANK', false],
      ['PERCENTRANK.INC', false],
      ['PERCENTRANK.EXC', true],
    ] as const
  ).map(
    ([name, exclusive]): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array, x, [significance])`,
      description: 'Returns the rank of a value in a data set as a percentage.',
      minArgs: 2,
      maxArgs: 3,
      args: ['ref', 'scalar', 'scalar'],
      impl: (args, ctx) => percentRank(args, ctx, exclusive),
    }),
  ),
  {
    name: 'TRIMMEAN',
    category: C,
    syntax: 'TRIMMEAN(array, percent)',
    description: 'Returns the mean of the interior of a data set.',
    minArgs: 2,
    maxArgs: 2,
    args: ['ref', 'scalar'],
    impl: (args, ctx) => {
      const v = collectNumbers([args[0] ?? null], ctx);
      if (isError(v)) return v;
      const pct = num(args[1], ctx);
      if (isError(pct)) return pct;
      if (pct < 0 || pct >= 1 || v.length === 0) return ERRORS.NUM;
      const cut = Math.floor((v.length * pct) / 2);
      const s = [...v].sort((a, b) => a - b).slice(cut, v.length - cut);
      return sum(s) / s.length;
    },
  },
  ...(['CORREL', 'PEARSON'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array1, array2)`,
      description: 'Returns the correlation coefficient between two data sets.',
      minArgs: 2,
      maxArgs: 2,
      args: ['value', 'value'],
      impl: (args, ctx) => {
        const r = regression(args[0] ?? null, args[1] ?? null, ctx);
        if (isError(r)) return r;
        if (r.sxx === 0 || r.syy === 0) return ERRORS.DIV0;
        return r.sxy / Math.sqrt(r.sxx * r.syy);
      },
    }),
  ),
  {
    name: 'RSQ',
    category: C,
    syntax: "RSQ(known_y's, known_x's)",
    description: 'Returns the square of the Pearson correlation coefficient.',
    minArgs: 2,
    maxArgs: 2,
    args: ['value', 'value'],
    impl: (args, ctx) => {
      const r = regression(args[0] ?? null, args[1] ?? null, ctx);
      if (isError(r)) return r;
      if (r.sxx === 0 || r.syy === 0) return ERRORS.DIV0;
      return (r.sxy * r.sxy) / (r.sxx * r.syy);
    },
  },
  ...(
    [
      ['COVAR', false],
      ['COVARIANCE.P', false],
      ['COVARIANCE.S', true],
    ] as const
  ).map(
    ([name, sample]): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array1, array2)`,
      description: sample ? 'Returns the sample covariance.' : 'Returns the population covariance.',
      minArgs: 2,
      maxArgs: 2,
      args: ['value', 'value'],
      impl: (args, ctx) => {
        const r = regression(args[0] ?? null, args[1] ?? null, ctx);
        if (isError(r)) return r;
        if (r.n < (sample ? 2 : 1)) return ERRORS.DIV0;
        return r.sxy / (sample ? r.n - 1 : r.n);
      },
    }),
  ),
  {
    name: 'SLOPE',
    category: C,
    syntax: "SLOPE(known_y's, known_x's)",
    description: 'Returns the slope of the linear regression line.',
    minArgs: 2,
    maxArgs: 2,
    args: ['value', 'value'],
    impl: (args, ctx) => {
      const r = regression(args[0] ?? null, args[1] ?? null, ctx);
      if (isError(r)) return r;
      return r.sxx === 0 ? ERRORS.DIV0 : r.sxy / r.sxx;
    },
  },
  {
    name: 'INTERCEPT',
    category: C,
    syntax: "INTERCEPT(known_y's, known_x's)",
    description: 'Returns the intercept of the linear regression line.',
    minArgs: 2,
    maxArgs: 2,
    args: ['value', 'value'],
    impl: (args, ctx) => {
      const r = regression(args[0] ?? null, args[1] ?? null, ctx);
      if (isError(r)) return r;
      return r.sxx === 0 ? ERRORS.DIV0 : r.my - (r.sxy / r.sxx) * r.mx;
    },
  },
  {
    name: 'STEYX',
    category: C,
    syntax: "STEYX(known_y's, known_x's)",
    description: 'Returns the standard error of the predicted y-value for each x in the regression.',
    minArgs: 2,
    maxArgs: 2,
    args: ['value', 'value'],
    impl: (args, ctx) => {
      const r = regression(args[0] ?? null, args[1] ?? null, ctx);
      if (isError(r)) return r;
      if (r.n < 3 || r.sxx === 0) return ERRORS.DIV0;
      return Math.sqrt((r.syy - (r.sxy * r.sxy) / r.sxx) / (r.n - 2));
    },
  },
  ...(['FORECAST', 'FORECAST.LINEAR'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(x, known_y's, known_x's)`,
      description: 'Predicts a value along a linear trend.',
      minArgs: 3,
      maxArgs: 3,
      args: ['scalar', 'value', 'value'],
      impl: (args, ctx) => {
        const x = num(args[0], ctx);
        if (isError(x)) return x;
        const r = regression(args[1] ?? null, args[2] ?? null, ctx);
        if (isError(r)) return r;
        if (r.sxx === 0) return ERRORS.DIV0;
        const b = r.sxy / r.sxx;
        return r.my - b * r.mx + b * x;
      },
    }),
  ),
  ...(
    [
      ['TREND', false],
      ['GROWTH', true],
    ] as const
  ).map(
    ([name, exponential]): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(known_y's, [known_x's], [new_x's], [const])`,
      description: exponential ? 'Returns values along an exponential trend.' : 'Returns values along a linear trend.',
      minArgs: 1,
      maxArgs: 4,
      args: ['value', 'value', 'value', 'scalar'],
      impl: (args, ctx) => {
        const ysRaw = ctx.toArray(args[0] ?? null);
        const ys = ysRaw.data.map((v) => (typeof v === 'number' ? v : Number.NaN));
        if (ys.some(Number.isNaN)) return ERRORS.VALUE;
        if (exponential && ys.some((y) => y <= 0)) return ERRORS.NUM;
        const ty = exponential ? ys.map(Math.log) : ys;
        const xsArr = args[1] === undefined || args[1] === null ? undefined : ctx.toArray(args[1]);
        const xs = xsArr === undefined ? ys.map((_, i) => i + 1) : xsArr.data.map((v) => (typeof v === 'number' ? v : Number.NaN));
        if (xs.length !== ys.length || xs.some(Number.isNaN)) return ERRORS.REF;
        const useConst = optBool(args[3], true);
        if (isError(useConst)) return useConst;
        const n = xs.length;
        const mx = useConst ? sum(xs) / n : 0;
        const my = useConst ? sum(ty) / n : 0;
        let sxy = 0;
        let sxx = 0;
        for (let i = 0; i < n; i++) {
          sxy += ((xs[i] ?? 0) - mx) * ((ty[i] ?? 0) - my);
          sxx += ((xs[i] ?? 0) - mx) ** 2;
        }
        if (sxx === 0) return ERRORS.DIV0;
        const b = sxy / sxx;
        const a = my - b * mx;
        const target = args[2] === undefined || args[2] === null ? (xsArr ?? column(xs)) : ctx.toArray(args[2]);
        return arrayOf(target.rows, target.cols, (r, c) => {
          const x = target.data[r * target.cols + c];
          if (typeof x !== 'number') return ERRORS.VALUE;
          const y = a + b * x;
          return exponential ? Math.exp(y) : y;
        });
      },
    }),
  ),
  {
    name: 'LINEST',
    category: C,
    syntax: "LINEST(known_y's, [known_x's], [const], [stats])",
    description: 'Returns the slope and intercept of a simple linear regression.',
    minArgs: 1,
    maxArgs: 4,
    args: ['value', 'value', 'scalar', 'scalar'],
    impl: (args, ctx) => {
      const ys = flatten(args[0] ?? null, ctx);
      const xs = args[1] === undefined || args[1] === null ? ys.map((_, i) => i + 1) : flatten(args[1], ctx);
      if (xs.length !== ys.length) return ERRORS.REF;
      const useConst = optBool(args[2], true);
      if (isError(useConst)) return useConst;
      const n = ys.length;
      const nx = xs.map((v) => (typeof v === 'number' ? v : 0));
      const ny = ys.map((v) => (typeof v === 'number' ? v : 0));
      const mx = useConst ? sum(nx) / n : 0;
      const my = useConst ? sum(ny) / n : 0;
      let sxy = 0;
      let sxx = 0;
      for (let i = 0; i < n; i++) {
        sxy += ((nx[i] ?? 0) - mx) * ((ny[i] ?? 0) - my);
        sxx += ((nx[i] ?? 0) - mx) ** 2;
      }
      if (sxx === 0) return ERRORS.DIV0;
      const b = sxy / sxx;
      return arrayOf(1, 2, (_, c) => (c === 0 ? b : my - b * mx));
    },
  },
  {
    name: 'FREQUENCY',
    category: C,
    syntax: 'FREQUENCY(data_array, bins_array)',
    description: 'Returns how often values occur within ranges of values, as a vertical array.',
    minArgs: 2,
    maxArgs: 2,
    args: ['ref', 'ref'],
    impl: (args, ctx) => {
      const data = collectNumbers([args[0] ?? null], ctx);
      if (isError(data)) return data;
      const bins = collectNumbers([args[1] ?? null], ctx);
      if (isError(bins)) return bins;
      const sortedBins = bins.map((b, i) => ({ b, i })).sort((x, y) => x.b - y.b);
      const counts = new Array<number>(bins.length + 1).fill(0);
      for (const v of data) {
        const hit = sortedBins.find((x) => v <= x.b);
        const slot = hit === undefined ? counts.length - 1 : hit.i;
        counts[slot] = (counts[slot] ?? 0) + 1;
      }
      return column(counts);
    },
  },
  {
    name: 'STANDARDIZE',
    category: C,
    syntax: 'STANDARDIZE(x, mean, standard_dev)',
    description: 'Returns a normalized value.',
    minArgs: 3,
    maxArgs: 3,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined, undefined]);
      if (isError(v)) return v;
      const [x = 0, m = 0, s = 0] = v;
      return s <= 0 ? ERRORS.NUM : (x - m) / s;
    },
  },
  ...(['NORM.DIST', 'NORMDIST'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(x, mean, standard_dev, cumulative)`,
      description: 'Returns the normal distribution.',
      minArgs: 4,
      maxArgs: 4,
      impl: (args, ctx) => {
        const v = nums(args, ctx, [undefined, undefined, undefined]);
        if (isError(v)) return v;
        const [x = 0, m = 0, sd = 0] = v;
        const cum = optBool(args[3], false);
        if (isError(cum)) return cum;
        if (sd <= 0) return ERRORS.NUM;
        const z = (x - m) / sd;
        return cum ? normSDist(z) : Math.exp((-z * z) / 2) / (sd * Math.sqrt(2 * Math.PI));
      },
    }),
  ),
  {
    name: 'NORM.S.DIST',
    category: C,
    syntax: 'NORM.S.DIST(z, cumulative)',
    description: 'Returns the standard normal distribution.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const z = num(args[0], ctx);
      if (isError(z)) return z;
      const cum = optBool(args[1], false);
      if (isError(cum)) return cum;
      return cum ? normSDist(z) : Math.exp((-z * z) / 2) / Math.sqrt(2 * Math.PI);
    },
  },
  {
    name: 'NORMSDIST',
    category: C,
    syntax: 'NORMSDIST(z)',
    description: 'Returns the standard normal cumulative distribution.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const z = num(args[0], ctx);
      return isError(z) ? z : normSDist(z);
    },
  },
  ...(['NORM.INV', 'NORMINV'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(probability, mean, standard_dev)`,
      description: 'Returns the inverse of the normal cumulative distribution.',
      minArgs: 3,
      maxArgs: 3,
      impl: (args, ctx) => {
        const v = nums(args, ctx, [undefined, undefined, undefined]);
        if (isError(v)) return v;
        const [p = 0, m = 0, s = 0] = v;
        if (p <= 0 || p >= 1 || s <= 0) return ERRORS.NUM;
        return m + s * normSInv(p);
      },
    }),
  ),
  ...(['NORM.S.INV', 'NORMSINV'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(probability)`,
      description: 'Returns the inverse of the standard normal cumulative distribution.',
      minArgs: 1,
      maxArgs: 1,
      impl: (args, ctx) => {
        const p = num(args[0], ctx);
        if (isError(p)) return p;
        return p <= 0 || p >= 1 ? ERRORS.NUM : normSInv(p);
      },
    }),
  ),
  ...(['EXPON.DIST', 'EXPONDIST'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(x, lambda, cumulative)`,
      description: 'Returns the exponential distribution.',
      minArgs: 3,
      maxArgs: 3,
      impl: (args, ctx) => {
        const v = nums(args, ctx, [undefined, undefined]);
        if (isError(v)) return v;
        const [x = 0, l = 0] = v;
        const cum = optBool(args[2], false);
        if (isError(cum)) return cum;
        if (x < 0 || l <= 0) return ERRORS.NUM;
        return cum ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x);
      },
    }),
  ),
  ...(['BINOM.DIST', 'BINOMDIST'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(number_s, trials, probability_s, cumulative)`,
      description: 'Returns the individual term binomial distribution probability.',
      minArgs: 4,
      maxArgs: 4,
      impl: (args, ctx) => {
        const v = nums(args, ctx, [undefined, undefined, undefined]);
        if (isError(v)) return v;
        const [s = 0, n = 0, prob = 0] = v;
        const cum = optBool(args[3], false);
        if (isError(cum)) return cum;
        const k = trunc(s);
        const trials = trunc(n);
        if (k < 0 || k > trials || prob < 0 || prob > 1) return ERRORS.NUM;
        const pmf = (i: number): number => combin(trials, i) * prob ** i * (1 - prob) ** (trials - i);
        if (!cum) return pmf(k);
        let total = 0;
        for (let i = 0; i <= k; i++) total += pmf(i);
        return total;
      },
    }),
  ),
  ...(['POISSON.DIST', 'POISSON'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(x, mean, cumulative)`,
      description: 'Returns the Poisson distribution.',
      minArgs: 3,
      maxArgs: 3,
      impl: (args, ctx) => {
        const v = nums(args, ctx, [undefined, undefined]);
        if (isError(v)) return v;
        const [x = 0, mean = 0] = v;
        const cum = optBool(args[2], false);
        if (isError(cum)) return cum;
        const k = trunc(x);
        if (k < 0 || mean < 0) return ERRORS.NUM;
        const pmf = (i: number): number => Math.exp(i * Math.log(mean) - mean - logFactorial(i));
        if (!cum) return mean === 0 ? (k === 0 ? 1 : 0) : pmf(k);
        let total = 0;
        for (let i = 0; i <= k; i++) total += mean === 0 ? (i === 0 ? 1 : 0) : pmf(i);
        return total;
      },
    }),
  ),
  {
    name: 'GAMMA',
    category: C,
    syntax: 'GAMMA(number)',
    description: 'Returns the Gamma function value.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const x = num(args[0], ctx);
      if (isError(x)) return x;
      if (x <= 0 && Number.isInteger(x)) return ERRORS.NUM;
      return checked(gamma(x));
    },
  },
  ...(['GAMMALN', 'GAMMALN.PRECISE'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(x)`,
      description: 'Returns the natural logarithm of the gamma function.',
      minArgs: 1,
      maxArgs: 1,
      impl: (args, ctx) => {
        const x = num(args[0], ctx);
        if (isError(x)) return x;
        return x <= 0 ? ERRORS.NUM : gammaLn(x);
      },
    }),
  ),
  {
    name: 'FISHER',
    category: C,
    syntax: 'FISHER(x)',
    description: 'Returns the Fisher transformation.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const x = num(args[0], ctx);
      if (isError(x)) return x;
      return x <= -1 || x >= 1 ? ERRORS.NUM : 0.5 * Math.log((1 + x) / (1 - x));
    },
  },
  {
    name: 'FISHERINV',
    category: C,
    syntax: 'FISHERINV(y)',
    description: 'Returns the inverse of the Fisher transformation.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const y = num(args[0], ctx);
      if (isError(y)) return y;
      const e2 = Math.exp(2 * y);
      return (e2 - 1) / (e2 + 1);
    },
  },
  ...(['CONFIDENCE.NORM', 'CONFIDENCE'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(alpha, standard_dev, size)`,
      description: 'Returns the confidence interval for a population mean, using a normal distribution.',
      minArgs: 3,
      maxArgs: 3,
      impl: (args, ctx) => {
        const v = nums(args, ctx, [undefined, undefined, undefined]);
        if (isError(v)) return v;
        const [a = 0, s = 0, n = 0] = v;
        const size = trunc(n);
        if (a <= 0 || a >= 1 || s <= 0 || size < 1) return ERRORS.NUM;
        return (normSInv(1 - a / 2) * s) / Math.sqrt(size);
      },
    }),
  ),
  {
    name: 'PERMUT',
    category: C,
    syntax: 'PERMUT(number, number_chosen)',
    description: 'Returns the number of permutations for a given number of objects.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const nn = trunc(v[0] ?? 0);
      const kk = trunc(v[1] ?? 0);
      if (nn < 0 || kk < 0 || kk > nn) return ERRORS.NUM;
      let r = 1;
      for (let i = 0; i < kk; i++) r *= nn - i;
      return checked(r);
    },
  },
  {
    name: 'PERMUTATIONA',
    category: C,
    syntax: 'PERMUTATIONA(number, number_chosen)',
    description: 'Returns the number of permutations with repetitions.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined]);
      if (isError(v)) return v;
      const nn = trunc(v[0] ?? 0);
      const kk = trunc(v[1] ?? 0);
      if (nn < 0 || kk < 0) return ERRORS.NUM;
      return checked(nn ** kk);
    },
  },
];

