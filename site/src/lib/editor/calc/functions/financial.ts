// Financial functions (annuities, cash flows, depreciation).

import type { FnContext, FunctionSpec } from '../function-spec.ts';
import { type CalcError, type CalcValue, ERRORS, isError } from '../types.ts';
import { checked, collectNumbers, flatten, nums, trunc } from './helpers.ts';

const C = 'Financial' as const;

export function pmt(r: number, n: number, present: number, future: number, type: number): number {
  if (r === 0) return -(present + future) / n;
  const f = (1 + r) ** n;
  return -(r * (future + present * f)) / ((1 + r * type) * (f - 1));
}

export function fv(r: number, n: number, payment: number, present: number, type: number): number {
  if (r === 0) return -(present + payment * n);
  const f = (1 + r) ** n;
  return -(present * f + (payment * (1 + r * type) * (f - 1)) / r);
}

function pv(r: number, n: number, payment: number, future: number, type: number): number {
  if (r === 0) return -(future + payment * n);
  const f = (1 + r) ** n;
  return -(future + (payment * (1 + r * type) * (f - 1)) / r) / f;
}

function ipmt(r: number, per: number, n: number, presentValue: number, future: number, type: number): number {
  const p = pmt(r, n, presentValue, future, type);
  if (per === 1 && type === 1) return 0;
  const interest = fv(r, per - 1, p, presentValue, type) * r;
  return type === 1 ? interest / (1 + r) : interest;
}

const STEP_TOLERANCE = 1e-10;

/** Newton's method from `guess`; undefined when it does not settle. */
function newton(f: (x: number) => number, guess: number): number | undefined {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    if (y === 0) return x;
    const h = Math.max(1e-7, Math.abs(x) * 1e-7);
    const slope = (f(x + h) - f(x - h)) / (2 * h);
    if (slope === 0 || !Number.isFinite(slope)) return undefined;
    const next = x - y / slope;
    if (!Number.isFinite(next) || next <= -1) return undefined;
    if (Math.abs(next - x) < STEP_TOLERANCE * Math.max(1, Math.abs(x))) return next;
    x = next;
  }
  return undefined;
}

/**
 * A rate in (-1, ∞) where `f` is zero, or undefined (Excel's #NUM!). Newton
 * from the caller's guess finds the root Excel reports when there are
 * several; when it wanders off (IRR({-100,10,10}) from 0.1 heads below -1) a
 * scan for the sign change nearest the guess, then bisection, still finds it.
 */
function solve(f: (x: number) => number, guess: number): number | undefined {
  const fast = newton(f, guess);
  if (fast !== undefined) return fast;
  const grid: number[] = [];
  for (let x = -0.99; x < 1; x += 0.01) grid.push(x);
  for (let x = 1; x <= 10; x += 0.1) grid.push(x);
  let best: [number, number] | undefined;
  for (let i = 1; i < grid.length; i++) {
    const a = grid[i - 1] ?? 0;
    const b = grid[i] ?? 0;
    if (Math.sign(f(a)) * Math.sign(f(b)) > 0) continue;
    if (best === undefined || Math.abs(a - guess) < Math.abs(best[0] - guess)) best = [a, b];
  }
  if (best === undefined) return undefined;
  let [lo, hi] = best;
  const loSign = Math.sign(f(lo));
  for (let i = 0; i < 200 && hi - lo > STEP_TOLERANCE; i++) {
    const mid = (lo + hi) / 2;
    if (Math.sign(f(mid)) === loSign) lo = mid;
    else hi = mid;
  }
  return newton(f, (lo + hi) / 2) ?? (lo + hi) / 2;
}

const npv = (r: number, values: readonly number[]): number => values.reduce((s, v, i) => s + v / (1 + r) ** (i + 1), 0);

/** Numbers of a cash-flow argument in order; text and blanks are skipped. */
const flows = (v: CalcValue | undefined, ctx: FnContext): number[] | CalcError => collectNumbers([v ?? null], ctx);

const annuity = (
  name: string,
  syntax: string,
  description: string,
  defaults: ReadonlyArray<number | undefined>,
  fn: (v: number[]) => number | CalcError,
): FunctionSpec => ({
  name,
  category: C,
  syntax,
  description,
  minArgs: defaults.filter((d) => d === undefined).length,
  maxArgs: defaults.length,
  impl: (args, ctx) => {
    const v = nums(args, ctx, defaults);
    if (isError(v)) return v;
    const r = fn(v);
    return isError(r) ? r : checked(r);
  },
});

const U = undefined;

export const FINANCIAL_FUNCTIONS: FunctionSpec[] = [
  annuity('PMT', 'PMT(rate, nper, pv, [fv], [type])', 'Calculates the payment for a loan based on constant payments and a constant interest rate.', [U, U, U, 0, 0], ([r = 0, n = 0, p = 0, f = 0, t = 0]) =>
    n === 0 ? ERRORS.NUM : pmt(r, n, p, f, t !== 0 ? 1 : 0),
  ),
  annuity('FV', 'FV(rate, nper, pmt, [pv], [type])', 'Returns the future value of an investment.', [U, U, U, 0, 0], ([r = 0, n = 0, p = 0, v = 0, t = 0]) => fv(r, n, p, v, t !== 0 ? 1 : 0)),
  annuity('PV', 'PV(rate, nper, pmt, [fv], [type])', 'Returns the present value of an investment.', [U, U, U, 0, 0], ([r = 0, n = 0, p = 0, f = 0, t = 0]) => pv(r, n, p, f, t !== 0 ? 1 : 0)),
  annuity('NPER', 'NPER(rate, pmt, pv, [fv], [type])', 'Returns the number of periods for an investment.', [U, U, U, 0, 0], ([r = 0, p = 0, v = 0, f = 0, t = 0]) => {
    const type = t !== 0 ? 1 : 0;
    if (r === 0) return p === 0 ? ERRORS.NUM : -(v + f) / p;
    const num = p * (1 + r * type) - f * r;
    const den = p * (1 + r * type) + v * r;
    if (num / den <= 0) return ERRORS.NUM;
    return Math.log(num / den) / Math.log(1 + r);
  }),
  annuity('IPMT', 'IPMT(rate, per, nper, pv, [fv], [type])', 'Returns the interest payment for an investment for a given period.', [U, U, U, U, 0, 0], ([r = 0, per = 0, n = 0, p = 0, f = 0, t = 0]) =>
    per < 1 || per > n ? ERRORS.NUM : ipmt(r, per, n, p, f, t !== 0 ? 1 : 0),
  ),
  annuity('PPMT', 'PPMT(rate, per, nper, pv, [fv], [type])', 'Returns the payment on the principal for an investment for a given period.', [U, U, U, U, 0, 0], ([r = 0, per = 0, n = 0, p = 0, f = 0, t = 0]) => {
    const type = t !== 0 ? 1 : 0;
    return per < 1 || per > n ? ERRORS.NUM : pmt(r, n, p, f, type) - ipmt(r, per, n, p, f, type);
  }),
  annuity('CUMIPMT', 'CUMIPMT(rate, nper, pv, start_period, end_period, type)', 'Returns the cumulative interest paid between two periods.', [U, U, U, U, U, U], ([r = 0, n = 0, p = 0, s = 0, e = 0, t = 0]) => {
    const start = Math.ceil(s);
    const end = Math.floor(e);
    if (r <= 0 || n <= 0 || p <= 0 || start < 1 || end < start || (t !== 0 && t !== 1)) return ERRORS.NUM;
    let total = 0;
    for (let i = start; i <= end; i++) total += ipmt(r, i, n, p, 0, t);
    return total;
  }),
  annuity('CUMPRINC', 'CUMPRINC(rate, nper, pv, start_period, end_period, type)', 'Returns the cumulative principal paid on a loan between two periods.', [U, U, U, U, U, U], ([r = 0, n = 0, p = 0, s = 0, e = 0, t = 0]) => {
    const start = Math.ceil(s);
    const end = Math.floor(e);
    if (r <= 0 || n <= 0 || p <= 0 || start < 1 || end < start || (t !== 0 && t !== 1)) return ERRORS.NUM;
    let total = 0;
    for (let i = start; i <= end; i++) total += pmt(r, n, p, 0, t) - ipmt(r, i, n, p, 0, t);
    return total;
  }),
  annuity('RATE', 'RATE(nper, pmt, pv, [fv], [type], [guess])', 'Returns the interest rate per period of an annuity.', [U, U, U, 0, 0, 0.1], ([n = 0, p = 0, v = 0, f = 0, t = 0, g = 0.1]) => {
    const type = t !== 0 ? 1 : 0;
    const r = solve((x) => (x === 0 ? v + p * n + f : v * (1 + x) ** n + (p * (1 + x * type) * ((1 + x) ** n - 1)) / x + f), g);
    return r === undefined ? ERRORS.NUM : r;
  }),
  annuity('SLN', 'SLN(cost, salvage, life)', 'Returns the straight-line depreciation of an asset for one period.', [U, U, U], ([c = 0, s = 0, l = 0]) => (l === 0 ? ERRORS.DIV0 : (c - s) / l)),
  annuity('SYD', 'SYD(cost, salvage, life, per)', "Returns the sum-of-years' digits depreciation of an asset for a specified period.", [U, U, U, U], ([c = 0, s = 0, l = 0, p = 0]) =>
    l <= 0 || p <= 0 || p > l ? ERRORS.NUM : ((c - s) * (l - p + 1) * 2) / (l * (l + 1)),
  ),
  annuity('DB', 'DB(cost, salvage, life, period, [month])', 'Returns the depreciation of an asset for a specified period using the fixed-declining balance method.', [U, U, U, U, 12], ([c = 0, s = 0, l = 0, p = 0, m = 12]) => {
    const life = trunc(l);
    const period = trunc(p);
    const month = trunc(m);
    if (c < 0 || s < 0 || life <= 0 || period <= 0 || month < 1 || month > 12 || period > life + 1) return ERRORS.NUM;
    if (c === 0) return 0;
    const rate = Math.round((1 - (s / c) ** (1 / life)) * 1000) / 1000;
    let total = 0;
    let dep = 0;
    for (let i = 1; i <= period; i++) {
      if (i === 1) dep = (c * rate * month) / 12;
      else if (i === life + 1) dep = ((c - total) * rate * (12 - month)) / 12;
      else dep = (c - total) * rate;
      total += dep;
    }
    return dep;
  }),
  annuity('DDB', 'DDB(cost, salvage, life, period, [factor])', 'Returns the depreciation of an asset for a specified period using the double-declining balance method.', [U, U, U, U, 2], ([c = 0, s = 0, l = 0, p = 0, f = 2]) => {
    if (c < 0 || s < 0 || l <= 0 || p <= 0 || p > l || f <= 0) return ERRORS.NUM;
    let acc = 0;
    let dep = 0;
    for (let i = 1; i <= Math.ceil(p); i++) {
      dep = Math.max(0, Math.min(((c - acc) * f) / l, c - s - acc));
      acc += dep;
    }
    return dep;
  }),
  annuity('EFFECT', 'EFFECT(nominal_rate, npery)', 'Returns the effective annual interest rate.', [U, U], ([n = 0, p = 0]) => {
    const periods = trunc(p);
    return n <= 0 || periods < 1 ? ERRORS.NUM : (1 + n / periods) ** periods - 1;
  }),
  annuity('NOMINAL', 'NOMINAL(effect_rate, npery)', 'Returns the annual nominal interest rate.', [U, U], ([e = 0, p = 0]) => {
    const periods = trunc(p);
    return e <= 0 || periods < 1 ? ERRORS.NUM : periods * ((1 + e) ** (1 / periods) - 1);
  }),
  {
    name: 'NPV',
    category: C,
    syntax: 'NPV(rate, value1, [value2], ...)',
    description: 'Returns the net present value of an investment based on periodic cash flows and a discount rate.',
    minArgs: 2,
    maxArgs: 255,
    args: ['scalar', 'ref'],
    impl: (args, ctx) => {
      const r = nums(args, ctx, [U]);
      if (isError(r)) return r;
      const values = collectNumbers(args.slice(1), ctx);
      if (isError(values)) return values;
      const rate = r[0] ?? 0;
      return rate === -1 ? ERRORS.DIV0 : npv(rate, values);
    },
  },
  {
    name: 'IRR',
    category: C,
    syntax: 'IRR(values, [guess])',
    description: 'Returns the internal rate of return for a series of cash flows.',
    minArgs: 1,
    maxArgs: 2,
    args: ['ref', 'scalar'],
    impl: (args, ctx) => {
      const values = flows(args[0], ctx);
      if (isError(values)) return values;
      const g = nums(args.slice(1), ctx, [0.1]);
      if (isError(g)) return g;
      if (!values.some((v) => v > 0) || !values.some((v) => v < 0)) return ERRORS.NUM;
      const r = solve((x) => values.reduce((s, v, i) => s + v / (1 + x) ** i, 0), g[0] ?? 0.1);
      return r === undefined ? ERRORS.NUM : r;
    },
  },
  {
    name: 'MIRR',
    category: C,
    syntax: 'MIRR(values, finance_rate, reinvest_rate)',
    description: 'Returns the internal rate of return where positive and negative cash flows are financed at different rates.',
    minArgs: 3,
    maxArgs: 3,
    args: ['ref', 'scalar', 'scalar'],
    impl: (args, ctx) => {
      const values = flows(args[0], ctx);
      if (isError(values)) return values;
      const rates = nums(args.slice(1), ctx, [U, U]);
      if (isError(rates)) return rates;
      const [fr = 0, rr = 0] = rates;
      const n = values.length;
      let pos = 0;
      let neg = 0;
      values.forEach((v, i) => {
        if (v > 0) pos += v * (1 + rr) ** (n - 1 - i);
        else neg += v / (1 + fr) ** i;
      });
      if (pos === 0 || neg === 0) return ERRORS.DIV0;
      return (-pos / neg) ** (1 / (n - 1)) - 1;
    },
  },
  {
    name: 'XNPV',
    category: C,
    syntax: 'XNPV(rate, values, dates)',
    description: 'Returns the net present value for a schedule of cash flows that is not necessarily periodic.',
    minArgs: 3,
    maxArgs: 3,
    args: ['scalar', 'value', 'value'],
    impl: (args, ctx) => {
      const r = nums(args, ctx, [U]);
      if (isError(r)) return r;
      const v = flatten(args[1] ?? null, ctx);
      const d = flatten(args[2] ?? null, ctx);
      if (v.length !== d.length) return ERRORS.NUM;
      const rate = r[0] ?? 0;
      const d0 = d[0];
      if (typeof d0 !== 'number') return ERRORS.VALUE;
      let total = 0;
      for (let i = 0; i < v.length; i++) {
        const value = v[i];
        const date = d[i];
        if (typeof value !== 'number' || typeof date !== 'number') return ERRORS.VALUE;
        if (Math.floor(date) < Math.floor(d0)) return ERRORS.NUM;
        total += value / (1 + rate) ** ((Math.floor(date) - Math.floor(d0)) / 365);
      }
      return checked(total);
    },
  },
  {
    name: 'XIRR',
    category: C,
    syntax: 'XIRR(values, dates, [guess])',
    description: 'Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic.',
    minArgs: 2,
    maxArgs: 3,
    args: ['value', 'value', 'scalar'],
    impl: (args, ctx) => {
      const v = flatten(args[0] ?? null, ctx);
      const d = flatten(args[1] ?? null, ctx);
      if (v.length !== d.length || v.length < 2) return ERRORS.NUM;
      const values: number[] = [];
      const days: number[] = [];
      for (let i = 0; i < v.length; i++) {
        const value = v[i];
        const date = d[i];
        if (typeof value !== 'number' || typeof date !== 'number') return ERRORS.VALUE;
        values.push(value);
        days.push(Math.floor(date));
      }
      if (!values.some((x) => x > 0) || !values.some((x) => x < 0)) return ERRORS.NUM;
      const g = nums(args.slice(2), ctx, [0.1]);
      if (isError(g)) return g;
      const d0 = days[0] ?? 0;
      const r = solve((x) => values.reduce((s, value, i) => s + value / (1 + x) ** (((days[i] ?? 0) - d0) / 365), 0), g[0] ?? 0.1);
      return r === undefined ? ERRORS.NUM : r;
    },
  },
];
