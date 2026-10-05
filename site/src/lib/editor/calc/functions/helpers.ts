// Argument handling shared by the function families.

import { approxAdd, toBoolean, toNumber, toText } from '../coerce.ts';
import type { FnContext } from '../function-spec.ts';
import {
  type Area,
  type CalcArray,
  type CalcError,
  type CalcScalar,
  type CalcValue,
  ERRORS,
  isArray,
  isError,
  isLambda,
  isRef,
  makeArray,
} from '../types.ts';

/** An argument the evaluator already reduced to a scalar (`scalar` kind). */
export function scalar(v: CalcValue | undefined): CalcScalar {
  if (v === undefined) return null;
  if (isArray(v)) return v.data[0] ?? null;
  if (isRef(v) || isLambda(v)) return ERRORS.VALUE;
  return v;
}

export function num(v: CalcValue | undefined, ctx: FnContext): number | CalcError {
  return toNumber(scalar(v), ctx.host.date1904);
}

/** Optional numeric argument: omitted or blank → `fallback`. */
export function optNum(v: CalcValue | undefined, fallback: number, ctx: FnContext): number | CalcError {
  return v === undefined || v === null ? fallback : num(v, ctx);
}

export function str(v: CalcValue | undefined): string | CalcError {
  return toText(scalar(v));
}

export function bool(v: CalcValue | undefined): boolean | CalcError {
  return toBoolean(scalar(v));
}

export function optBool(v: CalcValue | undefined, fallback: boolean): boolean | CalcError {
  return v === undefined ? fallback : bool(v);
}

/** First error among the already-converted arguments. */
export function anyError(...values: unknown[]): CalcError | undefined {
  return values.find(isError);
}

export function checked(n: number): number | CalcError {
  return Number.isFinite(n) ? n : ERRORS.NUM;
}

export function arrayOf(rows: number, cols: number, fill: (r: number, c: number) => CalcScalar): CalcArray {
  const data: CalcScalar[] = new Array<CalcScalar>(rows * cols);
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) data[r * cols + c] = fill(r, c);
  return makeArray(rows, cols, data);
}

export function column(values: readonly CalcScalar[]): CalcArray {
  return makeArray(values.length, 1, values.slice());
}

/**
 * Every value the arguments contribute, row by row. References and arrays
 * yield their cells (`fromRange: true`); a direct scalar yields itself.
 * Returning false from `visit` stops the walk.
 */
export function forEachValue(
  args: readonly CalcValue[],
  ctx: FnContext,
  visit: (value: CalcScalar, fromRange: boolean) => boolean | void,
): void {
  for (const v of args) {
    if (isRef(v)) {
      for (const area of v.areas) {
        for (const x of ctx.host.areaValues(area).data) if (visit(x, true) === false) return;
      }
    } else if (isArray(v)) {
      for (const x of v.data) if (visit(x, true) === false) return;
    } else if (visit(isLambda(v) ? ERRORS.VALUE : v, false) === false) {
      return;
    }
  }
}

export type NumberMode =
  /** SUM / AVERAGE: numbers from ranges; direct arguments coerced (text that is not numeric is #VALUE!). */
  | 'plain'
  /** AVERAGEA / MAXA: ranges also count text as 0 and booleans as 1 / 0. */
  | 'a';

/** Collect the numbers aggregate functions see, or the first error encountered. */
export function collectNumbers(args: readonly CalcValue[], ctx: FnContext, mode: NumberMode = 'plain'): number[] | CalcError {
  const out: number[] = [];
  let error: CalcError | undefined;
  forEachValue(args, ctx, (x, fromRange) => {
    if (typeof x === 'number') {
      out.push(x);
    } else if (isError(x)) {
      error = x;
    } else if (fromRange) {
      if (mode === 'a' && x !== null) out.push(typeof x === 'boolean' ? (x ? 1 : 0) : 0);
    } else {
      const n = toNumber(x, ctx.host.date1904);
      if (isError(n)) error = n;
      else out.push(n);
    }
    return error === undefined;
  });
  return error ?? out;
}

/** Values of a `value`-kind argument flattened row-major. */
export function flatten(v: CalcValue, ctx: FnContext): CalcScalar[] {
  return ctx.toArray(v).data;
}

/** Numbers of an array-ish argument with positions kept (non-numbers become undefined). */
export function numericVector(v: CalcValue, ctx: FnContext): Array<number | undefined> | CalcError {
  const out: Array<number | undefined> = [];
  for (const x of flatten(v, ctx)) {
    if (isError(x)) return x;
    out.push(typeof x === 'number' ? x : undefined);
  }
  return out;
}

/** SUM's total: Excel rounds only the last addition, so SUM(0.1,0.2,-0.3) is 0 but SUM(0.1,0.2,-0.3,0) is not. */
export function sum(values: readonly number[]): number {
  let s = 0;
  const last = values.length - 1;
  for (let i = 0; i < last; i++) s += values[i] ?? 0;
  return last < 0 ? 0 : approxAdd(s, values[last] ?? 0);
}

/** Kahan-free mean and sample / population variance, two-pass for stability. */
export function variance(values: readonly number[], sample: boolean): number | CalcError {
  const n = values.length;
  if (n < (sample ? 2 : 1)) return ERRORS.DIV0;
  const mean = sum(values) / n;
  let ss = 0;
  for (const v of values) ss += (v - mean) ** 2;
  return ss / (sample ? n - 1 : n);
}

/** Truncate toward zero, the way Excel reads a fractional count or index. */
export function trunc(n: number): number {
  return n < 0 ? Math.ceil(n) : Math.floor(n);
}

/**
 * A range or array argument with its full logical size. Very large references
 * come back trimmed to the populated extent (see `EvalHost.areaValues`), so
 * positional code reads through `at`, which treats the trimmed-off tail as
 * blank.
 */
export interface Grid {
  readonly rows: number;
  readonly cols: number;
  readonly values: CalcArray;
  readonly area: Area | undefined;
  at(r: number, c: number): CalcScalar;
}

export function grid(v: CalcValue, ctx: FnContext): Grid | CalcError {
  if (isRef(v)) {
    if (v.areas.length !== 1) return ERRORS.VALUE;
    const area = v.areas[0];
    if (area === undefined) return ERRORS.REF;
    const values = ctx.host.areaValues(area);
    return makeGrid(values, area.r2 - area.r1 + 1, area.c2 - area.c1 + 1, area);
  }
  if (isError(v)) return v;
  const values = ctx.toArray(v);
  return makeGrid(values, values.rows, values.cols, undefined);
}

const makeGrid = (values: CalcArray, rows: number, cols: number, area: Area | undefined): Grid => ({
  rows,
  cols,
  values,
  area,
  at: (r, c) => (r < values.rows && c < values.cols ? (values.data[r * values.cols + c] ?? null) : null),
});

/**
 * Read leading numeric arguments in one go. `defaults[i]` is used when
 * argument i was omitted; `undefined` marks a required one. Returns the first
 * coercion error instead.
 */
export function nums(args: ReadonlyArray<CalcValue | undefined>, ctx: FnContext, defaults: ReadonlyArray<number | undefined>): number[] | CalcError {
  const out: number[] = [];
  for (let i = 0; i < defaults.length; i++) {
    const v = args[i];
    const fallback = defaults[i];
    if (v === undefined && fallback !== undefined) {
      out.push(fallback);
      continue;
    }
    const n = num(v, ctx);
    if (isError(n)) return n;
    out.push(n);
  }
  return out;
}
