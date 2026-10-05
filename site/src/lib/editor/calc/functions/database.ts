// Database functions: DSUM(database, field, criteria) and friends. The
// criteria range is a header row plus condition rows; conditions in one row
// must all hold (AND), any row may match (OR), and a blank condition matches
// everything.

import { toText } from '../coerce.ts';
import type { FnContext, FunctionSpec } from '../function-spec.ts';
import { type CalcScalar, type CalcValue, ERRORS, isError } from '../types.ts';
import { parseCriterion } from './criteria.ts';
import { grid, sum, variance } from './helpers.ts';
import { stdev } from './stats.ts';

const C = 'Database' as const;

const header = (v: CalcScalar): string => {
  const t = toText(v);
  return isError(t) ? '' : t.trim().toLowerCase();
};

/** Values of `field` in the database rows matching the criteria, or an error. */
const selected = (args: CalcValue[], ctx: FnContext, needField: boolean): CalcScalar[] | typeof ERRORS.VALUE => {
  const db = grid(args[0] ?? null, ctx);
  const crit = grid(args[2] ?? null, ctx);
  if (isError(db) || isError(crit) || db.rows < 1 || crit.rows < 1) return ERRORS.VALUE;
  const headers = Array.from({ length: db.cols }, (_, c) => header(db.at(0, c)));
  let fieldCol = -1;
  const f = args[1] ?? null;
  if (typeof f === 'number') fieldCol = Math.trunc(f) - 1;
  else if (typeof f === 'string') fieldCol = headers.indexOf(f.trim().toLowerCase());
  if (needField && (fieldCol < 0 || fieldCol >= db.cols)) return ERRORS.VALUE;

  const conditions: Array<Array<{ col: number; test: (v: CalcScalar) => boolean }>> = [];
  for (let r = 1; r < crit.rows; r++) {
    const row: Array<{ col: number; test: (v: CalcScalar) => boolean }> = [];
    for (let c = 0; c < crit.cols; c++) {
      const cond = crit.at(r, c);
      if (cond === null) continue;
      const col = headers.indexOf(header(crit.at(0, c)));
      if (col < 0) return ERRORS.VALUE;
      row.push({ col, test: parseCriterion(cond, ctx.host.date1904).test });
    }
    conditions.push(row);
  }
  const out: CalcScalar[] = [];
  for (let r = 1; r < db.rows; r++) {
    const match = conditions.length === 0 || conditions.some((row) => row.every((t) => t.test(db.at(r, t.col))));
    if (match) out.push(needField ? db.at(r, fieldCol) : db.at(r, 0));
  }
  return out;
};

const dfn = (name: string, description: string, fn: (values: CalcScalar[]) => CalcValue, needField = true): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(database, field, criteria)`,
  description,
  minArgs: 3,
  maxArgs: 3,
  args: ['ref', 'scalar', 'ref'],
  impl: (args, ctx) => {
    const values = selected(args, ctx, needField || (args[1] !== null && args[1] !== undefined));
    return isError(values) ? values : fn(values);
  },
});

const numbers = (values: readonly CalcScalar[]): number[] => values.filter((v): v is number => typeof v === 'number');

export const DATABASE_FUNCTIONS: FunctionSpec[] = [
  dfn('DSUM', 'Adds the numbers in the field column of records in the database that match the criteria.', (v) => sum(numbers(v))),
  dfn('DCOUNT', 'Counts the cells that contain numbers in a database.', (v) => numbers(v).length, false),
  dfn('DCOUNTA', 'Counts nonblank cells in a database.', (v) => v.filter((x) => x !== null).length, false),
  dfn('DAVERAGE', 'Returns the average of selected database entries.', (v) => {
    const n = numbers(v);
    return n.length === 0 ? ERRORS.DIV0 : sum(n) / n.length;
  }),
  dfn('DMAX', 'Returns the maximum value from selected database entries.', (v) => {
    const n = numbers(v);
    return n.length === 0 ? 0 : Math.max(...n);
  }),
  dfn('DMIN', 'Returns the minimum value from selected database entries.', (v) => {
    const n = numbers(v);
    return n.length === 0 ? 0 : Math.min(...n);
  }),
  dfn('DPRODUCT', 'Multiplies the values in a field of records that match the criteria.', (v) => {
    const n = numbers(v);
    return n.length === 0 ? 0 : n.reduce((a, b) => a * b, 1);
  }),
  dfn('DGET', 'Extracts from a database a single record that matches the specified criteria.', (v) =>
    v.length === 0 ? ERRORS.VALUE : v.length > 1 ? ERRORS.NUM : (v[0] ?? null),
  ),
  dfn('DSTDEV', 'Estimates the standard deviation based on a sample of selected database entries.', (v) => stdev(numbers(v), true)),
  dfn('DSTDEVP', 'Calculates the standard deviation based on the entire population of selected database entries.', (v) => stdev(numbers(v), false)),
  dfn('DVAR', 'Estimates variance based on a sample from selected database entries.', (v) => variance(numbers(v), true)),
  dfn('DVARP', 'Calculates variance based on the entire population of selected database entries.', (v) => variance(numbers(v), false)),
];
