// SUMIF / COUNTIF / AVERAGEIF and their multi-criteria forms, MAXIFS / MINIFS.

import type { FnContext, FunctionSpec } from '../function-spec.ts';
import { type CalcError, type CalcValue, ERRORS, isError, isRef } from '../types.ts';
import { type Criterion, matchingPositions, parseCriterion } from './criteria.ts';
import { type Grid, grid, scalar } from './helpers.ts';

interface Condition {
  readonly grid: Grid;
  readonly criterion: Criterion;
}

interface Matches {
  readonly rows: number;
  readonly cols: number;
  /** Matching (row, col) offsets inside the ranges, flattened as row * cols + col. */
  readonly cells: number[];
  /** Matches in the blank tail trimmed off very large ranges (only COUNT needs them). */
  readonly blankTail: number;
}

const conditionsFrom = (pairs: ReadonlyArray<[CalcValue, CalcValue]>, ctx: FnContext): Condition[] | CalcError => {
  const out: Condition[] = [];
  for (const [range, crit] of pairs) {
    const g = grid(range, ctx);
    if (isError(g)) return g;
    out.push({ grid: g, criterion: parseCriterion(scalar(crit), ctx.host.date1904) });
  }
  return out;
};

const toCells = (g: Grid, positions: readonly number[]): number[] =>
  positions.map((i) => Math.floor(i / g.values.cols) * g.cols + (i % g.values.cols));

function match(conditions: readonly Condition[], ctx: FnContext): Matches | CalcError {
  const first = conditions[0];
  if (first === undefined) return ERRORS.VALUE;
  const { rows, cols } = first.grid;
  if (conditions.some((c) => c.grid.rows !== rows || c.grid.cols !== cols)) return ERRORS.VALUE;

  // Drive from the keyed condition with the fewest hits; test the rest per hit.
  let driver = first;
  let cells: number[] | undefined;
  for (const c of conditions) {
    if (c.criterion.key === undefined) continue;
    const hits = toCells(c.grid, matchingPositions(c.grid.values, c.criterion, ctx.host.date1904));
    if (cells === undefined || hits.length < cells.length) {
      cells = hits;
      driver = c;
    }
  }
  cells ??= toCells(first.grid, matchingPositions(first.grid.values, first.criterion, ctx.host.date1904));

  const others = conditions.filter((c) => c !== driver);
  const kept = others.length === 0 ? cells : cells.filter((cell) => {
    const r = Math.floor(cell / cols);
    const c = cell % cols;
    return others.every((o) => o.criterion.test(o.grid.at(r, c)));
  });
  const trimmed = driver.grid.values.data.length;
  const blankTail = conditions.every((c) => c.criterion.test(null)) ? rows * cols - trimmed : 0;
  return { rows, cols, cells: kept, blankTail };
}

/** The range whose values are aggregated, resized to the criteria range's shape like Excel does. */
const valueGrid = (v: CalcValue | undefined, rows: number, cols: number, ctx: FnContext): Grid | CalcError => {
  if (v === undefined) return ERRORS.VALUE;
  if (isRef(v)) {
    const a = v.areas[0];
    if (v.areas.length !== 1 || a === undefined) return ERRORS.VALUE;
    return grid({ kind: 'ref', areas: [{ ...a, r2: a.r1 + rows - 1, c2: a.c1 + cols - 1 }] }, ctx);
  }
  return grid(v, ctx);
};

type Reducer = 'sum' | 'count' | 'average' | 'max' | 'min';

const reduce = (kind: Reducer, m: Matches, values: Grid | undefined): CalcValue => {
  if (kind === 'count') return m.cells.length + m.blankTail;
  let total = 0;
  let n = 0;
  let best: number | undefined;
  for (const cell of m.cells) {
    const v = values?.at(Math.floor(cell / m.cols), cell % m.cols) ?? null;
    if (isError(v)) return v;
    if (typeof v !== 'number') continue;
    total += v;
    n++;
    if (best === undefined || (kind === 'max' ? v > best : v < best)) best = v;
  }
  switch (kind) {
    case 'sum':
      return total;
    case 'average':
      return n === 0 ? ERRORS.DIV0 : total / n;
    default:
      return best ?? 0;
  }
};

/** `xxxIF(range, criteria, [values])`. */
const single = (kind: Reducer) => (args: CalcValue[], ctx: FnContext): CalcValue => {
  const conditions = conditionsFrom([[args[0] ?? null, args[1] ?? null]], ctx);
  if (isError(conditions)) return conditions;
  const m = match(conditions, ctx);
  if (isError(m)) return m;
  if (kind === 'count') return reduce(kind, m, undefined);
  const values = args[2] === undefined || args[2] === null ? conditions[0]?.grid : valueGrid(args[2], m.rows, m.cols, ctx);
  if (isError(values)) return values;
  return reduce(kind, m, values);
};

/** `xxxIFS(values, range1, criteria1, ...)`; COUNTIFS has no values argument. */
const multi = (kind: Reducer) => (args: CalcValue[], ctx: FnContext): CalcValue => {
  const offset = kind === 'count' ? 0 : 1;
  const rest = args.slice(offset);
  if (rest.length === 0 || rest.length % 2 !== 0) return ERRORS.VALUE;
  const pairs: Array<[CalcValue, CalcValue]> = [];
  for (let i = 0; i < rest.length; i += 2) pairs.push([rest[i] ?? null, rest[i + 1] ?? null]);
  const conditions = conditionsFrom(pairs, ctx);
  if (isError(conditions)) return conditions;
  if (kind !== 'count') {
    const values = grid(args[0] ?? null, ctx);
    if (isError(values)) return values;
    const first = conditions[0];
    if (first !== undefined && (values.rows !== first.grid.rows || values.cols !== first.grid.cols)) return ERRORS.VALUE;
    const m = match(conditions, ctx);
    return isError(m) ? m : reduce(kind, m, values);
  }
  const m = match(conditions, ctx);
  return isError(m) ? m : reduce(kind, m, undefined);
};

export const CONDITIONAL_FUNCTIONS: FunctionSpec[] = [
  {
    name: 'SUMIF',
    category: 'Math & Trig',
    syntax: 'SUMIF(range, criteria, [sum_range])',
    description: 'Adds the cells specified by a given criteria.',
    minArgs: 2,
    maxArgs: 3,
    args: ['ref', 'scalar', 'ref'],
    impl: single('sum'),
  },
  {
    name: 'SUMIFS',
    category: 'Math & Trig',
    syntax: 'SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    description: 'Adds the cells in a range that meet multiple criteria.',
    minArgs: 3,
    maxArgs: 255,
    args: ['ref'],
    rest: ['ref', 'scalar'],
    impl: multi('sum'),
  },
  {
    name: 'COUNTIF',
    category: 'Statistical',
    syntax: 'COUNTIF(range, criteria)',
    description: 'Counts the number of cells within a range that meet the given criteria.',
    minArgs: 2,
    maxArgs: 2,
    args: ['ref', 'scalar'],
    impl: single('count'),
  },
  {
    name: 'COUNTIFS',
    category: 'Statistical',
    syntax: 'COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    description: 'Counts the number of cells within a range that meet multiple criteria.',
    minArgs: 2,
    maxArgs: 254,
    args: [],
    rest: ['ref', 'scalar'],
    impl: multi('count'),
  },
  {
    name: 'AVERAGEIF',
    category: 'Statistical',
    syntax: 'AVERAGEIF(range, criteria, [average_range])',
    description: 'Returns the average of all the cells in a range that meet a given criteria.',
    minArgs: 2,
    maxArgs: 3,
    args: ['ref', 'scalar', 'ref'],
    impl: single('average'),
  },
  {
    name: 'AVERAGEIFS',
    category: 'Statistical',
    syntax: 'AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    description: 'Returns the average of all cells that meet multiple criteria.',
    minArgs: 3,
    maxArgs: 255,
    args: ['ref'],
    rest: ['ref', 'scalar'],
    impl: multi('average'),
  },
  {
    name: 'MAXIFS',
    category: 'Statistical',
    syntax: 'MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    description: 'Returns the maximum value among cells specified by a given set of conditions.',
    minArgs: 3,
    maxArgs: 255,
    args: ['ref'],
    rest: ['ref', 'scalar'],
    impl: multi('max'),
  },
  {
    name: 'MINIFS',
    category: 'Statistical',
    syntax: 'MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    description: 'Returns the minimum value among cells specified by a given set of conditions.',
    minArgs: 3,
    maxArgs: 255,
    args: ['ref'],
    rest: ['ref', 'scalar'],
    impl: multi('min'),
  },
];
