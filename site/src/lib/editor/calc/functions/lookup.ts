// Lookup & Reference functions, including the dynamic-array shapers
// (FILTER, SORT, UNIQUE, TAKE, VSTACK…).

import { lettersFromColumn, quoteSheetName } from '../address.ts';
import { compareScalars, round15 } from '../coerce.ts';
import { fromStorageFormula } from '../storage.ts';
import { topLeft } from '../evaluator.ts';
import type { FnContext, FunctionSpec } from '../function-spec.ts';
import {
  type Area,
  type CalcArray,
  type CalcError,
  type CalcScalar,
  type CalcValue,
  ERRORS,
  isArray,
  isError,
  isRef,
  makeArray,
  MAX_COL,
  MAX_ROW,
} from '../types.ts';
import { hasWildcard, wildcardRegex } from './criteria.ts';
import { arrayOf, type Grid, grid, num, nums, optBool, scalar, str, trunc } from './helpers.ts';

const C = 'Lookup & Reference' as const;

// ---- matching --------------------------------------------------------------

type Comparable = Exclude<CalcScalar, CalcError | null>;

const sameType = (a: CalcScalar, b: Comparable): a is Comparable => a !== null && !isError(a) && typeof a === typeof b;

/** Exact-match predicate; text matches case-insensitively and, when allowed, with wildcards. */
const exactMatcher = (target: Comparable, wildcards: boolean): ((v: CalcScalar) => boolean) => {
  if (typeof target === 'string') {
    if (wildcards && hasWildcard(target)) {
      const re = wildcardRegex(target);
      return (v) => typeof v === 'string' && re.test(v);
    }
    const lower = target.toLowerCase();
    return (v) => typeof v === 'string' && v.toLowerCase() === lower;
  }
  if (typeof target === 'number') {
    const t = round15(target);
    return (v) => typeof v === 'number' && round15(v) === t;
  }
  return (v) => v === target;
};

/**
 * Excel's sorted-lookup binary search: the last position whose value is <=
 * target (or >= for descending data). Cells of another type, blanks and
 * errors are stepped over, which is what makes `LOOKUP(2, 1/(cond), …)`
 * return the last match.
 */
const binarySearch = (n: number, get: (i: number) => CalcScalar, target: Comparable, descending: boolean): number => {
  let lo = 0;
  let hi = n - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let m = mid;
    while (m >= lo && !sameType(get(m), target)) m--;
    if (m < lo) {
      lo = mid + 1;
      continue;
    }
    const v = get(m);
    const c = sameType(v, target) ? compareScalars(v, target) : 0;
    if (descending ? c >= 0 : c <= 0) {
      best = m;
      lo = mid + 1;
    } else {
      hi = m - 1;
    }
  }
  return best;
};

/** A one-dimensional view over a range or array (a single row or column of it). */
interface Vector {
  readonly length: number;
  at(i: number): CalcScalar;
}

const vectorOf = (g: Grid, orientation: 'row' | 'col' | 'auto', index = 0): Vector => {
  const byCol = orientation === 'col' || (orientation === 'auto' && g.cols === 1) || (orientation === 'auto' && g.rows !== 1);
  return byCol ? { length: g.rows, at: (i) => g.at(i, index) } : { length: g.cols, at: (i) => g.at(index, i) };
};

/** MATCH-style search returning a 0-based position or -1. `mode`: 0 exact, 1 sorted ≤, -1 sorted ≥ (descending). */
const matchIn = (vec: Vector, target: CalcScalar, mode: number): number => {
  if (target === null) return -1;
  if (isError(target)) return -1;
  if (mode === 0) {
    const test = exactMatcher(target, true);
    for (let i = 0; i < vec.length; i++) if (test(vec.at(i))) return i;
    return -1;
  }
  return binarySearch(vec.length, (i) => vec.at(i), target, mode < 0);
};

const lookupTarget = (v: CalcValue | undefined): CalcScalar => {
  const s = scalar(v);
  // A blank lookup value searches for 0 / "" the way Excel coerces it.
  return s === null ? 0 : s;
};

// ---- reference helpers -----------------------------------------------------

const singleArea = (v: CalcValue | undefined): Area | CalcError => {
  if (v === undefined) return ERRORS.VALUE;
  if (isError(v)) return v;
  if (!isRef(v) || v.areas.length !== 1) return ERRORS.VALUE;
  return v.areas[0] ?? ERRORS.REF;
};

const refOf = (area: Area): CalcValue => ({ kind: 'ref', areas: [area] });

const validArea = (a: Area): boolean => a.r1 >= 1 && a.c1 >= 1 && a.r2 <= MAX_ROW && a.c2 <= MAX_COL && a.r1 <= a.r2 && a.c1 <= a.c2;

/** Rows `[r1, r2)` × cols `[c1, c2)` of a grid, as a reference when it came from one. */
const slice = (g: Grid, r1: number, r2: number, c1: number, c2: number): CalcValue => {
  if (g.area !== undefined) {
    return refOf({ sheet: g.area.sheet, r1: g.area.r1 + r1, r2: g.area.r1 + r2 - 1, c1: g.area.c1 + c1, c2: g.area.c1 + c2 - 1 });
  }
  return arrayOf(r2 - r1, c2 - c1, (r, c) => g.at(r1 + r, c1 + c));
};

const rowsOf = (a: CalcArray): CalcScalar[][] => Array.from({ length: a.rows }, (_, r) => a.data.slice(r * a.cols, (r + 1) * a.cols));

/** Excel has no empty array: an empty result is #CALC!. */
const fromRows = (rows: readonly CalcScalar[][], cols: number): CalcArray | CalcError =>
  rows.length === 0 ? ERRORS.CALC : makeArray(rows.length, cols, rows.flat());

const transpose = (a: CalcArray): CalcArray => arrayOf(a.cols, a.rows, (r, c) => a.data[c * a.cols + r] ?? null);

/** Sort ordering for SORT / SORTBY: Excel's type order with blanks last. */
const sortCompare = (a: CalcScalar, b: CalcScalar): number => {
  if (a === null || b === null) return a === null ? (b === null ? 0 : 1) : -1;
  if (isError(a) || isError(b)) return isError(a) ? (isError(b) ? 0 : 1) : -1;
  return compareScalars(a, b);
};

const scalarKey = (v: CalcScalar): string => {
  if (v === null) return 'b';
  if (isError(v)) return `e${v.code}`;
  if (typeof v === 'string') return `s${v.toLowerCase()}`;
  return `${typeof v}${String(v)}`;
};

// ---- specs -----------------------------------------------------------------

const vhlookup = (vertical: boolean) => (args: CalcValue[], ctx: FnContext): CalcValue => {
  const target = lookupTarget(args[0]);
  if (isError(target)) return target;
  const table = grid(args[1] ?? null, ctx);
  if (isError(table)) return table;
  const idx = num(args[2], ctx);
  if (isError(idx)) return idx;
  const n = trunc(idx);
  const approx = optBool(args[3], true);
  if (isError(approx)) return approx;
  if (n < 1) return ERRORS.VALUE;
  if (n > (vertical ? table.cols : table.rows)) return ERRORS.REF;
  const key = vertical ? vectorOf(table, 'col', 0) : vectorOf(table, 'row', 0);
  const pos = matchIn(key, target, approx ? 1 : 0);
  if (pos < 0) return ERRORS.NA;
  return vertical ? table.at(pos, n - 1) : table.at(n - 1, pos);
};

/** XLOOKUP / XMATCH search. Returns the 0-based position or -1. */
const xsearch = (vec: Vector, target: CalcScalar, matchMode: number, searchMode: number): number => {
  if (target === null || isError(target)) return -1;
  if (searchMode === 2 || searchMode === -2) {
    const pos = binarySearch(vec.length, (i) => vec.at(i), target, searchMode === -2);
    if (matchMode === 0) return pos >= 0 && exactMatcher(target, false)(vec.at(pos)) ? pos : -1;
    if (pos >= 0 && exactMatcher(target, false)(vec.at(pos))) return pos;
    if (matchMode === -1) return pos;
    const next = pos + 1;
    return next < vec.length ? next : -1;
  }
  const order = searchMode === -1 ? Array.from({ length: vec.length }, (_, i) => vec.length - 1 - i) : Array.from({ length: vec.length }, (_, i) => i);
  const exact = exactMatcher(target, matchMode === 2);
  let best = -1;
  let bestValue: Comparable | undefined;
  for (const i of order) {
    const v = vec.at(i);
    if (exact(v)) return i;
    if ((matchMode === -1 || matchMode === 1) && sameType(v, target)) {
      const c = compareScalars(v, target);
      const better = bestValue === undefined || (matchMode === -1 ? compareScalars(v, bestValue) > 0 : compareScalars(v, bestValue) < 0);
      if (((matchMode === -1 && c < 0) || (matchMode === 1 && c > 0)) && better) {
        best = i;
        bestValue = v;
      }
    }
  }
  return best;
};

export const LOOKUP_FUNCTIONS: FunctionSpec[] = [
  {
    name: 'VLOOKUP',
    category: C,
    syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    description: 'Looks for a value in the leftmost column of a table, and returns a value in the same row from a column you specify.',
    minArgs: 3,
    maxArgs: 4,
    args: ['scalar', 'ref', 'scalar', 'scalar'],
    impl: vhlookup(true),
  },
  {
    name: 'HLOOKUP',
    category: C,
    syntax: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
    description: 'Looks for a value in the top row of a table, and returns a value in the same column from a row you specify.',
    minArgs: 3,
    maxArgs: 4,
    args: ['scalar', 'ref', 'scalar', 'scalar'],
    impl: vhlookup(false),
  },
  {
    name: 'LOOKUP',
    category: C,
    syntax: 'LOOKUP(lookup_value, lookup_vector, [result_vector])',
    description: 'Looks up a value either from a one-row or one-column range or from an array.',
    minArgs: 2,
    maxArgs: 3,
    args: ['scalar', 'ref', 'ref'],
    impl: (args, ctx) => {
      const target = lookupTarget(args[0]);
      if (isError(target)) return target;
      const look = grid(args[1] ?? null, ctx);
      if (isError(look)) return look;
      // Array form: search the first column (or row, for wide arrays), return from the last.
      const wide = look.cols > look.rows;
      const key = vectorOf(look, wide ? 'row' : 'col', 0);
      const pos = matchIn(key, target, 1);
      if (pos < 0) return ERRORS.NA;
      if (args[2] === undefined) return wide ? look.at(look.rows - 1, pos) : look.at(pos, look.cols - 1);
      const result = grid(args[2], ctx);
      if (isError(result)) return result;
      const out = vectorOf(result, 'auto');
      return pos < out.length ? out.at(pos) : ERRORS.NA;
    },
  },
  {
    name: 'MATCH',
    category: C,
    syntax: 'MATCH(lookup_value, lookup_array, [match_type])',
    description: 'Returns the relative position of an item in an array that matches a specified value.',
    minArgs: 2,
    maxArgs: 3,
    args: ['scalar', 'ref', 'scalar'],
    impl: (args, ctx) => {
      const target = lookupTarget(args[0]);
      if (isError(target)) return target;
      const look = grid(args[1] ?? null, ctx);
      if (isError(look)) return look;
      if (look.rows > 1 && look.cols > 1) return ERRORS.NA;
      const mode = args[2] === undefined ? 1 : num(args[2], ctx);
      if (isError(mode)) return mode;
      const pos = matchIn(vectorOf(look, 'auto'), target, Math.sign(trunc(mode)));
      return pos < 0 ? ERRORS.NA : pos + 1;
    },
  },
  {
    name: 'XMATCH',
    category: C,
    syntax: 'XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])',
    description: 'Returns the relative position of an item in an array or range of cells.',
    minArgs: 2,
    maxArgs: 4,
    args: ['scalar', 'ref', 'scalar', 'scalar'],
    impl: (args, ctx) => {
      const target = lookupTarget(args[0]);
      if (isError(target)) return target;
      const look = grid(args[1] ?? null, ctx);
      if (isError(look)) return look;
      if (look.rows > 1 && look.cols > 1) return ERRORS.VALUE;
      const v = nums(args.slice(2), ctx, [0, 1]);
      if (isError(v)) return v;
      const pos = xsearch(vectorOf(look, 'auto'), target, trunc(v[0] ?? 0), trunc(v[1] ?? 1));
      return pos < 0 ? ERRORS.NA : pos + 1;
    },
  },
  {
    name: 'XLOOKUP',
    category: C,
    syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])',
    description: 'Searches a range or an array, and returns an item corresponding to the first match it finds.',
    minArgs: 3,
    maxArgs: 6,
    args: ['scalar', 'ref', 'ref', 'value', 'scalar', 'scalar'],
    impl: (args, ctx) => {
      const target = lookupTarget(args[0]);
      if (isError(target)) return target;
      const look = grid(args[1] ?? null, ctx);
      if (isError(look)) return look;
      const result = grid(args[2] ?? null, ctx);
      if (isError(result)) return result;
      if (look.rows > 1 && look.cols > 1) return ERRORS.VALUE;
      const vertical = look.cols === 1 && look.rows > 1 ? true : look.rows === 1 && look.cols > 1 ? false : result.rows >= result.cols;
      if (vertical ? result.rows !== look.rows : result.cols !== look.cols) return ERRORS.VALUE;
      const v = nums(args.slice(4), ctx, [0, 1]);
      if (isError(v)) return v;
      const pos = xsearch(vectorOf(look, vertical ? 'col' : 'row'), target, trunc(v[0] ?? 0), trunc(v[1] ?? 1));
      if (pos < 0) return args[3] === undefined ? ERRORS.NA : (args[3] ?? null);
      return vertical ? slice(result, pos, pos + 1, 0, result.cols) : slice(result, 0, result.rows, pos, pos + 1);
    },
  },
  {
    name: 'INDEX',
    category: C,
    syntax: 'INDEX(array, row_num, [column_num], [area_num])',
    description: 'Returns a value or reference of the cell at the intersection of a particular row and column, in a given range.',
    minArgs: 2,
    maxArgs: 4,
    args: ['ref', 'scalar', 'scalar', 'scalar'],
    impl: (args, ctx) => {
      const source = args[0] ?? null;
      const v = nums(args.slice(1), ctx, [undefined, 0, 1]);
      if (isError(v)) return v;
      let [r = 0, c = 0] = v.map(trunc);
      const areaNum = trunc(v[2] ?? 1);
      let target: CalcValue = source;
      if (isRef(source)) {
        const area = source.areas[areaNum - 1];
        if (area === undefined) return ERRORS.REF;
        target = refOf(area);
      }
      const g = grid(target, ctx);
      if (isError(g)) return g;
      // With one index on a single row, that index picks the column.
      if (args[2] === undefined && g.rows === 1 && g.cols > 1) [r, c] = [1, r];
      if (r < 0 || c < 0 || r > g.rows || c > g.cols) return ERRORS.REF;
      if (r === 0 && c === 0) return slice(g, 0, g.rows, 0, g.cols);
      if (r === 0) return slice(g, 0, g.rows, c - 1, c);
      if (c === 0) {
        if (g.cols === 1) return slice(g, r - 1, r, 0, 1);
        return slice(g, r - 1, r, 0, g.cols);
      }
      return slice(g, r - 1, r, c - 1, c);
    },
  },
  {
    name: 'OFFSET',
    category: C,
    syntax: 'OFFSET(reference, rows, cols, [height], [width])',
    description: 'Returns a reference to a range that is a given number of rows and columns from a given reference.',
    minArgs: 3,
    maxArgs: 5,
    args: ['ref', 'scalar', 'scalar', 'scalar', 'scalar'],
    volatile: true,
    impl: (args, ctx) => {
      const base = singleArea(args[0]);
      if (isError(base)) return base;
      const v = nums(args.slice(1), ctx, [undefined, undefined, base.r2 - base.r1 + 1, base.c2 - base.c1 + 1]);
      if (isError(v)) return v;
      const [dr = 0, dc = 0, h = 1, w = 1] = v.map(trunc);
      if (h === 0 || w === 0) return ERRORS.REF;
      const r1 = base.r1 + dr;
      const c1 = base.c1 + dc;
      const area = {
        sheet: base.sheet,
        r1: Math.min(r1, r1 + h + (h > 0 ? -1 : 1)),
        r2: Math.max(r1, r1 + h + (h > 0 ? -1 : 1)),
        c1: Math.min(c1, c1 + w + (w > 0 ? -1 : 1)),
        c2: Math.max(c1, c1 + w + (w > 0 ? -1 : 1)),
      };
      return validArea(area) ? refOf(area) : ERRORS.REF;
    },
  },
  {
    name: 'INDIRECT',
    category: C,
    syntax: 'INDIRECT(ref_text, [a1])',
    description: 'Returns the reference specified by a text string.',
    minArgs: 1,
    maxArgs: 2,
    volatile: true,
    impl: (args, ctx) => {
      const text = str(args[0]);
      if (isError(text)) return text;
      const a1 = optBool(args[1], true);
      if (isError(a1)) return a1;
      return ctx.referenceFromText(text, a1) ?? ERRORS.REF;
    },
  },
  ...(['ROW', 'COLUMN'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}([reference])`,
      description: name === 'ROW' ? 'Returns the row number of a reference.' : 'Returns the column number of a reference.',
      minArgs: 0,
      maxArgs: 1,
      args: ['ref'],
      impl: (args, ctx) => {
        if (args[0] === undefined) return name === 'ROW' ? ctx.row : ctx.col;
        const a = singleArea(args[0]);
        if (isError(a)) return a;
        if (name === 'ROW') return a.r1 === a.r2 ? a.r1 : arrayOf(a.r2 - a.r1 + 1, 1, (r) => a.r1 + r);
        return a.c1 === a.c2 ? a.c1 : arrayOf(1, a.c2 - a.c1 + 1, (_, c) => a.c1 + c);
      },
    }),
  ),
  ...(['ROWS', 'COLUMNS'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array)`,
      description: name === 'ROWS' ? 'Returns the number of rows in a reference or array.' : 'Returns the number of columns in a reference or array.',
      minArgs: 1,
      maxArgs: 1,
      args: ['ref'],
      impl: (args, ctx) => {
        const v = args[0] ?? null;
        if (isError(v)) return v;
        if (isRef(v)) {
          const a = singleArea(v);
          if (isError(a)) return a;
          return name === 'ROWS' ? a.r2 - a.r1 + 1 : a.c2 - a.c1 + 1;
        }
        const arr = ctx.toArray(v);
        return name === 'ROWS' ? arr.rows : arr.cols;
      },
    }),
  ),
  {
    name: 'AREAS',
    category: C,
    syntax: 'AREAS(reference)',
    description: 'Returns the number of areas in a reference.',
    minArgs: 1,
    maxArgs: 1,
    args: ['ref'],
    impl: (args) => {
      const v = args[0] ?? null;
      if (isError(v)) return v;
      return isRef(v) ? v.areas.length : ERRORS.VALUE;
    },
  },
  {
    name: 'ADDRESS',
    category: C,
    syntax: 'ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])',
    description: 'Creates a cell reference as text, given specified row and column numbers.',
    minArgs: 2,
    maxArgs: 5,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined, 1]);
      if (isError(v)) return v;
      const [row = 0, col = 0, abs = 1] = v.map(trunc);
      const a1 = optBool(args[3], true);
      if (isError(a1)) return a1;
      if (row < 1 || row > MAX_ROW || col < 1 || col > MAX_COL || abs < 1 || abs > 4) return ERRORS.VALUE;
      const rowAbs = abs === 1 || abs === 2;
      const colAbs = abs === 1 || abs === 3;
      let ref: string;
      if (a1) {
        ref = `${colAbs ? '$' : ''}${lettersFromColumn(col)}${rowAbs ? '$' : ''}${row}`;
      } else {
        ref = `R${rowAbs ? row : `[${row}]`}C${colAbs ? col : `[${col}]`}`;
      }
      if (args[4] === undefined || args[4] === null) return ref;
      const sheet = str(args[4]);
      if (isError(sheet)) return sheet;
      return `${quoteSheetName(sheet)}!${ref}`;
    },
  },
  {
    name: 'CHOOSE',
    category: C,
    syntax: 'CHOOSE(index_num, value1, [value2], ...)',
    description: 'Chooses a value from a list of values based on an index number.',
    minArgs: 2,
    maxArgs: 255,
    lazy: true,
    impl: (args, ctx) => {
      const idx = ctx.deref(args[0]?.() ?? null);
      const pick = (v: CalcScalar): number | CalcError => {
        if (isError(v)) return v;
        const n = num(v, ctx);
        if (isError(n)) return n;
        const i = trunc(n);
        return i < 1 || i >= args.length ? ERRORS.VALUE : i;
      };
      if (isArray(idx)) {
        const choices = args.slice(1).map((t) => ctx.deref(t()));
        return arrayOf(idx.rows, idx.cols, (r, c) => {
          const i = pick(idx.data[r * idx.cols + c] ?? null);
          if (isError(i)) return i;
          return topLeft(choices[i - 1] ?? ERRORS.VALUE);
        });
      }
      const i = pick(idx);
      if (isError(i)) return i;
      return args[i]?.() ?? ERRORS.VALUE;
    },
  },
  {
    name: 'TRANSPOSE',
    category: C,
    syntax: 'TRANSPOSE(array)',
    description: 'Converts a vertical range of cells to a horizontal range, or vice versa.',
    minArgs: 1,
    maxArgs: 1,
    args: ['value'],
    impl: (args, ctx) => transpose(ctx.toArray(args[0] ?? null)),
  },
  {
    name: 'FORMULATEXT',
    category: C,
    syntax: 'FORMULATEXT(reference)',
    description: 'Returns the formula at the given reference as text.',
    minArgs: 1,
    maxArgs: 1,
    args: ['ref'],
    impl: (args, ctx) => {
      const a = singleArea(args[0]);
      if (isError(a)) return a;
      const text = ctx.host.formulaText(a.sheet, a.r1, a.c1);
      // Excel's FORMULATEXT shows the formula as typed, without the file's _xlfn. prefixes.
      return text === undefined ? ERRORS.NA : `=${fromStorageFormula(text)}`;
    },
  },
  {
    name: 'HYPERLINK',
    category: C,
    syntax: 'HYPERLINK(link_location, [friendly_name])',
    description: 'Creates a shortcut that jumps to a document or web page; the cell shows the friendly name.',
    minArgs: 1,
    maxArgs: 2,
    impl: (args) => (args[1] === undefined ? scalar(args[0]) : scalar(args[1])),
  },
  {
    name: 'FILTER',
    category: C,
    syntax: 'FILTER(array, include, [if_empty])',
    description: 'Filters a range or array based on criteria you define.',
    minArgs: 2,
    maxArgs: 3,
    args: ['value', 'value', 'value'],
    impl: (args, ctx) => {
      const a = ctx.toArray(args[0] ?? null);
      const inc = ctx.toArray(args[1] ?? null);
      const byRows = inc.cols === 1 && inc.rows === a.rows;
      const byCols = inc.rows === 1 && inc.cols === a.cols;
      if (!byRows && !byCols) return ERRORS.VALUE;
      const keep: number[] = [];
      for (let i = 0; i < inc.data.length; i++) {
        const v = inc.data[i] ?? null;
        if (isError(v)) return v;
        const b = typeof v === 'string' ? ERRORS.VALUE : v === null ? false : Boolean(v);
        if (isError(b)) return b;
        if (b) keep.push(i);
      }
      if (keep.length === 0) return args[2] === undefined ? ERRORS.CALC : (args[2] ?? null);
      if (byRows) return arrayOf(keep.length, a.cols, (r, c) => a.data[(keep[r] ?? 0) * a.cols + c] ?? null);
      return arrayOf(a.rows, keep.length, (r, c) => a.data[r * a.cols + (keep[c] ?? 0)] ?? null);
    },
  },
  {
    name: 'SORT',
    category: C,
    syntax: 'SORT(array, [sort_index], [sort_order], [by_col])',
    description: 'Sorts the contents of a range or array.',
    minArgs: 1,
    maxArgs: 4,
    args: ['value', 'value', 'value', 'scalar'],
    impl: (args, ctx) => {
      const byCol = optBool(args[3], false);
      if (isError(byCol)) return byCol;
      const source = ctx.toArray(args[0] ?? null);
      const a = byCol ? transpose(source) : source;
      const indexes = args[1] === undefined || args[1] === null ? [1] : ctx.toArray(args[1]).data.map((v) => (typeof v === 'number' ? trunc(v) : Number.NaN));
      const orders = args[2] === undefined || args[2] === null ? [1] : ctx.toArray(args[2]).data.map((v) => (typeof v === 'number' ? trunc(v) : Number.NaN));
      if (indexes.some((i) => !(i >= 1 && i <= a.cols)) || orders.some((o) => o !== 1 && o !== -1)) return ERRORS.VALUE;
      const rows = rowsOf(a);
      rows.sort((x, y) => {
        for (let k = 0; k < indexes.length; k++) {
          const i = (indexes[k] ?? 1) - 1;
          const o = orders[Math.min(k, orders.length - 1)] ?? 1;
          const c = sortCompare(x[i] ?? null, y[i] ?? null);
          if (c !== 0) return x[i] === null || y[i] === null ? c : c * o;
        }
        return 0;
      });
      const sorted = fromRows(rows, a.cols);
      return byCol && !isError(sorted) ? transpose(sorted) : sorted;
    },
  },
  {
    name: 'SORTBY',
    category: C,
    syntax: 'SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...)',
    description: 'Sorts the contents of a range or array based on the values in a corresponding range or array.',
    minArgs: 2,
    maxArgs: 255,
    args: ['value'],
    rest: ['value', 'scalar'],
    impl: (args, ctx) => {
      const a = ctx.toArray(args[0] ?? null);
      const keys: Array<{ values: CalcScalar[]; order: number; byCol: boolean }> = [];
      for (let i = 1; i < args.length; i += 2) {
        const by = ctx.toArray(args[i] ?? null);
        const o = args[i + 1] === undefined ? 1 : num(args[i + 1], ctx);
        if (isError(o)) return o;
        if (o !== 1 && o !== -1) return ERRORS.VALUE;
        const byCol = by.rows === 1 && by.cols === a.cols && a.cols > 1;
        if (!byCol && !(by.cols === 1 && by.rows === a.rows)) return ERRORS.VALUE;
        keys.push({ values: by.data, order: o, byCol });
      }
      const byCol = keys[0]?.byCol ?? false;
      const n = byCol ? a.cols : a.rows;
      const order = Array.from({ length: n }, (_, i) => i);
      order.sort((x, y) => {
        for (const k of keys) {
          const c = sortCompare(k.values[x] ?? null, k.values[y] ?? null);
          if (c !== 0) return c * k.order;
        }
        return 0;
      });
      return byCol
        ? arrayOf(a.rows, a.cols, (r, c) => a.data[r * a.cols + (order[c] ?? 0)] ?? null)
        : arrayOf(a.rows, a.cols, (r, c) => a.data[(order[r] ?? 0) * a.cols + c] ?? null);
    },
  },
  {
    name: 'UNIQUE',
    category: C,
    syntax: 'UNIQUE(array, [by_col], [exactly_once])',
    description: 'Returns a list of unique values in a list or range.',
    minArgs: 1,
    maxArgs: 3,
    args: ['value', 'scalar', 'scalar'],
    impl: (args, ctx) => {
      const byCol = optBool(args[1], false);
      if (isError(byCol)) return byCol;
      const once = optBool(args[2], false);
      if (isError(once)) return once;
      const source = ctx.toArray(args[0] ?? null);
      const a = byCol ? transpose(source) : source;
      const counts = new Map<string, { row: CalcScalar[]; n: number }>();
      for (const row of rowsOf(a)) {
        const key = row.map(scalarKey).join('\u0000');
        const hit = counts.get(key);
        if (hit === undefined) counts.set(key, { row, n: 1 });
        else hit.n++;
      }
      const rows = [...counts.values()].filter((x) => !once || x.n === 1).map((x) => x.row);
      const out = fromRows(rows, a.cols);
      return byCol && !isError(out) ? transpose(out) : out;
    },
  },
  ...(['TAKE', 'DROP'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array, rows, [columns])`,
      description: name === 'TAKE' ? 'Returns a specified number of contiguous rows or columns from the start or end of an array.' : 'Excludes a specified number of rows or columns from the start or end of an array.',
      minArgs: 2,
      maxArgs: 3,
      args: ['value', 'scalar', 'scalar'],
      impl: (args, ctx) => {
        const a = ctx.toArray(args[0] ?? null);
        const span = (v: CalcValue | undefined, size: number): [number, number] | CalcError => {
          if (v === undefined || v === null) return [0, size];
          const n = num(v, ctx);
          if (isError(n)) return n;
          const k = trunc(n);
          if (name === 'TAKE') return k >= 0 ? [0, Math.min(k, size)] : [Math.max(0, size + k), size];
          return k >= 0 ? [Math.min(k, size), size] : [0, Math.max(0, size + k)];
        };
        const rs = span(args[1], a.rows);
        if (isError(rs)) return rs;
        const cs = span(args[2], a.cols);
        if (isError(cs)) return cs;
        const rows = rs[1] - rs[0];
        const cols = cs[1] - cs[0];
        if (rows <= 0 || cols <= 0) return ERRORS.CALC;
        return arrayOf(rows, cols, (r, c) => a.data[(rs[0] + r) * a.cols + cs[0] + c] ?? null);
      },
    }),
  ),
  ...(['CHOOSECOLS', 'CHOOSEROWS'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array, ${name === 'CHOOSECOLS' ? 'col_num1, [col_num2]' : 'row_num1, [row_num2]'}, ...)`,
      description: name === 'CHOOSECOLS' ? 'Returns the specified columns from an array.' : 'Returns the specified rows from an array.',
      minArgs: 2,
      maxArgs: 255,
      args: ['value', 'value'],
      impl: (args, ctx) => {
        const a = ctx.toArray(args[0] ?? null);
        const size = name === 'CHOOSECOLS' ? a.cols : a.rows;
        const picks: number[] = [];
        for (const v of args.slice(1)) {
          for (const x of ctx.toArray(v).data) {
            if (isError(x)) return x;
            if (typeof x !== 'number') return ERRORS.VALUE;
            const k = trunc(x);
            const idx = k < 0 ? size + k : k - 1;
            if (k === 0 || idx < 0 || idx >= size) return ERRORS.VALUE;
            picks.push(idx);
          }
        }
        return name === 'CHOOSECOLS'
          ? arrayOf(a.rows, picks.length, (r, c) => a.data[r * a.cols + (picks[c] ?? 0)] ?? null)
          : arrayOf(picks.length, a.cols, (r, c) => a.data[(picks[r] ?? 0) * a.cols + c] ?? null);
      },
    }),
  ),
  ...(['VSTACK', 'HSTACK'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array1, [array2], ...)`,
      description: name === 'VSTACK' ? 'Appends arrays vertically and in sequence to return a larger array.' : 'Appends arrays horizontally and in sequence to return a larger array.',
      minArgs: 1,
      maxArgs: 254,
      args: ['value'],
      impl: (args, ctx) => {
        const arrays = args.map((v) => ctx.toArray(v));
        const pad = (a: CalcArray, r: number, c: number): CalcScalar => (r < a.rows && c < a.cols ? (a.data[r * a.cols + c] ?? null) : ERRORS.NA);
        if (name === 'VSTACK') {
          const cols = Math.max(...arrays.map((a) => a.cols));
          const rows: CalcScalar[][] = [];
          for (const a of arrays) for (let r = 0; r < a.rows; r++) rows.push(Array.from({ length: cols }, (_, c) => pad(a, r, c)));
          return fromRows(rows, cols);
        }
        const rows = Math.max(...arrays.map((a) => a.rows));
        const cols = arrays.reduce((s, a) => s + a.cols, 0);
        return arrayOf(rows, cols, (r, c) => {
          let offset = c;
          for (const a of arrays) {
            if (offset < a.cols) return pad(a, r, offset);
            offset -= a.cols;
          }
          return ERRORS.NA;
        });
      },
    }),
  ),
  ...(['TOCOL', 'TOROW'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(array, [ignore], [scan_by_column])`,
      description: name === 'TOCOL' ? 'Returns the array in a single column.' : 'Returns the array in a single row.',
      minArgs: 1,
      maxArgs: 3,
      args: ['value', 'scalar', 'scalar'],
      impl: (args, ctx) => {
        const a = ctx.toArray(args[0] ?? null);
        const ignore = args[1] === undefined || args[1] === null ? 0 : num(args[1], ctx);
        if (isError(ignore)) return ignore;
        const byCol = optBool(args[2], false);
        if (isError(byCol)) return byCol;
        const values: CalcScalar[] = [];
        const outer = byCol ? a.cols : a.rows;
        const inner = byCol ? a.rows : a.cols;
        for (let i = 0; i < outer; i++) {
          for (let j = 0; j < inner; j++) {
            const v = byCol ? (a.data[j * a.cols + i] ?? null) : (a.data[i * a.cols + j] ?? null);
            if ((ignore === 1 || ignore === 3) && v === null) continue;
            if ((ignore === 2 || ignore === 3) && isError(v)) continue;
            values.push(v);
          }
        }
        if (values.length === 0) return ERRORS.CALC;
        return name === 'TOCOL' ? makeArray(values.length, 1, values) : makeArray(1, values.length, values);
      },
    }),
  ),
  ...(['WRAPROWS', 'WRAPCOLS'] as const).map(
    (name): FunctionSpec => ({
      name,
      category: C,
      syntax: `${name}(vector, wrap_count, [pad_with])`,
      description: name === 'WRAPROWS' ? 'Wraps a row or column of values by rows after a specified number of elements.' : 'Wraps a row or column of values by columns after a specified number of elements.',
      minArgs: 2,
      maxArgs: 3,
      args: ['value', 'scalar', 'scalar'],
      impl: (args, ctx) => {
        const a = ctx.toArray(args[0] ?? null);
        if (a.rows > 1 && a.cols > 1) return ERRORS.VALUE;
        const n = num(args[1], ctx);
        if (isError(n)) return n;
        const k = trunc(n);
        if (k < 1) return ERRORS.NUM;
        const padWith = args[2] === undefined ? ERRORS.NA : scalar(args[2]);
        const outer = Math.ceil(a.data.length / k);
        const at = (i: number): CalcScalar => (i < a.data.length ? (a.data[i] ?? null) : padWith);
        return name === 'WRAPROWS' ? arrayOf(outer, k, (r, c) => at(r * k + c)) : arrayOf(k, outer, (r, c) => at(c * k + r));
      },
    }),
  ),
  {
    name: 'EXPAND',
    category: C,
    syntax: 'EXPAND(array, rows, [columns], [pad_with])',
    description: 'Expands or pads an array to specified row and column dimensions.',
    minArgs: 2,
    maxArgs: 4,
    args: ['value', 'scalar', 'scalar', 'scalar'],
    impl: (args, ctx) => {
      const a = ctx.toArray(args[0] ?? null);
      const rows = args[1] === null ? a.rows : num(args[1], ctx);
      if (isError(rows)) return rows;
      const cols = args[2] === undefined || args[2] === null ? a.cols : num(args[2], ctx);
      if (isError(cols)) return cols;
      const r = trunc(rows);
      const c = trunc(cols);
      if (r < a.rows || c < a.cols) return ERRORS.VALUE;
      const padWith = args[3] === undefined ? ERRORS.NA : scalar(args[3]);
      return arrayOf(r, c, (i, j) => (i < a.rows && j < a.cols ? (a.data[i * a.cols + j] ?? null) : padWith));
    },
  },
];
