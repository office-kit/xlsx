// COUNTIF-style criteria (`">=10"`, `"<>"`, `"a*"`, 42) and the equality
// index that keeps thousands of SUMIF / COUNTIF formulas over one column from
// rescanning it once each.

import { compareText, parseNumericText, round15 } from '../coerce.ts';
import { type CalcArray, type CalcScalar, errorFromCode, isError } from '../types.ts';

export interface Criterion {
  test(v: CalcScalar): boolean;
  /** Hash key of the single value an equality criterion matches, when it has one. */
  readonly key?: string;
}

const WILDCARD = /[*?~]/;

/** Excel wildcards: `*` any run, `?` one character, `~` escapes the next one. */
export function wildcardRegex(pattern: string): RegExp {
  let src = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i] ?? '';
    if (ch === '~' && i + 1 < pattern.length) {
      src += escapeRegex(pattern[++i] ?? '');
    } else if (ch === '*') {
      src += '.*';
    } else if (ch === '?') {
      src += '.';
    } else {
      src += escapeRegex(ch);
    }
  }
  return new RegExp(`^${src}$`, 'is');
}

const escapeRegex = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export function hasWildcard(s: string): boolean {
  return WILDCARD.test(s);
}

const numberKey = (n: number): string => `n${round15(n)}`;
const textKey = (s: string): string => `s${s.toLowerCase()}`;

/** Keys a cell value is found under in an equality index. Numeric text also answers to its number. */
const valueKeys = (v: CalcScalar, date1904: boolean): string[] => {
  if (typeof v === 'number') return [numberKey(v)];
  if (typeof v === 'boolean') return [v ? 'b1' : 'b0'];
  if (typeof v === 'string') {
    const n = parseNumericText(v, date1904);
    return n === undefined ? [textKey(v)] : [textKey(v), numberKey(n)];
  }
  if (v === null) return ['blank'];
  return [`e${v.code}`];
};

type Op = '=' | '<>' | '<' | '>' | '<=' | '>=';

const OPERATOR = /^(<=|>=|<>|=|<|>)?(.*)$/s;
const OPS: readonly Op[] = ['=', '<>', '<', '>', '<=', '>='];
const toOp = (s: string | undefined): Op | undefined => OPS.find((o) => o === s);

export function parseCriterion(c: CalcScalar, date1904: boolean): Criterion {
  if (typeof c === 'number') return numericCriterion('=', c, date1904);
  if (typeof c === 'boolean') return { key: c ? 'b1' : 'b0', test: (v) => v === c };
  if (c === null) return numericCriterion('=', 0, date1904);
  if (isError(c)) return { key: `e${c.code}`, test: (v) => isError(v) && v.code === c.code };

  const m = OPERATOR.exec(c);
  const op: Op = toOp(m?.[1]) ?? '=';
  const hasOp = m?.[1] !== undefined;
  const operand = m?.[2] ?? '';

  if (operand === '') {
    if (op === '=') {
      // `""` matches blanks and empty strings; `"="` only truly empty cells.
      return hasOp ? { key: 'blank', test: (v) => v === null } : { test: (v) => v === null || v === '' };
    }
    if (op === '<>') return { test: (v) => v !== null };
    return { test: () => false };
  }

  const n = parseNumericText(operand, date1904);
  if (n !== undefined) return numericCriterion(op, n, date1904);
  const upper = operand.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE';
    if (op === '=') return { key: b ? 'b1' : 'b0', test: (v) => v === b };
    if (op === '<>') return { test: (v) => v !== b };
    return { test: (v) => typeof v === 'boolean' && compareOrdered(op, Number(v) - Number(b)) };
  }
  const err = errorFromCode(operand);
  if (err !== undefined) {
    if (op === '=') return { key: `e${err.code}`, test: (v) => isError(v) && v.code === err.code };
    if (op === '<>') return { test: (v) => !(isError(v) && v.code === err.code) };
    return { test: () => false };
  }
  if (op === '=' || op === '<>') {
    let match: (v: CalcScalar) => boolean;
    let key: string | undefined;
    if (hasWildcard(operand)) {
      const re = wildcardRegex(operand);
      match = (v) => typeof v === 'string' && re.test(v);
    } else {
      const lower = operand.toLowerCase();
      key = textKey(operand);
      match = (v) => typeof v === 'string' && v.toLowerCase() === lower;
    }
    if (op === '=') return key !== undefined ? { key, test: match } : { test: match };
    return { test: (v) => !match(v) };
  }
  return { test: (v) => typeof v === 'string' && compareOrdered(op, compareText(v, operand)) };
}

const compareOrdered = (op: Op, c: number): boolean => {
  switch (op) {
    case '<':
      return c < 0;
    case '>':
      return c > 0;
    case '<=':
      return c <= 0;
    case '>=':
      return c >= 0;
    case '=':
      return c === 0;
    case '<>':
      return c !== 0;
  }
};

function numericCriterion(op: Op, n: number, date1904: boolean): Criterion {
  const target = round15(n);
  const asNumber = (v: CalcScalar): number | undefined => {
    if (typeof v === 'number') return round15(v);
    if (typeof v === 'string' && op === '=') {
      const parsed = parseNumericText(v, date1904);
      return parsed === undefined ? undefined : round15(parsed);
    }
    return undefined;
  };
  if (op === '=') return { key: numberKey(n), test: (v) => asNumber(v) === target };
  if (op === '<>') return { test: (v) => typeof v !== 'number' || round15(v) !== target };
  return {
    test: (v) => typeof v === 'number' && compareOrdered(op, round15(v) - target),
  };
}

// Index per materialised range array. The engine hands out one array object
// per range per recalculation pass, so the cache lives exactly as long as the
// values it indexes.
const indexes = new WeakMap<CalcArray, Map<string, number[]>>();
const INDEX_THRESHOLD = 64;

/** Positions in `range` matching `criterion`, in order. */
export function matchingPositions(range: CalcArray, criterion: Criterion, date1904: boolean): number[] {
  if (criterion.key !== undefined && range.data.length >= INDEX_THRESHOLD) {
    let index = indexes.get(range);
    if (index === undefined) {
      index = new Map();
      for (let i = 0; i < range.data.length; i++) {
        for (const key of valueKeys(range.data[i] ?? null, date1904)) {
          const list = index.get(key);
          if (list === undefined) index.set(key, [i]);
          else list.push(i);
        }
      }
      indexes.set(range, index);
    }
    return index.get(criterion.key) ?? [];
  }
  const out: number[] = [];
  for (let i = 0; i < range.data.length; i++) if (criterion.test(range.data[i] ?? null)) out.push(i);
  return out;
}
