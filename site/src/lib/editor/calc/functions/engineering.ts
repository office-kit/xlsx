// Engineering functions: base conversion, CONVERT, bit operations.

import type { FnContext, FunctionSpec } from '../function-spec.ts';
import { type CalcError, type CalcValue, ERRORS, isError } from '../types.ts';
import { num, nums, str, trunc } from './helpers.ts';

const C = 'Engineering' as const;

// Excel's base conversions use 10-digit two's complement for negatives.
const DIGITS = 10;

const RADIX = { BIN: 2, OCT: 8, DEC: 10, HEX: 16 } as const;
type Base = keyof typeof RADIX;
const BASES: readonly Base[] = ['BIN', 'OCT', 'DEC', 'HEX'];

const limit = (radix: number): number => radix ** DIGITS / 2;

/** Decimal → base text; `places` pads positive results. */
const toBase = (n: number, radix: number, places: CalcValue | undefined, ctx: FnContext): string | CalcError => {
  const v = trunc(n);
  if (v < -limit(radix) || v >= limit(radix)) return ERRORS.NUM;
  if (v < 0) return (radix ** DIGITS + v).toString(radix).toUpperCase();
  const text = v.toString(radix).toUpperCase();
  if (places === undefined || places === null) return text;
  const p = num(places, ctx);
  if (isError(p)) return p;
  const width = trunc(p);
  if (width < text.length || width > DIGITS) return ERRORS.NUM;
  return text.padStart(width, '0');
};

/** Base text → decimal, reading 10-digit values with the top bit set as negative. */
const fromBase = (v: CalcValue | undefined, radix: number): number | CalcError => {
  const t = str(v);
  if (isError(t)) return t;
  const text = t.trim();
  if (text.length > DIGITS) return ERRORS.NUM;
  if (text === '') return 0;
  let n = 0;
  for (const ch of text) {
    const d = Number.parseInt(ch, radix);
    if (Number.isNaN(d)) return ERRORS.NUM;
    n = n * radix + d;
  }
  return text.length === DIGITS && n >= limit(radix) ? n - radix ** DIGITS : n;
};

const conversions: FunctionSpec[] = [];
for (const from of BASES) {
  for (const to of BASES) {
    if (from === to) continue;
    const name = `${from}2${to}`;
    conversions.push({
      name,
      category: C,
      syntax: to === 'DEC' ? `${name}(number)` : `${name}(number, [places])`,
      description: `Converts a ${from.toLowerCase()} number to ${to === 'DEC' ? 'decimal' : to === 'BIN' ? 'binary' : to === 'OCT' ? 'octal' : 'hexadecimal'}.`,
      minArgs: 1,
      maxArgs: to === 'DEC' ? 1 : 2,
      impl: (args, ctx) => {
        const n = from === 'DEC' ? num(args[0], ctx) : fromBase(args[0], RADIX[from]);
        if (isError(n)) return n;
        return to === 'DEC' ? n : toBase(n, RADIX[to], args[1], ctx);
      },
    });
  }
}

interface Unit {
  readonly group: string;
  /** Size in the group's base unit; temperatures are handled separately. */
  readonly factor: number;
  readonly prefixable: boolean;
}

const UNITS: ReadonlyMap<string, Unit> = new Map<string, Unit>([
  // length (m)
  ['m', { group: 'length', factor: 1, prefixable: true }],
  ['mi', { group: 'length', factor: 1609.344, prefixable: false }],
  ['Nmi', { group: 'length', factor: 1852, prefixable: false }],
  ['in', { group: 'length', factor: 0.0254, prefixable: false }],
  ['ft', { group: 'length', factor: 0.3048, prefixable: false }],
  ['yd', { group: 'length', factor: 0.9144, prefixable: false }],
  ['ang', { group: 'length', factor: 1e-10, prefixable: true }],
  // mass (g)
  ['g', { group: 'mass', factor: 1, prefixable: true }],
  ['lbm', { group: 'mass', factor: 453.59237, prefixable: false }],
  ['ozm', { group: 'mass', factor: 28.349523125, prefixable: false }],
  ['ton', { group: 'mass', factor: 907184.74, prefixable: false }],
  ['stone', { group: 'mass', factor: 6350.29318, prefixable: false }],
  // time (s)
  ['sec', { group: 'time', factor: 1, prefixable: true }],
  ['s', { group: 'time', factor: 1, prefixable: true }],
  ['mn', { group: 'time', factor: 60, prefixable: false }],
  ['min', { group: 'time', factor: 60, prefixable: false }],
  ['hr', { group: 'time', factor: 3600, prefixable: false }],
  ['day', { group: 'time', factor: 86400, prefixable: false }],
  ['d', { group: 'time', factor: 86400, prefixable: false }],
  ['yr', { group: 'time', factor: 31557600, prefixable: false }],
  // volume (l)
  ['l', { group: 'volume', factor: 1, prefixable: true }],
  ['L', { group: 'volume', factor: 1, prefixable: true }],
  ['lt', { group: 'volume', factor: 1, prefixable: true }],
  ['tsp', { group: 'volume', factor: 0.00492892159375, prefixable: false }],
  ['tbs', { group: 'volume', factor: 0.01478676478125, prefixable: false }],
  ['oz', { group: 'volume', factor: 0.0295735295625, prefixable: false }],
  ['cup', { group: 'volume', factor: 0.2365882365, prefixable: false }],
  ['pt', { group: 'volume', factor: 0.473176473, prefixable: false }],
  ['qt', { group: 'volume', factor: 0.946352946, prefixable: false }],
  ['gal', { group: 'volume', factor: 3.785411784, prefixable: false }],
  ['m3', { group: 'volume', factor: 1000, prefixable: true }],
  // area (m2)
  ['m2', { group: 'area', factor: 1, prefixable: true }],
  ['ft2', { group: 'area', factor: 0.09290304, prefixable: false }],
  ['in2', { group: 'area', factor: 0.00064516, prefixable: false }],
  ['yd2', { group: 'area', factor: 0.83612736, prefixable: false }],
  ['mi2', { group: 'area', factor: 2589988.110336, prefixable: false }],
  ['ha', { group: 'area', factor: 10000, prefixable: false }],
  ['us_acre', { group: 'area', factor: 4046.87260987425, prefixable: false }],
  ['uk_acre', { group: 'area', factor: 4046.8564224, prefixable: false }],
  // energy (J)
  ['J', { group: 'energy', factor: 1, prefixable: true }],
  ['cal', { group: 'energy', factor: 4.1868, prefixable: true }],
  ['BTU', { group: 'energy', factor: 1055.05585262, prefixable: false }],
  ['Wh', { group: 'energy', factor: 3600, prefixable: true }],
  ['eV', { group: 'energy', factor: 1.602176487e-19, prefixable: true }],
  // power (W)
  ['W', { group: 'power', factor: 1, prefixable: true }],
  ['HP', { group: 'power', factor: 745.69987158227, prefixable: false }],
  // pressure (Pa)
  ['Pa', { group: 'pressure', factor: 1, prefixable: true }],
  ['atm', { group: 'pressure', factor: 101325, prefixable: true }],
  ['mmHg', { group: 'pressure', factor: 133.322, prefixable: true }],
  ['psi', { group: 'pressure', factor: 6894.75729316836, prefixable: false }],
  // speed (m/s)
  ['m/s', { group: 'speed', factor: 1, prefixable: true }],
  ['m/h', { group: 'speed', factor: 1 / 3600, prefixable: true }],
  ['mph', { group: 'speed', factor: 0.44704, prefixable: false }],
  ['kn', { group: 'speed', factor: 1852 / 3600, prefixable: false }],
  // force (N)
  ['N', { group: 'force', factor: 1, prefixable: true }],
  ['dyn', { group: 'force', factor: 1e-5, prefixable: true }],
  ['lbf', { group: 'force', factor: 4.4482216152605, prefixable: false }],
  // temperature (handled by formula)
  ['C', { group: 'temperature', factor: 1, prefixable: false }],
  ['cel', { group: 'temperature', factor: 1, prefixable: false }],
  ['F', { group: 'temperature', factor: 1, prefixable: false }],
  ['fah', { group: 'temperature', factor: 1, prefixable: false }],
  ['K', { group: 'temperature', factor: 1, prefixable: true }],
  ['kel', { group: 'temperature', factor: 1, prefixable: true }],
]);

const PREFIXES: ReadonlyMap<string, number> = new Map([
  ['Y', 1e24],
  ['Z', 1e21],
  ['E', 1e18],
  ['P', 1e15],
  ['T', 1e12],
  ['G', 1e9],
  ['M', 1e6],
  ['k', 1e3],
  ['h', 1e2],
  ['da', 1e1],
  ['d', 1e-1],
  ['c', 1e-2],
  ['m', 1e-3],
  ['u', 1e-6],
  ['n', 1e-9],
  ['p', 1e-12],
  ['f', 1e-15],
  ['a', 1e-18],
]);

const lookupUnit = (name: string): { unit: Unit; scale: number; base: string } | undefined => {
  const direct = UNITS.get(name);
  if (direct !== undefined) return { unit: direct, scale: 1, base: name };
  for (const [prefix, scale] of PREFIXES) {
    if (!name.startsWith(prefix)) continue;
    const base = name.slice(prefix.length);
    const unit = UNITS.get(base);
    if (unit?.prefixable === true) {
      // Squared / cubed units scale by the prefix's power.
      const power = base.endsWith('2') ? 2 : base.endsWith('3') ? 3 : 1;
      return { unit, scale: scale ** power, base };
    }
  }
  return undefined;
};

const toKelvin = (v: number, unit: string): number => {
  if (unit === 'C' || unit === 'cel') return v + 273.15;
  if (unit === 'F' || unit === 'fah') return ((v - 32) * 5) / 9 + 273.15;
  return v;
};

const fromKelvin = (k: number, unit: string): number => {
  if (unit === 'C' || unit === 'cel') return k - 273.15;
  if (unit === 'F' || unit === 'fah') return ((k - 273.15) * 9) / 5 + 32;
  return k;
};

const bitArgs = (args: CalcValue[], ctx: FnContext): [number, number] | CalcError => {
  const v = nums(args, ctx, [undefined, undefined]);
  if (isError(v)) return v;
  const [a = 0, b = 0] = v;
  if (a < 0 || b < 0 || a >= 2 ** 48 || b >= 2 ** 48 || !Number.isInteger(a) || !Number.isInteger(b)) return ERRORS.NUM;
  return [a, b];
};

const bitwise = (name: string, description: string, op: (a: bigint, b: bigint) => bigint): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(number1, number2)`,
  description,
  minArgs: 2,
  maxArgs: 2,
  impl: (args, ctx) => {
    const v = bitArgs(args, ctx);
    if (isError(v)) return v;
    return Number(op(BigInt(v[0]), BigInt(v[1])));
  },
});

const shift = (name: string, left: boolean): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(number, shift_amount)`,
  description: left ? 'Returns a number shifted left by the specified number of bits.' : 'Returns a number shifted right by the specified number of bits.',
  minArgs: 2,
  maxArgs: 2,
  impl: (args, ctx) => {
    const v = nums(args, ctx, [undefined, undefined]);
    if (isError(v)) return v;
    const [n = 0, s = 0] = v;
    if (n < 0 || n >= 2 ** 48 || !Number.isInteger(n) || Math.abs(s) > 53) return ERRORS.NUM;
    const amount = trunc(s) * (left ? 1 : -1);
    const r = amount >= 0 ? n * 2 ** amount : Math.floor(n / 2 ** -amount);
    return r >= 2 ** 48 ? ERRORS.NUM : r;
  },
});

export const ENGINEERING_FUNCTIONS: FunctionSpec[] = [
  ...conversions,
  {
    name: 'CONVERT',
    category: C,
    syntax: 'CONVERT(number, from_unit, to_unit)',
    description: 'Converts a number from one measurement system to another.',
    minArgs: 3,
    maxArgs: 3,
    impl: (args, ctx) => {
      const n = num(args[0], ctx);
      if (isError(n)) return n;
      const fromName = str(args[1]);
      if (isError(fromName)) return fromName;
      const toName = str(args[2]);
      if (isError(toName)) return toName;
      const from = lookupUnit(fromName);
      const to = lookupUnit(toName);
      if (from === undefined || to === undefined || from.unit.group !== to.unit.group) return ERRORS.NA;
      if (from.unit.group === 'temperature') return fromKelvin(toKelvin(n * from.scale, from.base), to.base) / to.scale;
      return (n * from.unit.factor * from.scale) / (to.unit.factor * to.scale);
    },
  },
  {
    name: 'DELTA',
    category: C,
    syntax: 'DELTA(number1, [number2])',
    description: 'Tests whether two values are equal.',
    minArgs: 1,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, 0]);
      return isError(v) ? v : v[0] === v[1] ? 1 : 0;
    },
  },
  {
    name: 'GESTEP',
    category: C,
    syntax: 'GESTEP(number, [step])',
    description: 'Tests whether a number is greater than or equal to a threshold value.',
    minArgs: 1,
    maxArgs: 2,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, 0]);
      return isError(v) ? v : (v[0] ?? 0) >= (v[1] ?? 0) ? 1 : 0;
    },
  },
  bitwise('BITAND', "Returns a 'Bitwise And' of two numbers.", (a, b) => a & b),
  bitwise('BITOR', "Returns a 'Bitwise Or' of two numbers.", (a, b) => a | b),
  bitwise('BITXOR', "Returns a 'Bitwise Exclusive Or' of two numbers.", (a, b) => a ^ b),
  shift('BITLSHIFT', true),
  shift('BITRSHIFT', false),
];
