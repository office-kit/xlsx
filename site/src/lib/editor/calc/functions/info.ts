// Information functions.

import { cellAddress } from '../address.ts';
import type { FunctionSpec } from '../function-spec.ts';
import { type CalcScalar, type CalcValue, type ErrorCode, ERRORS, isArray, isError, isLambda, isRef } from '../types.ts';
import { num, scalar, str, trunc } from './helpers.ts';

const C = 'Information' as const;

const is = (name: string, description: string, test: (v: CalcScalar) => boolean): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(value)`,
  description,
  minArgs: 1,
  maxArgs: 1,
  acceptsErrors: true,
  impl: (args) => test(scalar(args[0])),
});

const ERROR_TYPE: Readonly<Record<ErrorCode, number>> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
  '#GETTING_DATA': 8,
  '#SPILL!': 9,
  '#CALC!': 14,
};

const parity = (name: string, even: boolean): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(number)`,
  description: even ? 'Returns TRUE if the number is even.' : 'Returns TRUE if the number is odd.',
  minArgs: 1,
  maxArgs: 1,
  impl: (args, ctx) => {
    const n = num(args[0], ctx);
    if (isError(n)) return n;
    const isEven = Math.abs(trunc(n)) % 2 === 0;
    return even ? isEven : !isEven;
  },
});

// CELL("format") codes for the common number formats.
const formatCode = (code: string): string => {
  const c = code.toLowerCase();
  if (c === 'general') return 'G';
  if (/[dmy]/.test(c.replace(/"[^"]*"/g, ''))) {
    if (/h|s/.test(c)) return 'D4';
    return c.includes('y') && c.includes('m') && c.includes('d') ? 'D1' : c.includes('y') ? 'D2' : 'D3';
  }
  const decimals = (/\.(0+)/.exec(c)?.[1] ?? '').length;
  if (c.includes('%')) return `P${decimals}`;
  if (c.includes('e+') || c.includes('e-')) return `S${decimals}`;
  if (c.includes('$')) return `C${decimals}`;
  if (c.includes(',')) return `,${decimals}`;
  return `F${decimals}`;
};

export const INFO_FUNCTIONS: FunctionSpec[] = [
  is('ISBLANK', 'Returns TRUE if the value is blank.', (v) => v === null),
  is('ISNUMBER', 'Returns TRUE if the value is a number.', (v) => typeof v === 'number'),
  is('ISTEXT', 'Returns TRUE if the value is text.', (v) => typeof v === 'string'),
  is('ISNONTEXT', 'Returns TRUE if the value is not text.', (v) => typeof v !== 'string'),
  is('ISLOGICAL', 'Returns TRUE if the value is a logical value.', (v) => typeof v === 'boolean'),
  is('ISERROR', 'Returns TRUE if the value is any error value.', (v) => isError(v)),
  is('ISERR', 'Returns TRUE if the value is any error value except #N/A.', (v) => isError(v) && v.code !== '#N/A'),
  is('ISNA', 'Returns TRUE if the value is the #N/A error value.', (v) => isError(v) && v.code === '#N/A'),
  parity('ISEVEN', true),
  parity('ISODD', false),
  {
    name: 'ISREF',
    category: C,
    syntax: 'ISREF(value)',
    description: 'Returns TRUE if the value is a reference.',
    minArgs: 1,
    maxArgs: 1,
    args: ['ref'],
    acceptsErrors: true,
    impl: (args) => args[0] !== undefined && isRef(args[0]),
  },
  {
    name: 'ISFORMULA',
    category: C,
    syntax: 'ISFORMULA(reference)',
    description: 'Checks whether there is a reference to a cell that contains a formula.',
    minArgs: 1,
    maxArgs: 1,
    args: ['ref'],
    impl: (args, ctx) => {
      const v = args[0];
      if (v === undefined || !isRef(v)) return isError(v) ? v : ERRORS.VALUE;
      const a = v.areas[0];
      return a !== undefined && ctx.host.formulaText(a.sheet, a.r1, a.c1) !== undefined;
    },
  },
  {
    name: 'NA',
    category: C,
    syntax: 'NA()',
    description: 'Returns the error value #N/A.',
    minArgs: 0,
    maxArgs: 0,
    impl: () => ERRORS.NA,
  },
  {
    name: 'ERROR.TYPE',
    category: C,
    syntax: 'ERROR.TYPE(error_val)',
    description: 'Returns a number corresponding to an error type.',
    minArgs: 1,
    maxArgs: 1,
    acceptsErrors: true,
    impl: (args) => {
      const v = scalar(args[0]);
      return isError(v) ? ERROR_TYPE[v.code] : ERRORS.NA;
    },
  },
  {
    name: 'TYPE',
    category: C,
    syntax: 'TYPE(value)',
    description: 'Returns a number indicating the data type of a value.',
    minArgs: 1,
    maxArgs: 1,
    args: ['value'],
    acceptsErrors: true,
    impl: (args) => {
      const v: CalcValue = args[0] ?? null;
      if (isArray(v)) return 64;
      if (isLambda(v)) return 128;
      if (isError(v)) return 16;
      if (typeof v === 'string') return 2;
      if (typeof v === 'boolean') return 4;
      return 1;
    },
  },
  {
    name: 'N',
    category: C,
    syntax: 'N(value)',
    description: 'Returns a value converted to a number.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args) => {
      const v = scalar(args[0]);
      if (typeof v === 'number') return v;
      if (typeof v === 'boolean') return v ? 1 : 0;
      return 0;
    },
  },
  {
    name: 'CELL',
    category: C,
    syntax: 'CELL(info_type, [reference])',
    description: 'Returns information about the formatting, location, or contents of a cell.',
    minArgs: 1,
    maxArgs: 2,
    args: ['scalar', 'ref'],
    impl: (args, ctx) => {
      const info = str(args[0]);
      if (isError(info)) return info;
      const ref = args[1];
      let sheet = ctx.sheet;
      let row = ctx.row;
      let col = ctx.col;
      if (ref !== undefined) {
        if (isError(ref)) return ref;
        if (!isRef(ref)) return ERRORS.VALUE;
        const a = ref.areas[0];
        if (a === undefined) return ERRORS.REF;
        sheet = a.sheet;
        row = a.r1;
        col = a.c1;
      }
      switch (info.toLowerCase()) {
        case 'address':
          return `$${cellAddress(row, col).replace(/(\d+)$/, '$$$1')}`;
        case 'row':
          return row;
        case 'col':
          return col;
        case 'contents':
          return ctx.host.cellValue(sheet, row, col) ?? 0;
        case 'type': {
          const v = ctx.host.cellValue(sheet, row, col);
          return v === null ? 'b' : typeof v === 'string' ? 'l' : 'v';
        }
        case 'format':
          return formatCode(ctx.host.numberFormat(sheet, row, col));
        case 'filename':
          return '';
        default:
          return ERRORS.VALUE;
      }
    },
  },
];
