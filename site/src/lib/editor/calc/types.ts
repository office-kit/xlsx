// Value model shared by the parser, the evaluator and the function library.
//
// A formula evaluates to a `CalcValue`: a scalar, a 2-D array, a reference
// (kept unresolved so ROW / OFFSET / INDEX can see coordinates and SUM can walk
// only populated cells) or a LAMBDA closure. Only scalars ever reach the
// workbook model.

export type ErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#SPILL!'
  | '#CALC!'
  | '#GETTING_DATA';

export interface CalcError {
  readonly kind: 'error';
  readonly code: ErrorCode;
}

/** `null` is a blank cell (or an omitted argument), never a value a formula can produce at top level. */
export type CalcScalar = number | string | boolean | CalcError | null;

export interface CellRef {
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
}

/** Row-major 2-D array. `rows * cols === data.length`, never empty. */
export interface CalcArray {
  readonly kind: 'array';
  readonly rows: number;
  readonly cols: number;
  readonly data: CalcScalar[];
}

/** One rectangle of a reference, 1-based inclusive, on the sheet named `sheet`. */
export interface Area {
  readonly sheet: string;
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

/** A reference. More than one area only for unions `(A1,B2)` and 3-D refs. */
export interface CalcRef {
  readonly kind: 'ref';
  readonly areas: readonly Area[];
}

export interface CalcLambda {
  readonly kind: 'lambda';
  readonly params: readonly string[];
  readonly body: import('./ast.ts').AstNode;
  readonly scope: Scope | undefined;
}

/** LET / LAMBDA bindings, chained to the enclosing scope. Keys are upper-cased. */
export interface Scope {
  readonly vars: Map<string, CalcValue>;
  readonly parent: Scope | undefined;
}

export type CalcValue = CalcScalar | CalcArray | CalcRef | CalcLambda;

export const MAX_ROW = 1_048_576;
export const MAX_COL = 16_384;

const makeError = (code: ErrorCode): CalcError => Object.freeze({ kind: 'error', code });

// Singletons so hot paths compare and allocate nothing.
export const ERRORS = {
  NULL: makeError('#NULL!'),
  DIV0: makeError('#DIV/0!'),
  VALUE: makeError('#VALUE!'),
  REF: makeError('#REF!'),
  NAME: makeError('#NAME?'),
  NUM: makeError('#NUM!'),
  NA: makeError('#N/A'),
  SPILL: makeError('#SPILL!'),
  CALC: makeError('#CALC!'),
  GETTING_DATA: makeError('#GETTING_DATA'),
} as const;

const ERROR_BY_CODE: ReadonlyMap<string, CalcError> = new Map(
  Object.values(ERRORS).map((e) => [e.code, e] as const),
);

/** The error singleton for `code` (case-insensitive), or undefined when it is not an error token. */
export function errorFromCode(code: string): CalcError | undefined {
  return ERROR_BY_CODE.get(code.toUpperCase());
}

export const ERROR_CODE_LIST: readonly ErrorCode[] = Object.values(ERRORS).map((e) => e.code);

// Accepts anything so functions can check intermediate results (`number[] | CalcError`) too.
export function isError(v: unknown): v is CalcError {
  return typeof v === 'object' && v !== null && 'kind' in v && v.kind === 'error';
}

export function isArray(v: CalcValue): v is CalcArray {
  return typeof v === 'object' && v !== null && v.kind === 'array';
}

export function isRef(v: CalcValue): v is CalcRef {
  return typeof v === 'object' && v !== null && v.kind === 'ref';
}

export function isLambda(v: CalcValue): v is CalcLambda {
  return typeof v === 'object' && v !== null && v.kind === 'lambda';
}

export function makeArray(rows: number, cols: number, data: CalcScalar[]): CalcArray {
  return { kind: 'array', rows, cols, data };
}

export function makeRef(area: Area): CalcRef {
  return { kind: 'ref', areas: [area] };
}

export function areaRows(a: Area): number {
  return a.r2 - a.r1 + 1;
}

export function areaCols(a: Area): number {
  return a.c2 - a.c1 + 1;
}

export function arrayAt(a: CalcArray, r: number, c: number): CalcScalar {
  return a.data[r * a.cols + c] ?? null;
}
