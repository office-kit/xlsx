// Contract between the evaluator and the function library.

import type { AstNode } from './ast.ts';
import type { EvalHost } from './host.ts';
import type { CalcArray, CalcLambda, CalcRef, CalcScalar, CalcValue } from './types.ts';

export type FunctionCategory =
  | 'Math & Trig'
  | 'Statistical'
  | 'Lookup & Reference'
  | 'Text'
  | 'Logical'
  | 'Date & Time'
  | 'Financial'
  | 'Information'
  | 'Engineering'
  | 'Database';

/**
 * How the evaluator prepares one argument before the call:
 * - `scalar`: references are read; an array (or multi-cell range) makes the
 *   whole call *lift* — it runs once per element and returns an array, which
 *   is how `=LEN(A1:A3)` spills three lengths.
 * - `value`: references are read into arrays, nothing lifts (SUMPRODUCT, SORT).
 * - `ref`: passed as written — references stay references (SUM, ROW, INDEX).
 */
export type ArgKind = 'scalar' | 'value' | 'ref';

export interface FnContext {
  readonly host: EvalHost;
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
  /** Read a reference into a scalar (single cell) or array; other values pass through. */
  deref(v: CalcValue): CalcScalar | CalcArray;
  /** `deref`, then wrap a scalar as a 1×1 array. */
  toArray(v: CalcValue): CalcArray;
  callLambda(fn: CalcLambda, args: readonly CalcValue[]): CalcValue;
  /** Resolve reference text for INDIRECT, relative to the calling cell. */
  referenceFromText(text: string, a1: boolean): CalcRef | undefined;
  /** Evaluate `node` with extra LET / LAMBDA bindings layered over the caller's scope (keys upper-cased). */
  evaluateWith(node: AstNode, vars: ReadonlyMap<string, CalcValue>): CalcValue;
  /** Close over the caller's scope. */
  makeLambda(params: readonly string[], body: AstNode): CalcLambda;
}

/** Deferred argument for functions that must not evaluate every branch (IF, IFERROR, LET…). */
export type Thunk = () => CalcValue;

interface SpecBase {
  readonly name: string;
  readonly category: FunctionCategory;
  /** Signature as Excel's tooltip shows it, e.g. `SUM(number1, [number2], ...)`. */
  readonly syntax: string;
  readonly description: string;
  readonly minArgs: number;
  readonly maxArgs: number;
  /** Recalculated on every update (RAND, NOW, OFFSET, INDIRECT…). */
  readonly volatile?: boolean;
}

export interface EagerSpec extends SpecBase {
  readonly lazy?: false;
  /** Kinds of the leading arguments; the default for unlisted ones is `rest`. */
  readonly args?: readonly ArgKind[];
  /** Repeating kinds for arguments past `args` (SUMIFS pairs). Defaults to the last of `args`, else `scalar`. */
  readonly rest?: readonly ArgKind[];
  /** Hand error scalars to `impl` instead of returning the first one automatically (ISERROR, IFERROR…). */
  readonly acceptsErrors?: boolean;
  readonly impl: (args: CalcValue[], ctx: FnContext) => CalcValue;
}

export interface LazySpec extends SpecBase {
  readonly lazy: true;
  readonly impl: (args: Thunk[], ctx: FnContext, nodes: readonly AstNode[]) => CalcValue;
}

export type FunctionSpec = EagerSpec | LazySpec;
