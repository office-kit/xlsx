// Formula syntax tree. Produced by `parseFormula`, consumed by the evaluator
// and the dependency extractor. References keep the coordinates *as written*
// plus their `$` anchors; the evaluator adds the shared-formula offset to the
// relative parts, so one tree serves every cell of a shared-formula group.

import type { CalcError, CalcScalar } from './types.ts';

export type RefKind = 'cell' | 'area' | 'cols' | 'rows';

/**
 * A reference rectangle as written. For `cols` (`A:C`) the row fields are
 * 1 and MAX_ROW; for `rows` (`1:3`) the column fields are 1 and MAX_COL.
 * Corners keep their written order (`B2:A1` stays that way); evaluation
 * normalises.
 */
export interface RefArea {
  readonly kind: RefKind;
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
  readonly r1Abs: boolean;
  readonly c1Abs: boolean;
  readonly r2Abs: boolean;
  readonly c2Abs: boolean;
}

export interface SheetPrefix {
  /** First (or only) sheet. */
  readonly sheet: string;
  /** Last sheet of a 3-D reference `Sheet1:Sheet3!A1`. */
  readonly sheet2?: string;
  /** External workbook token (`1` in `[1]Sheet1!A1`). Evaluates to #REF!. */
  readonly external?: string;
}

export type StructuredSpecial = '#All' | '#Data' | '#Headers' | '#Totals' | '#This Row';

export interface StructuredSpec {
  readonly specials: readonly StructuredSpecial[];
  readonly col1?: string;
  readonly col2?: string;
}

export type BinaryOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=' | ':' | ' ' | ',';

export type AstNode =
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'string'; readonly value: string }
  | { readonly type: 'bool'; readonly value: boolean }
  | { readonly type: 'error'; readonly error: CalcError }
  | { readonly type: 'array'; readonly rows: number; readonly cols: number; readonly data: readonly CalcScalar[] }
  | { readonly type: 'ref'; readonly prefix: SheetPrefix | undefined; readonly area: RefArea }
  | { readonly type: 'name'; readonly name: string; readonly prefix: SheetPrefix | undefined }
  | { readonly type: 'structured'; readonly table: string | undefined; readonly spec: StructuredSpec }
  | { readonly type: 'unary'; readonly op: '-' | '+' | '@'; readonly operand: AstNode }
  | { readonly type: 'postfix'; readonly op: '%' | '#'; readonly operand: AstNode }
  | {
      readonly type: 'binary';
      readonly op: BinaryOp;
      readonly left: AstNode;
      readonly right: AstNode;
      /** Written inside `( )`; Excel's final-subtraction rounding skips it. */
      readonly parenthesized?: true;
    }
  /** `name` is upper-cased with any `_xlfn.` / `_xlws.` storage prefix removed. */
  | { readonly type: 'call'; readonly name: string; readonly args: readonly AstNode[] }
  /** Calling the result of an expression: `LAMBDA(x, x+1)(2)`. */
  | { readonly type: 'invoke'; readonly callee: AstNode; readonly args: readonly AstNode[] }
  /** An omitted argument: the middle of `IF(A1,,2)`. */
  | { readonly type: 'missing' };
