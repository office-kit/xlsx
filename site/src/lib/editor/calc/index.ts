// Public surface of the formula engine used by the spreadsheet editor.

export type { AstNode } from './ast.ts';
export type { CalcEngineOptions } from './engine.ts';
export { CalcEngine } from './engine.ts';
export { toNumber as coerceToNumber, toText as coerceToText } from './coerce.ts';
export { CalcParseError } from './lexer.ts';
export { parseFormula } from './parser.ts';
export { fromStorageFormula, toStorageFormula } from './storage.ts';
export type { FormulaReferenceSpan } from './references.ts';
export { formulaReferences, toggleReferenceAt } from './references.ts';
export type { MoveEdit, StructureEdit } from './translate.ts';
export { adjustFormulaForMove, adjustFormulaForStructure, deleteSheetInFormula, renameSheetInFormula, translateFormula } from './translate.ts';
export type { FunctionInfo } from './functions/index.ts';
export { FUNCTION_CATALOG, isVolatileFunction } from './functions/index.ts';
export type { FunctionCategory } from './function-spec.ts';
export type { CalcArray, CalcError, CalcScalar, CellRef, ErrorCode } from './types.ts';
