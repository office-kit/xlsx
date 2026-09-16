// Data validations.
//
// A DataValidation entry attaches a constraint (one of seven type kinds,
// optional operator, two formula slots) to a sqref-style MultiCellRange.
// Stage-1 maps every OOXML attribute we have a use for; imeMode + numeric/value
// clamps land later when phase 7's Asian-locale support catches up.

import { normalizeFormulaText } from '../utils/formula-text.js';
import { type MultiCellRange, parseMultiCellRange } from './cell-range.js';

export type DataValidationType = 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';
export type DataValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'greaterThanOrEqual'
  | 'lessThan'
  | 'lessThanOrEqual';
export type DataValidationErrorStyle = 'stop' | 'warning' | 'information';

export interface DataValidation {
  /** Constraint kind. `'list'` is the dropdown form (formula1 = comma list or range). */
  type: DataValidationType;
  /** Comparison operator — only meaningful for whole/decimal/date/time/textLength. */
  operator?: DataValidationOperator;
  /** Lower-bound formula or list source. Always required for list/whole/decimal/date/time. */
  formula1?: string;
  /** Upper-bound formula. Used by between / notBetween only. */
  formula2?: string;
  /** Allow empty cells. */
  allowBlank?: boolean;
  /** Show the input message popover when the cell is selected. */
  showInputMessage?: boolean;
  /** Show the error message popover on invalid entry. */
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
  errorStyle?: DataValidationErrorStyle;
  promptTitle?: string;
  prompt?: string;
  /** Excel inverts this attribute: `showDropDown="1"` actually *hides* the dropdown arrow on list-type validators. We mirror the wire form. */
  showDropDown?: boolean;
  /** Apply-to range (sqref). */
  sqref: MultiCellRange;
}

export function makeDataValidation(
  opts: Omit<Partial<DataValidation>, 'sqref'> & {
    type: DataValidationType;
    sqref: MultiCellRange | string;
  },
): DataValidation {
  return {
    type: opts.type,
    sqref: typeof opts.sqref === 'string' ? parseMultiCellRange(opts.sqref) : opts.sqref,
    ...(opts.operator !== undefined ? { operator: opts.operator } : {}),
    ...(opts.formula1 !== undefined ? { formula1: normalizeFormulaText(opts.formula1) } : {}),
    ...(opts.formula2 !== undefined ? { formula2: normalizeFormulaText(opts.formula2) } : {}),
    ...(opts.allowBlank !== undefined ? { allowBlank: opts.allowBlank } : {}),
    ...(opts.showInputMessage !== undefined ? { showInputMessage: opts.showInputMessage } : {}),
    ...(opts.showErrorMessage !== undefined ? { showErrorMessage: opts.showErrorMessage } : {}),
    ...(opts.errorTitle !== undefined ? { errorTitle: opts.errorTitle } : {}),
    ...(opts.error !== undefined ? { error: opts.error } : {}),
    ...(opts.errorStyle !== undefined ? { errorStyle: opts.errorStyle } : {}),
    ...(opts.promptTitle !== undefined ? { promptTitle: opts.promptTitle } : {}),
    ...(opts.prompt !== undefined ? { prompt: opts.prompt } : {}),
    ...(opts.showDropDown !== undefined ? { showDropDown: opts.showDropDown } : {}),
  };
}

// ---- Worksheet ergonomic builders ---------------------------------------

import type { Worksheet } from './worksheet.js';

const resolveSqref = (sqref: MultiCellRange | string): MultiCellRange =>
  typeof sqref === 'string' ? parseMultiCellRange(sqref) : sqref;

export interface ValidationCommon {
  /** Show the dropdown / input prompt when the cell is selected. */
  prompt?: string;
  promptTitle?: string;
  /** Show an error dialog when the user types an invalid value. */
  error?: string;
  errorTitle?: string;
  errorStyle?: DataValidationErrorStyle;
  allowBlank?: boolean;
}

/**
 * Add a list-type dropdown validation to a range. `values` may be an inline
 * list (`['Red', 'Green', 'Blue']`) or a sheet reference
 * (`'Sheet1!$A$1:$A$10'`, with or without a leading `=`).
 */
export const addListValidation = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  values: ReadonlyArray<string> | string,
  opts: ValidationCommon = {},
): DataValidation => {
  const formula1 = Array.isArray(values)
    ? `"${(values as ReadonlyArray<string>).join(',')}"`
    : (values as string);
  const dv = makeDataValidation({
    type: 'list',
    sqref: resolveSqref(sqref),
    formula1,
    allowBlank: opts.allowBlank ?? true,
    showInputMessage: opts.prompt !== undefined,
    showErrorMessage: opts.error !== undefined || opts.errorStyle !== undefined,
    ...(opts.errorStyle !== undefined ? { errorStyle: opts.errorStyle } : {}),
    ...(opts.error !== undefined ? { error: opts.error } : {}),
    ...(opts.errorTitle !== undefined ? { errorTitle: opts.errorTitle } : {}),
    ...(opts.prompt !== undefined ? { prompt: opts.prompt } : {}),
    ...(opts.promptTitle !== undefined ? { promptTitle: opts.promptTitle } : {}),
  });
  ws.dataValidations.push(dv);
  return dv;
};

/**
 * Add a number-range validation. `between(min, max)` matches Excel's "Whole
 * Number" → "between" form by default. Use `kind: 'decimal'` for decimal
 * (default 'whole').
 */
export const addNumberValidation = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  range: { min: number; max?: number; operator?: DataValidationOperator; kind?: 'whole' | 'decimal' },
  opts: ValidationCommon = {},
): DataValidation => {
  const operator: DataValidationOperator =
    range.operator ?? (range.max !== undefined ? 'between' : 'greaterThanOrEqual');
  const dv = makeDataValidation({
    type: range.kind ?? 'whole',
    sqref: resolveSqref(sqref),
    operator,
    formula1: String(range.min),
    ...(range.max !== undefined ? { formula2: String(range.max) } : {}),
    allowBlank: opts.allowBlank ?? true,
    showInputMessage: opts.prompt !== undefined,
    showErrorMessage: opts.error !== undefined || opts.errorStyle !== undefined,
    ...(opts.errorStyle !== undefined ? { errorStyle: opts.errorStyle } : {}),
    ...(opts.error !== undefined ? { error: opts.error } : {}),
    ...(opts.errorTitle !== undefined ? { errorTitle: opts.errorTitle } : {}),
    ...(opts.prompt !== undefined ? { prompt: opts.prompt } : {}),
    ...(opts.promptTitle !== undefined ? { promptTitle: opts.promptTitle } : {}),
  });
  ws.dataValidations.push(dv);
  return dv;
};

/**
 * Add a date-range validation. Dates are passed as Excel serial numbers (use
 * `dateToExcel` to convert from JS `Date`).
 */
export const addDateValidation = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  range: { min: number; max?: number; operator?: DataValidationOperator },
  opts: ValidationCommon = {},
): DataValidation => {
  const operator: DataValidationOperator =
    range.operator ?? (range.max !== undefined ? 'between' : 'greaterThanOrEqual');
  const dv = makeDataValidation({
    type: 'date',
    sqref: resolveSqref(sqref),
    operator,
    formula1: String(range.min),
    ...(range.max !== undefined ? { formula2: String(range.max) } : {}),
    allowBlank: opts.allowBlank ?? true,
    showInputMessage: opts.prompt !== undefined,
    showErrorMessage: opts.error !== undefined || opts.errorStyle !== undefined,
    ...(opts.errorStyle !== undefined ? { errorStyle: opts.errorStyle } : {}),
    ...(opts.error !== undefined ? { error: opts.error } : {}),
    ...(opts.errorTitle !== undefined ? { errorTitle: opts.errorTitle } : {}),
    ...(opts.prompt !== undefined ? { prompt: opts.prompt } : {}),
    ...(opts.promptTitle !== undefined ? { promptTitle: opts.promptTitle } : {}),
  });
  ws.dataValidations.push(dv);
  return dv;
};

/** Add a custom-formula validation (`formula1` evaluated for each cell). */
export const addCustomValidation = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  formula: string,
  opts: ValidationCommon = {},
): DataValidation => {
  const dv = makeDataValidation({
    type: 'custom',
    sqref: resolveSqref(sqref),
    formula1: formula,
    allowBlank: opts.allowBlank ?? true,
    showInputMessage: opts.prompt !== undefined,
    showErrorMessage: opts.error !== undefined || opts.errorStyle !== undefined,
    ...(opts.errorStyle !== undefined ? { errorStyle: opts.errorStyle } : {}),
    ...(opts.error !== undefined ? { error: opts.error } : {}),
    ...(opts.errorTitle !== undefined ? { errorTitle: opts.errorTitle } : {}),
    ...(opts.prompt !== undefined ? { prompt: opts.prompt } : {}),
    ...(opts.promptTitle !== undefined ? { promptTitle: opts.promptTitle } : {}),
  });
  ws.dataValidations.push(dv);
  return dv;
};
