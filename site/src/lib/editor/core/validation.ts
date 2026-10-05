// Data validation: checking a value against a rule (on entry, and for
// Circle Invalid Data).

import type { CellValue } from '@office-kit/xlsx/cell';
import type { DataValidation } from '@office-kit/xlsx/worksheet';
import { translateFormula } from '../calc/index.ts';
import type { EditorController } from './controller.svelte.ts';
import { listSource, listValues } from './data.ts';

function scalar(v: CellValue): number | string | boolean | null {
  if (v === null) return null;
  if (v instanceof Date) return null;
  if (typeof v === 'object') {
    if (v.kind === 'formula') return v.cachedValueType === 'error' ? null : (v.cachedValue ?? null);
    if (v.kind === 'rich-text') return v.runs.map((r) => r.text).join('');
    return null;
  }
  return v;
}

/** Evaluate a rule operand (a constant or a formula relative to the rule's top-left cell). */
function operand(ctl: EditorController, dv: DataValidation, formula: string | undefined, row: number, col: number): number | undefined {
  if (formula === undefined) return undefined;
  const n = Number(formula);
  if (formula.trim() !== '' && Number.isFinite(n)) return n;
  const first = dv.sqref.ranges[0];
  const f = first ? translateFormula(formula, row - first.minRow, col - first.minCol) : formula;
  const v = ctl.doc.calc.evaluate(f, ctl.doc.ws.title, row, col);
  return typeof v === 'number' ? v : typeof v === 'string' && Number.isFinite(Number(v)) ? Number(v) : undefined;
}

function compare(op: DataValidation['operator'], x: number, a: number | undefined, b: number | undefined): boolean {
  if (a === undefined) return true;
  switch (op ?? 'between') {
    case 'between':
      return b === undefined ? x >= a : x >= Math.min(a, b) && x <= Math.max(a, b);
    case 'notBetween':
      return b === undefined ? x < a : x < Math.min(a, b) || x > Math.max(a, b);
    case 'equal':
      return x === a;
    case 'notEqual':
      return x !== a;
    case 'greaterThan':
      return x > a;
    case 'lessThan':
      return x < a;
    case 'greaterThanOrEqual':
      return x >= a;
    case 'lessThanOrEqual':
      return x <= a;
  }
}

function isError(v: CellValue): boolean {
  if (v === null || typeof v !== 'object' || v instanceof Date) return false;
  return v.kind === 'error' || (v.kind === 'formula' && v.cachedValueType === 'error');
}

export function validateValue(ctl: EditorController, dv: DataValidation, value: CellValue, row: number, col: number): boolean {
  // An error value is not a number, a date, a list item or a length; only a custom formula
  // gets to judge it.
  if (isError(value) && dv.type !== 'custom') return dv.type === undefined;
  const v = scalar(value);
  if (v === null || v === '') return dv.allowBlank !== false;
  const a = () => operand(ctl, dv, dv.formula1, row, col);
  const b = () => operand(ctl, dv, dv.formula2, row, col);
  switch (dv.type) {
    case 'whole':
      return typeof v === 'number' && Number.isInteger(v) && compare(dv.operator, v, a(), b());
    case 'decimal':
    case 'date':
    case 'time':
      return typeof v === 'number' && compare(dv.operator, v, a(), b());
    case 'textLength':
      return compare(dv.operator, String(v).length, a(), b());
    case 'list': {
      if (!dv.formula1) return false;
      // A range item matches by value (a typed 1/1/2024 is the date in the list), and also by
      // the text it shows, which is what picking it from the drop-down enters.
      if (listValues(ctl, dv.formula1).some((item) => item === v || (typeof item === 'string' && typeof v === 'string' && item.toLowerCase() === v.toLowerCase()))) return true;
      const text = typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
      return listSource(ctl, dv.formula1).some((item) => item.toLowerCase() === text.toLowerCase());
    }
    case 'custom': {
      if (!dv.formula1) return true;
      const first = dv.sqref.ranges[0];
      const f = first ? translateFormula(dv.formula1, row - first.minRow, col - first.minCol) : dv.formula1;
      const res = ctl.doc.calc.evaluate(f, ctl.doc.ws.title, row, col);
      return res === true || (typeof res === 'number' && res !== 0);
    }
    default:
      return true;
  }
}
