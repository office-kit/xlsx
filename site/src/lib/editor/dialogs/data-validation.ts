// Data Validation dialog model: converting a stored rule to the dialog's
// fields and back, and replacing the rules on the selected cells.

import type { DataValidation, DataValidationErrorStyle, DataValidationOperator, DataValidationType } from '@office-kit/xlsx/worksheet';
import { makeDataValidation } from '@office-kit/xlsx/worksheet';
import { CalcParseError, fromStorageFormula, parseFormula, toStorageFormula } from '../calc/index.ts';
import type { Range } from '../core/address.ts';
import { fromBoundaries, rangesIntersect, toBoundaries } from '../core/address.ts';
import type { EditorController } from '../core/controller.svelte.ts';
import { validationAt } from '../core/data.ts';
import { parseInput, serialToDate } from '../core/input.ts';
import type { MessageKey } from '../i18n/i18n.svelte.ts';

export type Allow = 'any' | DataValidationType;

export interface ValidationForm {
  allow: Allow;
  operator: DataValidationOperator;
  /** Minimum / value / source / formula, as typed (formulas keep their '='). */
  value1: string;
  value2: string;
  ignoreBlank: boolean;
  inCellDropdown: boolean;
  applyToAll: boolean;
  showInput: boolean;
  inputTitle: string;
  inputMessage: string;
  showError: boolean;
  errorStyle: DataValidationErrorStyle;
  errorTitle: string;
  errorMessage: string;
}

export function blankForm(): ValidationForm {
  return {
    allow: 'any',
    operator: 'between',
    value1: '',
    value2: '',
    ignoreBlank: true,
    inCellDropdown: true,
    applyToAll: false,
    showInput: true,
    inputTitle: '',
    inputMessage: '',
    showError: true,
    errorStyle: 'stop',
    errorTitle: '',
    errorMessage: '',
  };
}

export function usesOperator(allow: Allow): boolean {
  return allow === 'whole' || allow === 'decimal' || allow === 'date' || allow === 'time' || allow === 'textLength';
}

export function usesSecondValue(form: ValidationForm): boolean {
  return usesOperator(form.allow) && (form.operator === 'between' || form.operator === 'notBetween');
}

function pad(n: number): string {
  return String(n).padStart(2, '0');
}

function timeText(fraction: number): string {
  const total = Math.round((fraction - Math.floor(fraction)) * 86_400);
  return `${Math.floor(total / 3600)}:${pad(Math.floor((total % 3600) / 60))}:${pad(total % 60)}`;
}

function dateText(serial: number, ymd: boolean, date1904: boolean): string {
  const d = serialToDate(serial, date1904);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + 1;
  const day = d.getUTCDate();
  return ymd ? `${y}/${m}/${day}` : `${m}/${day}/${y}`;
}

/** A stored operand as the dialog shows it. */
function operandText(type: DataValidationType, f: string | undefined, ymd: boolean, date1904: boolean): string {
  if (f === undefined) return '';
  const n = Number(f);
  const numeric = f.trim() !== '' && Number.isFinite(n);
  if (type === 'list') return f.startsWith('"') && f.endsWith('"') ? f.slice(1, -1) : `=${fromStorageFormula(f)}`;
  if (type === 'custom' || !numeric) return `=${fromStorageFormula(f)}`;
  if (type === 'date') return dateText(n, ymd, date1904);
  if (type === 'time') return timeText(n);
  return f;
}

export function formFor(ctl: EditorController, dv: DataValidation | undefined): ValidationForm {
  const form = blankForm();
  if (!dv) return form;
  const ymd = ctl.dateOrder() === 'ymd';
  const date1904 = ctl.doc.wb.date1904;
  return {
    ...form,
    // A typeless rule (one that only carries an input message) reads back as 'custom' without a formula.
    allow: dv.type === 'custom' && dv.formula1 === undefined ? 'any' : dv.type,
    operator: dv.operator ?? 'between',
    value1: operandText(dv.type, dv.formula1, ymd, date1904),
    value2: operandText(dv.type, dv.formula2, ymd, date1904),
    ignoreBlank: dv.allowBlank !== false,
    // The file stores the inverse: showDropDown="1" hides the arrow.
    inCellDropdown: dv.showDropDown !== true,
    showInput: dv.showInputMessage !== false,
    inputTitle: dv.promptTitle ?? '',
    inputMessage: dv.prompt ?? '',
    showError: dv.showErrorMessage !== false,
    errorStyle: dv.errorStyle ?? 'stop',
    errorTitle: dv.errorTitle ?? '',
    errorMessage: dv.error ?? '',
  };
}

/** The rule at the active cell, if any. */
export function activeRule(ctl: EditorController): DataValidation | undefined {
  const { row, col } = ctl.doc.selection.active;
  return validationAt(ctl.doc.ws, row, col);
}

function isFormula(text: string): boolean {
  try {
    parseFormula(text);
    return true;
  } catch (e) {
    if (e instanceof CalcParseError) return false;
    throw e;
  }
}

type Operand = { ok: true; formula: string | undefined } | { ok: false; error: MessageKey };

function operand(ctl: EditorController, allow: DataValidationType, text: string): Operand {
  const s = text.trim();
  if (s === '') return { ok: false, error: allow === 'list' ? 'dlgDvListRequired' : 'dlgDvValueRequired' };
  if (s.startsWith('=')) return isFormula(s.slice(1)) ? { ok: true, formula: toStorageFormula(s.slice(1)) } : { ok: false, error: 'formulaError' };
  switch (allow) {
    case 'list':
      return { ok: true, formula: `"${s}"` };
    case 'custom':
      return isFormula(s) ? { ok: true, formula: toStorageFormula(s) } : { ok: false, error: 'formulaError' };
    case 'whole':
    case 'textLength': {
      const n = Number(s);
      if (!Number.isInteger(n) || (allow === 'textLength' && n < 0)) return { ok: false, error: 'dlgDvNotWhole' };
      return { ok: true, formula: String(n) };
    }
    case 'decimal': {
      const n = Number(s);
      return Number.isFinite(n) ? { ok: true, formula: String(n) } : { ok: false, error: 'dlgDvNotNumber' };
    }
    case 'date':
    case 'time': {
      const { value } = parseInput(s, { dateOrder: ctl.dateOrder(), date1904: ctl.doc.wb.date1904 });
      if (typeof value !== 'number') return { ok: false, error: allow === 'date' ? 'dlgDvNotDate' : 'dlgDvNotTime' };
      return { ok: true, formula: String(value) };
    }
  }
}

/** Rectangle `a` minus rectangle `b`, as up to four rectangles. */
function subtract(a: Range, b: Range): Range[] {
  if (!rangesIntersect(a, b)) return [a];
  const out: Range[] = [];
  if (a.r1 < b.r1) out.push({ ...a, r2: b.r1 - 1 });
  if (a.r2 > b.r2) out.push({ ...a, r1: b.r2 + 1 });
  const r1 = Math.max(a.r1, b.r1);
  const r2 = Math.min(a.r2, b.r2);
  if (a.c1 < b.c1) out.push({ r1, r2, c1: a.c1, c2: b.c1 - 1 });
  if (a.c2 > b.c2) out.push({ r1, r2, c1: b.c2 + 1, c2: a.c2 });
  return out;
}

/**
 * Validate the form and write it: the selection's old rules are cut out and
 * the new rule covers the selection. "Any value" writes no rule unless it
 * carries an input message. Returns the message to show when the input is
 * invalid.
 */
export function applyValidation(ctl: EditorController, form: ValidationForm, existing: DataValidation | undefined): MessageKey | undefined {
  let formula1: string | undefined;
  let formula2: string | undefined;
  if (form.allow !== 'any') {
    const a = operand(ctl, form.allow, form.value1);
    if (!a.ok) return a.error;
    formula1 = a.formula;
    if (usesSecondValue(form)) {
      const b = operand(ctl, form.allow, form.value2);
      if (!b.ok) return b.error;
      formula2 = b.formula;
      const lo = Number(formula1);
      const hi = Number(formula2);
      if (Number.isFinite(lo) && Number.isFinite(hi) && lo > hi) return 'dlgDvMinMax';
    }
    if (form.allow === 'list' && formula1 !== undefined && formula1.length > 257) return 'dlgDvListTooLong';
  }
  const doc = ctl.doc;
  const ws = doc.ws;
  const targets = doc.selection.ranges;
  const promptOnly = form.allow === 'any' && form.showInput && (form.inputTitle !== '' || form.inputMessage !== '');
  const rule =
    form.allow === 'any' && !promptOnly
      ? undefined
      : makeDataValidation({
          // The library has no "any" type; a 'custom' rule without a formula accepts every value.
          type: form.allow === 'any' ? 'custom' : form.allow,
          sqref: { ranges: targets.map(toBoundaries) },
          ...(usesOperator(form.allow) ? { operator: form.operator } : {}),
          ...(formula1 !== undefined ? { formula1 } : {}),
          ...(formula2 !== undefined ? { formula2 } : {}),
          allowBlank: form.ignoreBlank,
          ...(form.allow === 'list' && !form.inCellDropdown ? { showDropDown: true } : {}),
          showInputMessage: form.showInput,
          showErrorMessage: form.showError,
          ...(form.inputTitle ? { promptTitle: form.inputTitle } : {}),
          ...(form.inputMessage ? { prompt: form.inputMessage } : {}),
          errorStyle: form.errorStyle,
          ...(form.errorTitle ? { errorTitle: form.errorTitle } : {}),
          ...(form.errorMessage ? { error: form.errorMessage } : {}),
        });
  doc.transact('Data Validation', (tx) => {
    tx.sheet(ws, 'dataValidations');
    const next: DataValidation[] = [];
    for (const dv of ws.dataValidations) {
      // "Apply these changes to all other cells with the same settings" rewrites the shared rule.
      if (form.applyToAll && dv === existing) {
        if (rule) next.push({ ...rule, sqref: { ranges: [...dv.sqref.ranges, ...rule.sqref.ranges] } });
        continue;
      }
      let pieces = dv.sqref.ranges.map(fromBoundaries);
      for (const target of targets) pieces = pieces.flatMap((p) => subtract(p, target));
      if (pieces.length > 0) next.push({ ...dv, sqref: { ranges: pieces.map(toBoundaries) } });
    }
    if (rule && !(form.applyToAll && existing)) next.push(rule);
    ws.dataValidations = next;
  });
  return undefined;
}
