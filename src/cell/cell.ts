// Cell value model. Mirrors openpyxl/openpyxl/cell/cell.py.
//
// A Cell is a plain mutable object: the worksheet stores millions of these, so
// per-cell freezes / spreads aren't viable on the hot path. The public surface
// stays small — makeCell + getCoordinate + targeted setters — and uses
// discriminated unions for the special CellValue shapes (formula, rich text,
// duration, error).

import { columnLetterFromIndex, MAX_COL, MAX_ROW } from '../utils/coordinate.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { normalizeFormulaText, spellsFormula } from '../utils/formula-text.js';
import { ERROR_CODES } from '../utils/inference.js';
import { type RichText, richTextToString } from './rich-text.js';

/**
 * An Excel error token: the `#` sigil then the code, as the file stores it.
 * `ERROR_CODES` lists the ones this library knows by name and is what
 * {@link makeErrorValue} accepts; the type stays open because Excel has added
 * nine tokens since 2018 and a reader that refused an unlisted one would fail
 * a whole workbook over a single cell.
 */
export type ExcelErrorCode = `#${string}`;

/** Formula sub-kind — drives the OOXML `<f t="…">` attribute. */
export type FormulaKind = 'normal' | 'array' | 'shared' | 'dataTable';

export interface FormulaValue {
  readonly kind: 'formula';
  readonly formula: string;
  readonly t: FormulaKind;
  /** Cached value Excel last computed for the cell, used when `data_only` reads it back. */
  readonly cachedValue?: number | string | boolean;
  /** Mark an Excel error token as an error result. Without this, strings stay text, including "#N/A". */
  readonly cachedValueType?: 'error';
  /** Range string (`"A1:A10"`) for array / shared / dataTable formulas. */
  readonly ref?: string;
  /** Shared-formula index. */
  readonly si?: number;
  /** Data-table specific fields (mirrors openpyxl DataTableFormula). */
  readonly r1?: string;
  readonly r2?: string;
  readonly dt2D?: boolean;
  readonly dtr?: boolean;
  readonly del1?: boolean;
  readonly del2?: boolean;
  readonly aca?: boolean;
  readonly ca?: boolean;
}

export type CellValue =
  | number
  | string
  | boolean
  | Date
  | { kind: 'duration'; ms: number }
  | { kind: 'error'; code: ExcelErrorCode }
  | { kind: 'rich-text'; runs: RichText }
  | FormulaValue
  | null;

export interface Cell {
  /** 1-based row index. */
  row: number;
  /** 1-based column index. */
  col: number;
  /** Effective cell value. `null` represents an empty cell. */
  value: CellValue;
  /** Index into Workbook.styles.cellXfs. 0 = default. */
  styleId: number;
  /** Optional reference to a Hyperlink registered on the worksheet. */
  hyperlinkId?: number;
  /** Optional reference to a Comment registered on the worksheet. */
  commentId?: number;
}

/** Marker subtype for the placeholder cells inside a merged range (top-left holds the value). */
export interface MergedCell extends Cell {
  merged: true;
}

const validateCoord = (row: number, col: number): void => {
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW) {
    throw new OpenXmlSchemaError(`Cell row ${row} out of range [1, ${MAX_ROW}]`);
  }
  if (!Number.isInteger(col) || col < 1 || col > MAX_COL) {
    throw new OpenXmlSchemaError(`Cell col ${col} out of range [1, ${MAX_COL}]`);
  }
};

/** Build a fresh Cell. Validates coordinates against the OOXML grid bounds. */
export function makeCell(row: number, col: number, value: CellValue = null, styleId = 0): Cell {
  validateCoord(row, col);
  return { row, col, value, styleId };
}

/** Format a Cell's coordinate as the canonical "A1" string. */
export function getCoordinate(c: Cell): string {
  return `${columnLetterFromIndex(c.col)}${c.row}`;
}

/**
 * Direct value setter. No type inference, no validation beyond the union — the
 * caller is in charge. Use {@link bindValue} for the "do what I mean" path.
 */
export function setCellValue(c: Cell, value: CellValue): void {
  c.value = value;
}

/**
 * "Smart" setter: infers the cell value from a JS runtime value.
 * - `string` starting with `=` → formula, except a lone `'='`, which is text
 * - `string` matching an Excel error token → error variant
 * - other primitives / Date / null pass through verbatim
 *
 * This follows what Excel does with the same text typed into a cell. A string
 * sent down the formula path has to be a formula: `'==A1'` and `'= '` throw
 * here the way {@link setFormula} throws for them, as Excel refuses them too.
 * Data that may legitimately start with `=` (a CSV column holding `'==>'`)
 * belongs in {@link setCellValue}, which infers nothing.
 *
 * Intentionally not the default — explicit is clearer for typed code, and
 * inferring on every write costs measurable time on the worksheet write hot
 * path.
 */
export function bindValue(c: Cell, value: number | string | boolean | Date | null): void {
  if (typeof value === 'string') {
    if (spellsFormula(value)) {
      setFormula(c, value);
      return;
    }
    if (ERROR_CODES.has(value)) {
      c.value = { kind: 'error', code: value as ExcelErrorCode };
      return;
    }
    c.value = value;
    return;
  }
  c.value = value;
}

// ---- formula values --------------------------------------------------------

// A `<f>` with no text is the shape Excel reports as damaged, and `'='` on its
// own normalises to exactly that, so the kinds whose text is required reject it
// here rather than at save time.
const requireFormulaText = (fn: string, formula: string): string => {
  const text = normalizeFormulaText(formula, fn);
  if (text.length === 0) {
    throw new OpenXmlSchemaError(`${fn}: formula text must not be empty (got ${JSON.stringify(formula)})`);
  }
  return text;
};

/**
 * Build a formula cell value. Hand it to `setCell` when the write already
 * carries a style id, so placing a formatted formula stays a single call;
 * {@link setFormula} is the same thing applied to a cell you already hold.
 *
 * A leading `=` is stripped, so `'=SUM(A1:A3)'` and `'SUM(A1:A3)'` are
 * interchangeable. Exactly one comes off: text that still starts with `=`
 * after that (`'==A1'`) is rejected with an {@link OpenXmlSchemaError}, since
 * it is not a formula Excel accepts either and `A1` is not what it meant.
 *
 * A cached value is optional. Without one, Excel, LibreOffice and Google
 * Sheets compute the result on open, but viewers that never calculate (Quick
 * Look, Outlook and SharePoint previews, most thumbnailers) render the cell
 * empty. Supply one whenever the producer can compute it. For an Excel error,
 * supply `{ cachedValue: '#N/A', cachedValueType: 'error' }`; without the type,
 * the same token is a string result.
 */
export function makeFormula(
  formula: string,
  opts?: Pick<FormulaValue, 'cachedValue' | 'cachedValueType'>,
): FormulaValue {
  return Object.freeze({
    kind: 'formula',
    t: 'normal',
    formula: requireFormulaText('makeFormula', formula),
    ...(opts?.cachedValue !== undefined ? { cachedValue: opts.cachedValue } : {}),
    ...(opts?.cachedValueType !== undefined ? { cachedValueType: opts.cachedValueType } : {}),
  });
}

/**
 * Build an array (CSE) formula value spanning `ref`. Belongs on the top-left
 * cell of the range, since Excel reads `ref` to know how far the result
 * spreads.
 */
export function makeArrayFormula(
  ref: string,
  formula: string,
  opts?: Pick<FormulaValue, 'cachedValue' | 'cachedValueType'>,
): FormulaValue {
  return Object.freeze({
    kind: 'formula',
    t: 'array',
    formula: requireFormulaText('makeArrayFormula', formula),
    ref,
    ...(opts?.cachedValue !== undefined ? { cachedValue: opts.cachedValue } : {}),
    ...(opts?.cachedValueType !== undefined ? { cachedValueType: opts.cachedValueType } : {}),
  });
}

/**
 * Build a shared-formula value. The first cell in the group carries the
 * formula text + `ref`; subsequent cells with the same `si` carry only the
 * index, and Excel reconstructs their text by shifting the references, so
 * their `formula` is legitimately empty.
 */
export function makeSharedFormula(
  si: number,
  formula?: string,
  ref?: string,
  opts?: Pick<FormulaValue, 'cachedValue' | 'cachedValueType'>,
): FormulaValue {
  if (!Number.isInteger(si) || si < 0) {
    throw new OpenXmlSchemaError(`makeSharedFormula: si must be a non-negative integer; got ${si}`);
  }
  return Object.freeze({
    kind: 'formula',
    t: 'shared',
    formula: formula === undefined ? '' : normalizeFormulaText(formula, 'makeSharedFormula'),
    si,
    ...(ref !== undefined ? { ref } : {}),
    ...(opts?.cachedValue !== undefined ? { cachedValue: opts.cachedValue } : {}),
    ...(opts?.cachedValueType !== undefined ? { cachedValueType: opts.cachedValueType } : {}),
  });
}

/**
 * Plain `A1+B1` style formula, applied in place. Same value as
 * {@link makeFormula}, for a cell you already hold.
 */
export function setFormula(c: Cell, formula: string, opts?: Pick<FormulaValue, 'cachedValue' | 'cachedValueType'>): void {
  c.value = makeFormula(formula, opts);
}

/** Array (CSE) formula spanning a `ref` range, applied in place. */
export function setArrayFormula(
  c: Cell,
  ref: string,
  formula: string,
  opts?: Pick<FormulaValue, 'cachedValue' | 'cachedValueType'>,
): void {
  c.value = makeArrayFormula(ref, formula, opts);
}

/** Shared formula, applied in place. See {@link makeSharedFormula}. */
export function setSharedFormula(
  c: Cell,
  si: number,
  formula?: string,
  ref?: string,
  opts?: Pick<FormulaValue, 'cachedValue' | 'cachedValueType'>,
): void {
  c.value = makeSharedFormula(si, formula, ref, opts);
}

/**
 * Excel data-table formula (`<f t="dataTable">`). These appear as the "What-if
 * Analysis → Data Table" feature output: a 1- or 2-variable sensitivity grid
 * where the formula references one or two input cells. The wire format mirrors
 * openpyxl's `DataTableFormula`:
 *
 * - `ref`     — inclusive cell range the formula spans.
 * - `r1`, `r2`— input cell coordinates ("$A$1" etc.).
 * - `dt2D`    — true for two-variable tables (uses both r1 and r2).
 * - `dtr`     — row-direction flag (true) vs column-direction (false).
 * - `del1`/`del2` — Excel marks one of these true when the
 * corresponding input cell has been deleted; the formula keeps round-tripping
 * so Excel can show the warning state.
 * - `aca`/`ca` — alwaysCalculate / calculate flags.
 */
export interface DataTableFormulaOpts {
  ref: string;
  r1?: string;
  r2?: string;
  dt2D?: boolean;
  dtr?: boolean;
  del1?: boolean;
  del2?: boolean;
  aca?: boolean;
  ca?: boolean;
  cachedValue?: FormulaValue['cachedValue'];
  cachedValueType?: FormulaValue['cachedValueType'];
}

/**
 * Build a data-table formula value. Preserves all the dt-specific attributes
 * so the writer can re-emit `<f t="dataTable" r1="..." />` verbatim and Excel
 * keeps treating the cell as a Data Table cell. Excel writes these with no
 * formula text at all, so `formula` may be empty.
 */
export function makeDataTableFormula(formula: string, opts: DataTableFormulaOpts): FormulaValue {
  return Object.freeze({
    kind: 'formula',
    t: 'dataTable',
    formula: normalizeFormulaText(formula, 'makeDataTableFormula'),
    ref: opts.ref,
    ...(opts.r1 !== undefined ? { r1: opts.r1 } : {}),
    ...(opts.r2 !== undefined ? { r2: opts.r2 } : {}),
    ...(opts.dt2D !== undefined ? { dt2D: opts.dt2D } : {}),
    ...(opts.dtr !== undefined ? { dtr: opts.dtr } : {}),
    ...(opts.del1 !== undefined ? { del1: opts.del1 } : {}),
    ...(opts.del2 !== undefined ? { del2: opts.del2 } : {}),
    ...(opts.aca !== undefined ? { aca: opts.aca } : {}),
    ...(opts.ca !== undefined ? { ca: opts.ca } : {}),
    ...(opts.cachedValue !== undefined ? { cachedValue: opts.cachedValue } : {}),
    ...(opts.cachedValueType !== undefined ? { cachedValueType: opts.cachedValueType } : {}),
  });
}

/** Data-table formula, applied in place. See {@link makeDataTableFormula}. */
export function setDataTableFormula(c: Cell, formula: string, opts: DataTableFormulaOpts): void {
  c.value = makeDataTableFormula(formula, opts);
}

// ---- value-shape helpers ---------------------------------------------------

/** Build a `{ kind: 'error', code }` cell value. */
export function makeErrorValue(code: ExcelErrorCode): { kind: 'error'; code: ExcelErrorCode } {
  if (!ERROR_CODES.has(code)) {
    throw new OpenXmlSchemaError(`makeErrorValue: unknown error code "${code}"`);
  }
  return Object.freeze({ kind: 'error', code });
}

/** Build a `{ kind: 'duration', ms }` cell value. */
export function makeDurationValue(ms: number): { kind: 'duration'; ms: number } {
  if (!Number.isFinite(ms)) {
    throw new OpenXmlSchemaError(`makeDurationValue: ms "${ms}" is not finite`);
  }
  return Object.freeze({ kind: 'duration', ms });
}

/** True iff `c.value` is the formula variant. */
export function isFormulaCell(c: Cell): boolean {
  return isFormulaValue(c.value);
}

/** True iff `c.value` is the rich-text variant. */
export function isRichTextCell(c: Cell): boolean {
  return isRichTextValue(c.value);
}

/** Returns true iff the cell has no content. */
export function isEmptyCell(c: Cell): boolean {
  return c.value === null;
}

/**
 * Type guard for `MergedCell` — true iff the cell is a placeholder for a
 * merged-range covered cell (the top-left of a merged range holds the value;
 * the rest are `MergedCell`). Use this to filter merge-placeholders out of
 * value-walking loops.
 */
export function isMergedCell(c: Cell): c is MergedCell {
  return (c as MergedCell).merged === true;
}

/** Returns true iff the cell holds an Excel error value (`#REF!`, `#NAME?`, …). */
export function isErrorCell(c: Cell): boolean {
  return isErrorValue(c.value);
}

/**
 * Get the formula text from a formula-bearing cell, or `undefined` for
 * non-formula cells. Equivalent to: isFormulaValue(c.value) ? c.value.formula :
 * undefined but spares callers the type-narrow + member access.
 */
export function getFormulaText(c: Cell): string | undefined {
  return isFormulaValue(c.value) ? c.value.formula : undefined;
}

/**
 * Get the cached value Excel last computed for a formula cell, or `undefined`
 * for non-formula / uncached cells. Useful for `data_only` read paths that want
 * the displayed result without re-evaluating.
 */
export function getCachedFormulaValue(c: Cell): number | string | boolean | undefined {
  return isFormulaValue(c.value) ? c.value.cachedValue : undefined;
}

// ---- value-level type guards + coercion ----------------------------------

/** True iff `v` is the formula variant. */
export function isFormulaValue(v: CellValue): v is FormulaValue {
  return typeof v === 'object' && v !== null && (v as { kind?: string }).kind === 'formula';
}

/** True iff `v` is the rich-text variant. */
export function isRichTextValue(v: CellValue): v is { kind: 'rich-text'; runs: RichText } {
  return typeof v === 'object' && v !== null && (v as { kind?: string }).kind === 'rich-text';
}

/** True iff `v` is the error variant. */
export function isErrorValue(v: CellValue): v is { kind: 'error'; code: ExcelErrorCode } {
  return typeof v === 'object' && v !== null && (v as { kind?: string }).kind === 'error';
}

/** True iff `v` is the duration variant. */
export function isDurationValue(v: CellValue): v is { kind: 'duration'; ms: number } {
  return typeof v === 'object' && v !== null && (v as { kind?: string }).kind === 'duration';
}

export interface CellValueAsStringOptions {
  /** Renderer for `Date` cells. Defaults to `d => d.toISOString()`. */
  dateFormat?: (value: Date) => string;
  /** Replacement for the `null` cell value. Defaults to `''`. */
  emptyText?: string;
}

/**
 * Coerce a CellValue to its plain-string display form. Numbers / booleans
 * convert via `String`; rich text concatenates run text; formulas yield the
 * cached value (or empty string when uncached); errors yield their Excel token;
 * durations yield `"<ms> ms"` with no formatting; Dates yield
 * `Date.toISOString()`; `null` yields `""`.
 *
 * Pass `opts.dateFormat` to override the Date renderer (e.g. a locale-specific
 * format) and `opts.emptyText` to substitute a different placeholder for
 * `null` cells.
 */
export function cellValueAsString(v: CellValue, opts?: CellValueAsStringOptions): string {
  const emptyText = opts?.emptyText ?? '';
  if (v === null) return emptyText;
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return opts?.dateFormat ? opts.dateFormat(v) : v.toISOString();
  if (isRichTextValue(v)) return richTextToString(v.runs);
  if (isFormulaValue(v)) {
    if (v.cachedValue === undefined) return '';
    return typeof v.cachedValue === 'string' ? v.cachedValue : String(v.cachedValue);
  }
  if (isErrorValue(v)) return v.code;
  if (isDurationValue(v)) return `${v.ms} ms`;
  return '';
}

/**
 * Coerce a CellValue to `boolean | undefined`. Booleans pass through; `'TRUE'`
 * / `'true'` and `'FALSE'` / `'false'` (case-insensitive) parse to true /
 * false; numbers yield `false` for 0 and `true` for any other finite value
 * (matching Excel's truthy-number coercion); formula cells return their cached
 * boolean if any. Everything else (null, Date, error, duration, rich-text,
 * non-bool strings) yields `undefined`.
 */
export function cellValueAsBoolean(v: CellValue): boolean | undefined {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return undefined;
    return v !== 0;
  }
  if (typeof v === 'string') {
    const lc = v.toLowerCase();
    if (lc === 'true') return true;
    if (lc === 'false') return false;
    return undefined;
  }
  if (isFormulaValue(v) && typeof v.cachedValue === 'boolean') return v.cachedValue;
  return undefined;
}

/**
 * Coerce a CellValue to a `Date` when one is meaningful. Pass-through for
 * `Date`-typed values; ISO-8601 strings (anything `new Date(s)` parses to a
 * finite time) round-trip; durations are interpreted as `new Date(ms)`.
 * Numbers, booleans, formulas, errors, rich text, and null all return
 * `undefined` — this helper does **not** apply the Excel-serial-to-Date
 * conversion (use `excelToDate` for that).
 */
export function cellValueAsDate(v: CellValue): Date | undefined {
  if (v instanceof Date) return v;
  if (typeof v === 'string') {
    if (v === '') return undefined;
    const t = Date.parse(v);
    if (!Number.isFinite(t)) return undefined;
    return new Date(t);
  }
  if (isDurationValue(v)) return new Date(v.ms);
  return undefined;
}

/**
 * Coerce a CellValue to a number when one is meaningful. Booleans yield 0/1;
 * numeric strings parse via `Number(s)`; rich-text concats then parses;
 * formulas with a numeric cached value pass through. Returns `undefined` when
 * there's no sensible numeric reading (text strings, errors, dates, durations,
 * null, empty).
 */
export function cellValueAsNumber(v: CellValue): number | undefined {
  if (v === null) return undefined;
  if (typeof v === 'number') return Number.isFinite(v) ? v : undefined;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    if (v === '') return undefined;
    const n = Number(v);
    return Number.isFinite(n) ? n : undefined;
  }
  if (isFormulaValue(v) && typeof v.cachedValue === 'number') return v.cachedValue;
  return undefined;
}

/**
 * Map a CellValue to the most natural JS primitive for display / export.
 * Unlike `cellValueAsString`/`cellValueAsNumber`/etc., which each force a
 * single target type and return `undefined` when the value can't be coerced,
 * this returns whatever primitive best represents the union variant:
 *
 * - `null` → `null`
 * - `string` / `number` / `boolean` / `Date` → passthrough
 * - rich text → joined run text (`string`)
 * - formula → recursive on `cachedValue` (`null` when uncached)
 * - error → error code (`string`)
 * - duration → `ms` (`number`)
 */
export function cellValueAsPrimitive(v: CellValue): string | number | boolean | Date | null {
  if (v === null) return null;
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
  if (v instanceof Date) return v;
  if (isRichTextValue(v)) return richTextToString(v.runs);
  if (isFormulaValue(v)) {
    return v.cachedValue === undefined ? null : v.cachedValue;
  }
  if (isErrorValue(v)) return v.code;
  if (isDurationValue(v)) return v.ms;
  return null;
}
