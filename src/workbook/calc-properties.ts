// Workbook-level <calcPr>. Per ECMA-376 §18.2.2.
//
// `calcId` is a build identifier Excel uses to decide whether to force
// a recalc when re-opening; modern Excel treats any value as "use
// what's there". The other attrs cover Excel's calculation options
// (Tools → Options → Formulas).

export type CalcMode = 'manual' | 'auto' | 'autoNoTable';
export type RefMode = 'A1' | 'R1C1';

export interface CalcProperties {
  /** Excel build/calc-engine identifier; OpenOffice sets 191621, Excel 2016+ 162913 etc. */
  calcId?: number;
  calcMode?: CalcMode;
  /** Force a full recalc when the workbook is loaded. Excel default `true` for safety. */
  fullCalcOnLoad?: boolean;
  refMode?: RefMode;
  /** Allow circular references via iterative calculation. */
  iterate?: boolean;
  iterateCount?: number;
  iterateDelta?: number;
  /** Use full 15-digit precision (vs. displayed precision). */
  fullPrecision?: boolean;
  calcCompleted?: boolean;
  /** Run a recalc on save. */
  calcOnSave?: boolean;
  /** Use multi-threaded calculation. */
  concurrentCalc?: boolean;
  concurrentManualCount?: number;
  /** Force a full recalc on next interaction. */
  forceFullCalc?: boolean;
}

export const makeCalcProperties = (opts: CalcProperties = {}): CalcProperties => ({ ...opts });

// ---- Workbook ergonomic helpers ----------------------------------------

import type { Workbook } from './workbook.js';

const ensureCalcProperties = (wb: Workbook): CalcProperties => {
  if (!wb.calcProperties) wb.calcProperties = {};
  return wb.calcProperties;
};

/**
 * Set the recalculation mode. `'auto'` is Excel's default;
 * `'manual'` requires F9 to recompute formulas; `'autoNoTable'`
 * recomputes everything except data-table cells.
 */
export const setCalcMode = (wb: Workbook, mode: CalcMode): void => {
  ensureCalcProperties(wb).calcMode = mode;
};

/**
 * Toggle iterative calculation (Excel's "Enable iterative calculation"
 * option). When `enable` is true and `count` / `delta` are provided,
 * they replace the default Excel limits (100 iterations, 0.001 delta).
 */
export const setIterativeCalc = (
  wb: Workbook,
  enable: boolean,
  opts: { count?: number; delta?: number } = {},
): void => {
  const calc = ensureCalcProperties(wb);
  calc.iterate = enable;
  if (opts.count !== undefined) calc.iterateCount = opts.count;
  if (opts.delta !== undefined) calc.iterateDelta = opts.delta;
};

/** Toggle "Recalculate workbook before saving" (workbook-level). */
export const setCalcOnSave = (wb: Workbook, on: boolean): void => {
  ensureCalcProperties(wb).calcOnSave = on;
};

/**
 * Toggle "Recalculate workbook on load". A calculating app (Excel,
 * LibreOffice, Google Sheets) then recomputes every formula on open instead of
 * trusting the cached values in the file, which is what a generated workbook
 * wants when the producer could not compute a cached value for every formula,
 * or is not certain the ones it wrote are current. Viewers that never
 * calculate ignore the flag and keep showing only what is cached.
 */
export const setFullCalcOnLoad = (wb: Workbook, on: boolean): void => {
  ensureCalcProperties(wb).fullCalcOnLoad = on;
};

/** Toggle "Set precision as displayed" (false = full 15-digit precision). */
export const setFullPrecision = (wb: Workbook, on: boolean): void => {
  ensureCalcProperties(wb).fullPrecision = on;
};
