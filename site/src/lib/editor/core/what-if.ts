// Data ▸ What-If Analysis: Goal Seek, Data Table and the Scenario Manager.
//
// All three ask "what would these formulas show if those inputs were…": the
// inputs are set on the model for a moment, the engine recalculates their
// dependents, the results are read, and everything is put back — outside the
// undo history — before the command commits its outcome as one step.

import type { CellValue } from '@office-kit/xlsx/cell';
import { makeDataTableFormula } from '@office-kit/xlsx/cell';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import { addWorksheet } from '@office-kit/xlsx/workbook';
import type { Scenario, ScenarioList } from '@office-kit/xlsx/worksheet';
import type { CalcScalar } from '../calc/index.ts';
import type { MessageKey } from '../i18n/i18n.svelte.ts';
import { cellAddress, parseRangeAddress, type CellPos, type Range } from './address.ts';
import { getCellAt, setValueAt } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';

function isFormula(v: CellValue | undefined): boolean {
  return v !== null && v !== undefined && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula';
}

/**
 * Run `read` with each cell in `inputs` holding its value, then restore the
 * cells and their dependents exactly as they were.
 */
export function withInputs<T>(ctl: EditorController, inputs: ReadonlyArray<{ readonly pos: CellPos; readonly value: CellValue }>, read: () => T): T {
  const doc = ctl.doc;
  const ws = doc.ws;
  const sheet = ws.title;
  const saved = inputs.map(({ pos }) => ({ pos, cell: getCellAt(ws, pos.row, pos.col), value: getCellAt(ws, pos.row, pos.col)?.value ?? null }));
  const refs = inputs.map(({ pos }) => ({ sheet, row: pos.row, col: pos.col }));
  try {
    for (const { pos, value } of inputs) setValueAt(ws, pos.row, pos.col, value);
    doc.calc.update(refs);
    return read();
  } finally {
    for (const s of saved) {
      if (s.cell) s.cell.value = s.value;
      else {
        ws.rows.get(s.pos.row)?.delete(s.pos.col);
        if (ws.rows.get(s.pos.row)?.size === 0) ws.rows.delete(s.pos.row);
      }
    }
    doc.calc.update(refs);
  }
}

// ---- Goal Seek ----------------------------------------------------------------

export interface GoalSeekResult {
  readonly found: boolean;
  readonly value: number;
  readonly result: CalcScalar;
}

/** Excel's defaults: 100 iterations, stop within 0.001 of the target. */
const MAX_ITERATIONS = 100;
const TOLERANCE = 0.001;

export function goalSeekError(ctl: EditorController, setCell: CellPos, changing: CellPos): MessageKey | undefined {
  const ws = ctl.doc.ws;
  if (!isFormula(getCellAt(ws, setCell.row, setCell.col)?.value)) return 'dtGsNeedFormula';
  const v = getCellAt(ws, changing.row, changing.col)?.value;
  if (isFormula(v) || (v !== undefined && v !== null && typeof v !== 'number')) return 'dtGsNeedValue';
  return undefined;
}

/**
 * Find the value of `changing` that makes `setCell` show `target`: secant
 * steps from the current value, falling back to bisection once the target is
 * bracketed. The model is left unchanged.
 */
export function goalSeek(ctl: EditorController, setCell: CellPos, target: number, changing: CellPos): GoalSeekResult {
  const sheet = ctl.doc.ws.title;
  const evalAt = (x: number): CalcScalar => withInputs(ctl, [{ pos: changing, value: x }], () => ctl.doc.calc.cellValue(sheet, setCell.row, setCell.col));
  const f = (x: number): number | undefined => {
    const r = evalAt(x);
    return typeof r === 'number' ? r - target : undefined;
  };
  const start = getCellAt(ctl.doc.ws, changing.row, changing.col)?.value;
  let a = typeof start === 'number' ? start : 0;
  let fa = f(a);
  let b = a === 0 ? 0.01 : a * 1.01;
  let fb = f(b);
  let best = { x: a, err: fa === undefined ? Infinity : Math.abs(fa) };
  /** Points on either side of the target once one is found; bisection stays inside them. */
  let bracket: { lo: number; flo: number; hi: number } | undefined;
  for (let i = 0; i < MAX_ITERATIONS && fa !== undefined && fb !== undefined; i++) {
    if (Math.abs(fb) < best.err) best = { x: b, err: Math.abs(fb) };
    if (best.err <= Number.EPSILON * Math.max(1, Math.abs(target))) break;
    if (bracket) {
      if (Math.sign(fb) === Math.sign(bracket.flo)) {
        bracket.lo = b;
        bracket.flo = fb;
      } else bracket.hi = b;
    } else if (Math.sign(fa) !== Math.sign(fb)) bracket = { lo: a, flo: fa, hi: b };
    let next = fb === fa ? b + (b - a || 1) : b - (fb * (b - a)) / (fb - fa);
    if (bracket && !(next > Math.min(bracket.lo, bracket.hi) && next < Math.max(bracket.lo, bracket.hi))) next = (bracket.lo + bracket.hi) / 2;
    if (!Number.isFinite(next) || next === b) break;
    a = b;
    fa = fb;
    b = next;
    fb = f(b);
  }
  if (fb !== undefined && Math.abs(fb) < best.err) best = { x: b, err: Math.abs(fb) };
  return { found: best.err <= TOLERANCE, value: best.x, result: evalAt(best.x) };
}

/** Commit Goal Seek's answer as one undo step. */
export function setGoalSeekValue(ctl: EditorController, changing: CellPos, value: number): void {
  const ws = ctl.doc.ws;
  ctl.doc.transact('Goal Seek', (tx) => {
    tx.cells(ws, { r1: changing.row, c1: changing.col, r2: changing.row, c2: changing.col });
    setValueAt(ws, changing.row, changing.col, value, ctl.defaultStyleAt(changing.row, changing.col));
  });
}

// ---- Data Table ---------------------------------------------------------------

function toModel(v: CalcScalar): CellValue {
  if (v === null || typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  return { kind: 'error', code: v.code };
}

function cached(v: CalcScalar): { cachedValue?: number | string | boolean; cachedValueType?: 'error' } {
  if (v === null) return {};
  if (typeof v === 'object') return { cachedValue: v.code, cachedValueType: 'error' };
  return { cachedValue: v };
}

/**
 * Fill a one- or two-variable data table over `range` (Excel's layout: input
 * values along the first row and/or column, formulas in the other edge or at
 * the corner). The results are computed now and stored with Excel's
 * `dataTable` formula on the first result cell, so Excel recomputes them.
 */
export function fillDataTable(ctl: EditorController, range: Range, rowInput: CellPos | undefined, colInput: CellPos | undefined): MessageKey | undefined {
  const doc = ctl.doc;
  const ws = doc.ws;
  const sheet = ws.title;
  if (!rowInput && !colInput) return 'dtDtNeedInput';
  if (range.r2 <= range.r1 || range.c2 <= range.c1) return 'dtDtTooSmall';
  const inside = (p: CellPos | undefined) => p !== undefined && p.row >= range.r1 && p.row <= range.r2 && p.col >= range.c1 && p.col <= range.c2;
  if (inside(rowInput) || inside(colInput)) return 'dtDtInputInside';
  const val = (row: number, col: number): CellValue => toModel(doc.calc.cellValue(sheet, row, col));
  const results: CalcScalar[][] = [];
  for (let r = range.r1 + 1; r <= range.r2; r++) {
    const line: CalcScalar[] = [];
    for (let c = range.c1 + 1; c <= range.c2; c++) {
      const inputs: Array<{ pos: CellPos; value: CellValue }> = [];
      let formula: CellPos;
      if (rowInput && colInput) {
        inputs.push({ pos: rowInput, value: val(range.r1, c) }, { pos: colInput, value: val(r, range.c1) });
        formula = { row: range.r1, col: range.c1 };
      } else if (colInput) {
        inputs.push({ pos: colInput, value: val(r, range.c1) });
        formula = { row: range.r1, col: c };
      } else if (rowInput) {
        inputs.push({ pos: rowInput, value: val(range.r1, c) });
        formula = { row: r, col: range.c1 };
      } else return 'dtDtNeedInput';
      line.push(withInputs(ctl, inputs, () => doc.calc.cellValue(sheet, formula.row, formula.col)));
    }
    results.push(line);
  }
  const body: Range = { r1: range.r1 + 1, c1: range.c1 + 1, r2: range.r2, c2: range.c2 };
  const first = rowInput ?? colInput;
  if (!first) return 'dtDtNeedInput';
  doc.transact('Data Table', (tx) => {
    tx.cells(ws, body);
    results.forEach((line, i) =>
      line.forEach((v, j) => {
        const row = body.r1 + i;
        const col = body.c1 + j;
        const value: CellValue =
          i === 0 && j === 0
            ? makeDataTableFormula('', {
                ref: `${cellAddress(body.r1, body.c1)}:${cellAddress(body.r2, body.c2)}`,
                r1: cellAddress(first.row, first.col),
                ...(rowInput && colInput ? { r2: cellAddress(colInput.row, colInput.col), dt2D: true, dtr: true } : { dtr: rowInput !== undefined }),
                ...cached(v),
              })
            : toModel(v);
        setValueAt(ws, row, col, value, ctl.defaultStyleAt(row, col));
      }),
    );
  });
  return undefined;
}

// ---- Scenarios -------------------------------------------------------------------

export interface ScenarioDraft {
  readonly name: string;
  /** Changing cells as the user typed them ("B1,B2" or "B1:B3"). */
  readonly cells: string;
  readonly comment: string;
  readonly locked: boolean;
  readonly hidden: boolean;
}

/** Excel limits a scenario to 32 changing cells. */
const MAX_CHANGING = 32;

/** Single-cell addresses of a changing-cells text, or an error. */
export function changingCells(text: string): string[] | MessageKey {
  const out: string[] = [];
  for (const part of text.split(/[,;]/).map((s) => s.trim()).filter((s) => s !== '')) {
    const parsed = parseRangeAddress(part.replace(/^=/, ''));
    if (!parsed || parsed.sheet !== undefined) return 'invalidReference';
    const r = parsed.range;
    for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) out.push(cellAddress(row, col));
  }
  if (out.length === 0) return 'invalidReference';
  if (out.length > MAX_CHANGING) return 'dtScTooMany';
  return out;
}

function posOf(ref: string): CellPos | undefined {
  const r = parseRangeAddress(ref)?.range;
  return r ? { row: r.r1, col: r.c1 } : undefined;
}

/** A scenario value as Excel stores it: text, typed back to a number when it reads as one. */
function parseScenarioValue(val: string): CellValue {
  const n = Number(val);
  return val.trim() !== '' && Number.isFinite(n) ? n : val;
}

export function currentValues(ctl: EditorController, refs: readonly string[]): string[] {
  return refs.map((ref) => {
    const p = posOf(ref);
    const cell = p ? getCellAt(ctl.doc.ws, p.row, p.col) : undefined;
    if (!cell || cell.value === null) return '';
    return typeof cell.value === 'number' ? String(cell.value) : getCellDisplayText(ctl.doc.wb, cell);
  });
}

function scenarioList(ctl: EditorController): ScenarioList {
  return ctl.doc.ws.scenarios ?? { scenarios: [] };
}

export function scenarioNameError(ctl: EditorController, name: string, except?: number): MessageKey | undefined {
  if (name.trim() === '') return 'dtScNameEmpty';
  const lower = name.trim().toLowerCase();
  if (scenarioList(ctl).scenarios.some((s, i) => i !== except && s.name.toLowerCase() === lower)) return 'dtScNameTaken';
  return undefined;
}

/** Add (index undefined) or replace a scenario. */
export function saveScenario(ctl: EditorController, draft: ScenarioDraft, refs: readonly string[], values: readonly string[], index?: number): void {
  const ws = ctl.doc.ws;
  const scenario: Scenario = {
    name: draft.name.trim(),
    inputCells: refs.map((ref, i) => ({ ref, val: values[i] ?? '' })),
    ...(draft.comment ? { comment: draft.comment } : {}),
    ...(draft.locked ? { locked: true } : {}),
    ...(draft.hidden ? { hidden: true } : {}),
  };
  ctl.doc.transact('Scenario', (tx) => {
    tx.sheet(ws, 'scenarios');
    const list = scenarioList(ctl);
    const scenarios = list.scenarios.slice();
    if (index === undefined) scenarios.push(scenario);
    else scenarios[index] = scenario;
    ws.scenarios = { ...list, scenarios };
  });
}

export function deleteScenario(ctl: EditorController, index: number): void {
  const ws = ctl.doc.ws;
  ctl.doc.transact('Delete Scenario', (tx) => {
    tx.sheet(ws, 'scenarios');
    const list = scenarioList(ctl);
    const scenarios = list.scenarios.filter((_, i) => i !== index);
    if (scenarios.length === 0) delete ws.scenarios;
    else {
      const { current: _c, show: _s, ...rest } = list;
      ws.scenarios = { ...rest, scenarios };
    }
  });
}

/** Show: put the scenario's values into its changing cells. */
export function showScenario(ctl: EditorController, index: number): void {
  const ws = ctl.doc.ws;
  const scenario = scenarioList(ctl).scenarios[index];
  if (!scenario) return;
  const cells = scenario.inputCells.flatMap((c) => {
    const pos = posOf(c.ref);
    return pos ? [{ pos, value: parseScenarioValue(c.val) }] : [];
  });
  ctl.doc.transact('Show Scenario', (tx) => {
    tx.sheet(ws, 'scenarios');
    for (const { pos } of cells) tx.cells(ws, { r1: pos.row, c1: pos.col, r2: pos.row, c2: pos.col });
    for (const { pos, value } of cells) setValueAt(ws, pos.row, pos.col, value, ctl.defaultStyleAt(pos.row, pos.col));
    ws.scenarios = { ...scenarioList(ctl), current: index, show: index };
  });
}

/**
 * Summary: a new sheet tabulating each scenario's changing cells and the
 * values `resultRefs` take under it, next to the current values — Excel's
 * Scenario Summary report.
 */
export function scenarioSummary(ctl: EditorController, resultRefs: readonly string[], labels: { title: string; current: string; changing: string; result: string; sheetName: string }): void {
  const doc = ctl.doc;
  const ws = doc.ws;
  const sheet = ws.title;
  const list = scenarioList(ctl).scenarios;
  const changing = [...new Set(list.flatMap((s) => s.inputCells.map((c) => c.ref)))].flatMap((ref) => {
    const pos = posOf(ref);
    return pos ? [{ ref, pos }] : [];
  });
  const results = resultRefs.flatMap((ref) => {
    const pos = posOf(ref);
    return pos ? [{ ref, pos }] : [];
  });
  const current = (pos: CellPos) => doc.calc.cellValue(sheet, pos.row, pos.col);
  const columns = [
    { name: labels.current, changing: changing.map((c) => current(c.pos)), results: results.map((r) => current(r.pos)) },
    ...list.map((s) => {
      const inputs = s.inputCells.flatMap((c) => {
        const pos = posOf(c.ref);
        return pos ? [{ pos, value: parseScenarioValue(c.val) }] : [];
      });
      const byRef = new Map(s.inputCells.map((c) => [c.ref, parseScenarioValue(c.val)]));
      return {
        name: s.name,
        changing: changing.map((c): CalcScalar => {
          const v = byRef.get(c.ref);
          return v === undefined ? current(c.pos) : typeof v === 'number' || typeof v === 'string' ? v : null;
        }),
        results: withInputs(ctl, inputs, () => results.map((r) => current(r.pos))),
      };
    }),
  ];
  const names = new Set(doc.wb.sheets.map((s) => s.sheet.title.toLowerCase()));
  let title = labels.sheetName;
  for (let n = 2; names.has(title.toLowerCase()); n++) title = `${labels.sheetName} ${n}`;
  const index = doc.activeSheetIndex;
  doc.transact('Scenario Summary', (tx) => {
    tx.structural = true;
    tx.workbook('sheets', 'activeSheetIndex', 'definedNames');
    const out = addWorksheet(doc.wb, title, { index });
    setValueAt(out, 2, 2, labels.title);
    columns.forEach((c, j) => setValueAt(out, 3, 4 + j, c.name));
    setValueAt(out, 4, 2, labels.changing);
    changing.forEach((c, i) => {
      setValueAt(out, 5 + i, 3, c.ref);
      columns.forEach((col, j) => setValueAt(out, 5 + i, 4 + j, toModel(col.changing[i] ?? null)));
    });
    const resultTop = 5 + changing.length;
    if (results.length > 0) setValueAt(out, resultTop, 2, labels.result);
    results.forEach((r, i) => {
      setValueAt(out, resultTop + 1 + i, 3, r.ref);
      columns.forEach((c, j) => setValueAt(out, resultTop + 1 + i, 4 + j, toModel(c.results[i] ?? null)));
    });
  });
  doc.activateSheet(index);
}
