// Data ▸ Subtotal: insert a SUBTOTAL row after (or before) each run of equal
// values in a column, a grand total, and the row outline that collapses the
// detail — and Remove All, which takes them out again.
//
// Rows are inserted one by one so references into the data follow their
// rows, as with Excel's own insertions; the sheet is declared on the undo
// step once for the whole run.

import { getCellDisplayText } from '@office-kit/xlsx/styles';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { colLetter, MAX_COL, type Range } from './address.ts';
import { getCellAt, isBlank, setValueAt } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import { transformStyle } from './format.ts';
import { MAX_OUTLINE_LEVEL } from './outline.ts';
import { applyStructuralEdit, declareStructural } from './structure.ts';

export type SubtotalFunction = 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'stdDevp' | 'var' | 'varp';

/** SUBTOTAL function numbers. */
export const SUBTOTAL_CODES: Readonly<Record<SubtotalFunction, number>> = {
  sum: 9,
  count: 3,
  average: 1,
  max: 4,
  min: 5,
  product: 6,
  countNums: 2,
  stdDev: 7,
  stdDevp: 8,
  var: 10,
  varp: 11,
};

/** The Use function list, in Excel's order. */
export const SUBTOTAL_FUNCTIONS: readonly SubtotalFunction[] = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'];

export interface SubtotalOptions {
  /** Sheet column whose changes start a new group. */
  readonly groupBy: number;
  readonly fn: SubtotalFunction;
  /** Sheet columns that get a SUBTOTAL formula. */
  readonly columns: readonly number[];
  readonly replace: boolean;
  readonly pageBreaks: boolean;
  readonly summaryBelow: boolean;
}

export interface SubtotalLabels {
  /** Label of a group's row from the group's value ("East Total"). */
  readonly group: (value: string) => string;
  /** Label of the grand total row ("Grand Total"). */
  readonly grand: string;
}

const SUBTOTAL_RE = /^\s*SUBTOTAL\s*\(/i;

/** Whether the row holds a SUBTOTAL formula inside the range's columns: a row this command made. */
export function isSubtotalRow(ws: Worksheet, row: number, range: Range): boolean {
  const rowMap = ws.rows.get(row);
  if (!rowMap) return false;
  for (const [c, cell] of rowMap) {
    if (c < range.c1 || c > range.c2) continue;
    const v = cell.value;
    if (v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' && SUBTOTAL_RE.test(v.formula)) return true;
  }
  return false;
}

/** Runs of equal values (case-insensitive, as Excel compares) in `groupBy`, between subtotal rows. */
export function groupRuns(ws: Worksheet, keyAt: (row: number) => string, range: Range): Array<{ start: number; end: number; key: string }> {
  const runs: Array<{ start: number; end: number; key: string }> = [];
  let cur: { start: number; end: number; key: string } | undefined;
  for (let r = range.r1 + 1; r <= range.r2; r++) {
    if (isSubtotalRow(ws, r, range)) {
      cur = undefined;
      continue;
    }
    const key = keyAt(r);
    if (cur && cur.key.toLowerCase() === key.toLowerCase()) cur.end = r;
    else {
      cur = { start: r, end: r, key };
      runs.push(cur);
    }
  }
  return runs;
}

function outlineLevel(ws: Worksheet, row: number): number {
  return ws.rowDimensions.get(row)?.outlineLevel ?? 0;
}

function setOutlineLevel(ws: Worksheet, row: number, level: number): void {
  const { outlineLevel: _old, ...rest } = ws.rowDimensions.get(row) ?? {};
  if (level > 0) ws.rowDimensions.set(row, { ...rest, outlineLevel: Math.min(level, MAX_OUTLINE_LEVEL) });
  else if (Object.keys(rest).length > 0) ws.rowDimensions.set(row, rest);
  else ws.rowDimensions.delete(row);
}

/** Delete this command's rows from `range` and flatten its outline; returns the range left. */
function stripSubtotals(ctl: EditorController, range: Range): Range {
  const doc = ctl.doc;
  const ws = doc.ws;
  let r2 = range.r2;
  for (let r = range.r2; r > range.r1; r--) {
    if (!isSubtotalRow(ws, r, range)) continue;
    applyStructuralEdit(doc.wb, ws, { axis: 'row', at: r, count: -1 });
    r2--;
  }
  for (let r = range.r1; r <= r2; r++) {
    setOutlineLevel(ws, r, 0);
    const dim = ws.rowDimensions.get(r);
    if (dim?.hidden || dim?.collapsed) {
      const { hidden: _h, collapsed: _c, ...rest } = dim;
      if (Object.keys(rest).length > 0) ws.rowDimensions.set(r, rest);
      else ws.rowDimensions.delete(r);
    }
  }
  ws.rowBreaks = ws.rowBreaks.filter((b) => b.id === undefined || b.id < range.r1 || b.id > r2);
  return { ...range, r2 };
}

/**
 * Insert the subtotals. `range` is the list with its header row first.
 * Returns the range the list covers afterwards.
 */
export function applySubtotals(ctl: EditorController, range: Range, opts: SubtotalOptions, labels: SubtotalLabels): Range {
  const doc = ctl.doc;
  const ws = doc.ws;
  const code = SUBTOTAL_CODES[opts.fn];
  return doc.transact('Subtotal', (tx) => {
    declareStructural(tx, doc.wb, ws);
    const list = opts.replace ? stripSubtotals(ctl, range) : range;
    // Each inserted row pushes the list's last row down by one.
    let bottom = list.r2;
    const keyAt = (row: number) => {
      const cell = getCellAt(ws, row, opts.groupBy);
      return cell && !isBlank(cell) ? getCellDisplayText(doc.wb, cell) : '';
    };
    const runs = groupRuns(ws, keyAt, list);
    if (runs.length === 0) return list;
    // Excel bolds the label cells of the rows it adds.
    const boldMemo = new Map<number, number>();
    const bold = (styleId: number) => {
      let id = boldMemo.get(styleId);
      if (id === undefined) {
        id = transformStyle(doc.wb, styleId, { font: { bold: true } }, { top: false, bottom: false, left: false, right: false });
        boldMemo.set(styleId, id);
      }
      return id;
    };
    const writeRow = (row: number, label: string, from: number, to: number) => {
      const labelCell = setValueAt(ws, row, opts.groupBy, label, ctl.defaultStyleAt(row, opts.groupBy));
      labelCell.styleId = bold(labelCell.styleId);
      for (const col of opts.columns) {
        if (col === opts.groupBy) continue;
        const L = colLetter(col);
        // Number formats follow the column, as the inserted row inherits them in Excel.
        const style = getCellAt(ws, from, col)?.styleId ?? ctl.defaultStyleAt(row, col);
        setValueAt(ws, row, col, { kind: 'formula', t: 'normal', formula: `SUBTOTAL(${code},${L}${from}:${L}${to})` }, style);
      }
    };

    const firstData = list.r1 + 1;
    const nested = runs.some((run) => outlineLevel(ws, run.start) > 0);
    // Bottom-up, so earlier groups keep their row numbers while later ones grow.
    for (let i = runs.length - 1; i >= 0; i--) {
      const run = runs[i];
      if (!run) continue;
      const detailLevel = Math.max(outlineLevel(ws, run.start), 1);
      const at = opts.summaryBelow ? run.end + 1 : run.start;
      applyStructuralEdit(doc.wb, ws, { axis: 'row', at, count: 1 });
      const from = opts.summaryBelow ? run.start : run.start + 1;
      const to = opts.summaryBelow ? run.end : run.end + 1;
      writeRow(at, labels.group(run.key), from, to);
      for (let r = from; r <= to; r++) setOutlineLevel(ws, r, Math.max(outlineLevel(ws, r), 1) + 1);
      setOutlineLevel(ws, at, detailLevel);
      if (opts.pageBreaks && i < runs.length - 1) {
        const after = opts.summaryBelow ? at : to;
        ws.rowBreaks = [...ws.rowBreaks.filter((b) => b.id !== after), { id: after, max: MAX_COL - 1, man: true }];
      }
      bottom++;
    }
    // One grand total per list; nested subtotals reuse the existing one.
    if (!nested) {
      const at = opts.summaryBelow ? bottom + 1 : firstData;
      applyStructuralEdit(doc.wb, ws, { axis: 'row', at, count: 1 });
      bottom++;
      const from = opts.summaryBelow ? firstData : firstData + 1;
      const to = opts.summaryBelow ? at - 1 : bottom;
      writeRow(at, labels.grand, from, to);
      setOutlineLevel(ws, at, 0);
    }
    ws.sheetProperties = {
      ...ws.sheetProperties,
      outlinePr: { ...ws.sheetProperties?.outlinePr, summaryBelow: opts.summaryBelow },
    };
    return { ...list, r2: bottom };
  }) ?? range;
}

/** Subtotal ▸ Remove All. Returns the range left. */
export function removeSubtotals(ctl: EditorController, range: Range): Range {
  const doc = ctl.doc;
  return doc.transact('Remove Subtotals', (tx) => {
    declareStructural(tx, doc.wb, doc.ws);
    return stripSubtotals(ctl, range);
  }) ?? range;
}
