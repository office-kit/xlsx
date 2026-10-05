// Undo/redo for the spreadsheet editor.
//
// Serialising the whole workbook per edit (what the presentation editor does)
// does not scale to sheets with a million cells, so history here is
// patch-based: before mutating, a transaction declares what it is about to
// touch — a rectangle of cells, named worksheet fields, or workbook fields —
// and the declared parts are snapshotted. When the transaction ends the same
// parts are snapshotted again, so undo and redo are both "restore a snapshot"
// and never depend on an operation being invertible (formula #REF! rewrites,
// sorts and pastes are not).
//
// Cell values are immutable (frozen formula/rich-text objects or primitives),
// so a cell snapshot copies the cell's fields and shares the value. Worksheet
// and workbook fields are plain data and are deep-cloned with structuredClone,
// except the sheet list, whose entries must keep their identity.

import type { CellValue } from '@office-kit/xlsx/cell';
import type { SheetRef, Workbook } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { makeCell } from '@office-kit/xlsx/cell';
import type { Range } from './address.ts';
import { deleteCellsInRange, forEachCellInRange } from './cells.ts';

interface CellSnap {
  readonly row: number;
  readonly col: number;
  readonly value: CellValue;
  readonly styleId: number;
  readonly hyperlinkId: number | undefined;
  readonly commentId: number | undefined;
}

type SheetField = Exclude<keyof Worksheet, 'rows' | 'title'>;
type WorkbookField = keyof Workbook;

export type Part =
  | { readonly kind: 'cells'; readonly ws: Worksheet; readonly range: Range }
  | { readonly kind: 'sheet'; readonly ws: Worksheet; readonly fields: readonly SheetField[] }
  | { readonly kind: 'title'; readonly ws: Worksheet }
  | { readonly kind: 'workbook'; readonly fields: readonly WorkbookField[] };

type Snapshot = { readonly part: Part; readonly data: unknown };

/** A selection-like token restored with the step; the editor decides what it holds. */
export interface HistoryStep<V> {
  readonly label: string;
  readonly before: readonly Snapshot[];
  readonly after: readonly Snapshot[];
  readonly viewBefore: V;
  readonly viewAfter: V;
  /** Rows/columns/sheets/names moved: formula dependencies must be rebuilt, not patched. */
  readonly structural: boolean;
}

function snapCells(ws: Worksheet, range: Range): CellSnap[] {
  const out: CellSnap[] = [];
  forEachCellInRange(ws, range, (c) =>
    out.push({
      row: c.row,
      col: c.col,
      value: c.value,
      styleId: c.styleId,
      hyperlinkId: c.hyperlinkId,
      commentId: c.commentId,
    }),
  );
  return out;
}

function restoreCells(ws: Worksheet, range: Range, snaps: readonly CellSnap[]): void {
  deleteCellsInRange(ws, range);
  for (const s of snaps) {
    const cell = makeCell(s.row, s.col, s.value, s.styleId);
    if (s.hyperlinkId !== undefined) cell.hyperlinkId = s.hyperlinkId;
    if (s.commentId !== undefined) cell.commentId = s.commentId;
    let rowMap = ws.rows.get(s.row);
    if (!rowMap) {
      rowMap = new Map();
      ws.rows.set(s.row, rowMap);
    }
    rowMap.set(s.col, cell);
  }
}

function take(part: Part, wb: Workbook): Snapshot {
  switch (part.kind) {
    case 'cells':
      return { part, data: snapCells(part.ws, part.range) };
    case 'sheet': {
      const record: Record<string, unknown> = {};
      for (const f of part.fields) record[f] = structuredClone(part.ws[f]);
      return { part, data: record };
    }
    case 'title':
      return { part, data: part.ws.title };
    case 'workbook': {
      const record: Record<string, unknown> = {};
      // Sheet entries are kept by identity (they hold the sheets); their visibility is copied, since Hide/Unhide changes it in place.
      for (const f of part.fields) record[f] = f === 'sheets' ? wb.sheets.map((ref) => ({ ref, state: ref.state })) : structuredClone(wb[f]);
      return { part, data: record };
    }
  }
}

type SheetSnap = { readonly ref: SheetRef; readonly state: SheetRef['state'] };

function restore(snapshot: Snapshot, wb: Workbook): void {
  const { part, data } = snapshot;
  switch (part.kind) {
    case 'cells':
      restoreCells(part.ws, part.range, data as CellSnap[]);
      return;
    case 'sheet': {
      const record = data as Record<string, unknown>;
      // Clone again: the snapshot must survive later edits of the restored state.
      for (const f of part.fields) Object.assign(part.ws, { [f]: structuredClone(record[f]) });
      return;
    }
    case 'title':
      part.ws.title = data as string;
      return;
    case 'workbook': {
      const record = data as Record<string, unknown>;
      for (const f of part.fields) {
        if (f === 'sheets') {
          wb.sheets = (record[f] as SheetSnap[]).map(({ ref, state }) => {
            ref.state = state;
            return ref;
          });
        } else Object.assign(wb, { [f]: structuredClone(record[f]) });
      }
    }
  }
}

/** Thrown by a transaction guard to refuse an edit (sheet protection); the step is rolled back. */
export class EditRefusedError extends Error {
  constructor(readonly reason: 'protectedCell' | 'protectedStructure') {
    super(reason);
    this.name = 'EditRefusedError';
  }
}

/**
 * Inspects each part before it is declared; throws EditRefusedError to refuse
 * the edit. `structural` says the step moves cells or renames things, whose
 * formula rewrites touch cells the user didn't edit.
 */
export type TransactionGuard = (part: Part, structural: boolean) => void;

/** Collects the parts an edit touches; handed to the mutation callback. */
export class Transaction {
  readonly #wb: Workbook;
  readonly #guard: TransactionGuard | undefined;
  readonly parts: Part[] = [];
  readonly before: Snapshot[] = [];
  /** Set by edits that move cells or rename/add/remove sheets or names. */
  structural = false;

  constructor(wb: Workbook, guard?: TransactionGuard) {
    this.#wb = wb;
    this.#guard = guard;
  }

  #add(part: Part): void {
    this.#guard?.(part, this.structural);
    this.parts.push(part);
    this.before.push(take(part, this.#wb));
  }

  /** Declare that cells inside `range` (values, styles, links) are about to change. */
  cells(ws: Worksheet, range: Range): void {
    this.#add({ kind: 'cells', ws, range });
  }

  /** Declare that non-cell worksheet fields (merges, dimensions, views, …) are about to change. */
  sheet(ws: Worksheet, ...fields: SheetField[]): void {
    this.#add({ kind: 'sheet', ws, fields });
  }

  title(ws: Worksheet): void {
    this.#add({ kind: 'title', ws });
  }

  workbook(...fields: WorkbookField[]): void {
    this.#add({ kind: 'workbook', fields });
  }

  /** Everything a structural edit of `ws` can reach: all its cells and fields. */
  wholeSheet(ws: Worksheet): void {
    this.structural = true;
    this.cells(ws, { r1: 1, c1: 1, r2: 1_048_576, c2: 16_384 });
    const fields = Object.keys(ws).filter((k): k is SheetField => k !== 'rows' && k !== 'title');
    this.sheet(ws, ...fields);
  }

  /** Undo whatever the callback already changed (an edit refused part-way). */
  rollback(): void {
    for (let i = this.before.length - 1; i >= 0; i--) {
      const snap = this.before[i];
      if (snap) restore(snap, this.#wb);
    }
  }

  /** Snapshot the declared parts again after the mutation ran. */
  finish(): Snapshot[] {
    // Restore order matters for redo too: replay in declaration order.
    return this.parts.map((p) => take(p, this.#wb));
  }
}

const HISTORY_LIMIT = 200;

export class History<V> {
  readonly #wb: () => Workbook;
  #steps: HistoryStep<V>[] = [];
  #cursor = 0;

  constructor(wb: () => Workbook) {
    this.#wb = wb;
  }

  get canUndo(): boolean {
    return this.#cursor > 0;
  }

  get canRedo(): boolean {
    return this.#cursor < this.#steps.length;
  }

  get undoLabel(): string | undefined {
    return this.#steps[this.#cursor - 1]?.label;
  }

  get redoLabel(): string | undefined {
    return this.#steps[this.#cursor]?.label;
  }

  push(step: HistoryStep<V>): void {
    this.#steps.length = this.#cursor;
    this.#steps.push(step);
    if (this.#steps.length > HISTORY_LIMIT) this.#steps.shift();
    this.#cursor = this.#steps.length;
  }

  /** Restore the state before the last step; returns the step so the caller can restore its view. */
  undo(): HistoryStep<V> | undefined {
    const step = this.#steps[this.#cursor - 1];
    if (!step) return undefined;
    // Later declarations may overlap earlier ones; undo restores in reverse.
    for (let i = step.before.length - 1; i >= 0; i--) {
      const snap = step.before[i];
      if (snap) restore(snap, this.#wb());
    }
    this.#cursor--;
    return step;
  }

  redo(): HistoryStep<V> | undefined {
    const step = this.#steps[this.#cursor];
    if (!step) return undefined;
    for (const snap of step.after) restore(snap, this.#wb());
    this.#cursor++;
    return step;
  }

  clear(): void {
    this.#steps = [];
    this.#cursor = 0;
  }
}

/**
 * Every cell address a step wrote or removed (union of its before and after
 * cell snapshots). Cheap — proportional to the cells touched, not the ranges
 * declared — and it includes cells that were deleted, whose formula
 * dependents still need recalculating.
 */
/** Larger than any column index, so `row * COL_KEY_SPAN + col` is unique. */
const COL_KEY_SPAN = 16_385;

export function changedCells<V>(step: HistoryStep<V>): Array<{ ws: Worksheet; row: number; col: number }> {
  // Only value changes count: a format applied over a spilled range must not
  // read as the user overwriting the spill. Numeric keys per sheet: a string
  // key per cell made a 2M-cell sort spend seconds here.
  const before = new Map<Worksheet, Map<number, CellValue>>();
  for (const snap of step.before) {
    if (snap.part.kind !== 'cells') continue;
    const ws = snap.part.ws;
    let values = before.get(ws);
    if (!values) before.set(ws, (values = new Map()));
    for (const c of snap.data as CellSnap[]) values.set(c.row * COL_KEY_SPAN + c.col, c.value);
  }
  const out: Array<{ ws: Worksheet; row: number; col: number }> = [];
  for (const snap of step.after) {
    if (snap.part.kind !== 'cells') continue;
    const ws = snap.part.ws;
    const values = before.get(ws);
    for (const c of snap.data as CellSnap[]) {
      const key = c.row * COL_KEY_SPAN + c.col;
      const was = values?.get(key);
      values?.delete(key);
      if (!sameValue(was ?? null, c.value)) out.push({ ws, row: c.row, col: c.col });
    }
  }
  // Cells that existed before and are gone after.
  for (const [ws, values] of before) {
    for (const [key, value] of values) if (value !== null) out.push({ ws, row: Math.floor(key / COL_KEY_SPAN), col: key % COL_KEY_SPAN });
  }
  return out;
}

function sameValue(a: CellValue, b: CellValue): boolean {
  if (a === b) return true;
  return a instanceof Date && b instanceof Date && a.getTime() === b.getTime();
}
