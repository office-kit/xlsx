// Find & Replace and Go To Special over the sparse cell store.

import type { Cell } from '@office-kit/xlsx/cell';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { toStorageFormula } from '../calc/index.ts';
import type { CellPos } from './address.ts';
import type { EditorController } from './controller.svelte.ts';
import { editTextFor } from './input.ts';
import { selectRange } from './selection.ts';

export interface FindOptions {
  readonly query: string;
  readonly matchCase: boolean;
  readonly wholeCell: boolean;
  readonly byColumns: boolean;
  readonly lookIn: 'formulas' | 'values';
}

function cellText(ctl: EditorController, cell: Cell, lookIn: 'formulas' | 'values'): string {
  if (lookIn === 'values') return getCellDisplayText(ctl.doc.wb, cell);
  return editTextFor(cell.value, getCellDisplayText(ctl.doc.wb, cell), false);
}

/** Regex source for an Excel wildcard pattern: `*` any run, `?` one character, `~` escapes. */
function wildcardSource(query: string): string {
  let source = '';
  for (let i = 0; i < query.length; i++) {
    const ch = query[i] ?? '';
    if (ch === '~' && i + 1 < query.length) source += (query[++i] ?? '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (ch === '*') source += '.*';
    else if (ch === '?') source += '.';
    else source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return source;
}

export function toMatcher(opts: FindOptions): (text: string) => boolean {
  const source = wildcardSource(opts.query);
  const re = new RegExp(opts.wholeCell ? `^${source}$` : source, opts.matchCase ? 's' : 'is');
  return (text) => re.test(text);
}

function sortedCells(ws: Worksheet, byColumns: boolean): Cell[] {
  const cells: Cell[] = [];
  for (const rowMap of ws.rows.values()) for (const cell of rowMap.values()) if (cell.value !== null) cells.push(cell);
  cells.sort(byColumns ? (a, b) => a.col - b.col || a.row - b.row : (a, b) => a.row - b.row || a.col - b.col);
  return cells;
}

export function findAll(ctl: EditorController, opts: FindOptions, sheets: readonly Worksheet[]): Array<{ ws: Worksheet; cell: Cell; text: string }> {
  if (!opts.query) return [];
  const match = toMatcher(opts);
  const out: Array<{ ws: Worksheet; cell: Cell; text: string }> = [];
  for (const ws of sheets) {
    for (const cell of sortedCells(ws, opts.byColumns)) {
      const text = cellText(ctl, cell, opts.lookIn);
      if (match(text)) out.push({ ws, cell, text });
    }
  }
  return out;
}

/** Move the selection to the next match after the active cell (wrapping). */
export function findNext(ctl: EditorController, direction: 1 | -1): boolean {
  const opts = ctl.findState;
  if (!opts.query) {
    ctl.openDialog('find');
    return false;
  }
  const ws = ctl.doc.ws;
  const matches = findAll(ctl, opts, [ws]);
  if (matches.length === 0) {
    ctl.toast = 'notFound';
    return false;
  }
  const a = ctl.doc.selection.active;
  const after = (p: CellPos) => (opts.byColumns ? p.col > a.col || (p.col === a.col && p.row > a.row) : p.row > a.row || (p.row === a.row && p.col > a.col));
  const before = (p: CellPos) => (opts.byColumns ? p.col < a.col || (p.col === a.col && p.row < a.row) : p.row < a.row || (p.row === a.row && p.col < a.col));
  const hit = direction === 1 ? (matches.find((m) => after(m.cell)) ?? matches[0]) : ([...matches].reverse().find((m) => before(m.cell)) ?? matches.at(-1));
  if (!hit) return false;
  ctl.selectCell({ row: hit.cell.row, col: hit.cell.col });
  ctl.reveal(hit.cell.row, hit.cell.col);
  return true;
}

function replaceIn(ctl: EditorController, opts: FindOptions, replacement: string, targets: ReadonlyArray<{ ws: Worksheet; cell: Cell }>): void {
  const re = buildReplaceRegex(opts);
  ctl.doc.transact('Replace', (tx) => {
    for (const { ws, cell } of targets) {
      tx.cells(ws, { r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });
      const text = cellText(ctl, cell, 'formulas');
      const next = opts.wholeCell ? replacement : text.replace(re, replacement.replaceAll('$', '$$$$'));
      cell.value = next.startsWith('=') && next.length > 1 ? { kind: 'formula', t: 'normal', formula: toStorageFormula(next.slice(1)) } : coerce(next);
    }
  });
}

/** Replace in the formulas/text of every matching cell; returns the count. */
export function replaceAll(ctl: EditorController, opts: FindOptions, replacement: string, sheets: readonly Worksheet[]): number {
  const matches = findAll(ctl, { ...opts, lookIn: 'formulas' }, sheets);
  if (matches.length > 0) replaceIn(ctl, opts, replacement, matches);
  return matches.length;
}

function buildReplaceRegex(opts: FindOptions): RegExp {
  return new RegExp(wildcardSource(opts.query), opts.matchCase ? 'g' : 'gi');
}

function coerce(text: string): string | number {
  const n = Number(text);
  return text.trim() !== '' && Number.isFinite(n) ? n : text;
}

/** Replace button: replace the active cell if it matches, then move to the next match. */
export function replaceCurrent(ctl: EditorController, opts: FindOptions, replacement: string): void {
  const { row, col } = ctl.doc.selection.active;
  const ws = ctl.doc.ws;
  const cell = ws.rows.get(row)?.get(col);
  if (cell && toMatcher({ ...opts, lookIn: 'formulas' })(cellText(ctl, cell, 'formulas'))) replaceIn(ctl, opts, replacement, [{ ws, cell }]);
  findNext(ctl, 1);
}

export type GoToSpecial = 'blanks' | 'constants' | 'formulas' | 'comments' | 'currentRegion' | 'lastCell' | 'visible' | 'errors' | 'dataValidation' | 'conditionalFormats';

export function selectCells(ctl: EditorController, cells: readonly CellPos[]): void {
  if (cells.length === 0) {
    ctl.toast = 'noCellsFound';
    return;
  }
  const ranges = cells.slice(0, 2000).map((p) => ({ r1: p.row, c1: p.col, r2: p.row, c2: p.col }));
  const first = cells[0] ?? { row: 1, col: 1 };
  ctl.doc.setSelection({ ...selectRange(ranges[0] ?? { r1: 1, c1: 1, r2: 1, c2: 1 }, first), ranges });
}
