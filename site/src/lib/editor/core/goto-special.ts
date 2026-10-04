// Go To Special: compute the cells each option selects. With a single cell
// selected Excel searches the whole used range; with a larger selection it
// searches only inside it.

import type { CellValue } from '@office-kit/xlsx/cell';
import { columnIndexFromLetter, coordinateFromString } from '@office-kit/xlsx/utils';
import type { Range } from './address.ts';
import { fromBoundaries, inRange } from './address.ts';
import { getCellAt, usedRange } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import type { GoToSpecial } from './find.ts';
import { currentRegion, lastUsedCell } from './navigation.ts';
import { selectRange } from './selection.ts';

export interface ValueKinds {
  readonly numbers: boolean;
  readonly text: boolean;
  readonly logicals: boolean;
  readonly errors: boolean;
}

/** Most separate ranges a Go To Special selection keeps, so a huge result stays responsive. */
const MAX_RANGES = 2000;

type Kind = keyof ValueKinds;

function kindOf(v: Exclude<CellValue, null>): Kind | undefined {
  if (typeof v === 'number' || v instanceof Date) return 'numbers';
  if (typeof v === 'string') return 'text';
  if (typeof v === 'boolean') return 'logicals';
  switch (v.kind) {
    case 'error':
      return 'errors';
    case 'rich-text':
      return 'text';
    case 'formula':
      if (v.cachedValueType === 'error') return 'errors';
      if (typeof v.cachedValue === 'number') return 'numbers';
      if (typeof v.cachedValue === 'boolean') return 'logicals';
      return 'text';
    default:
      return 'numbers';
  }
}

/** Where to look: the used range for a single cell, else each selected range clipped to it. */
function searchAreas(ctl: EditorController): Range[] {
  const used = usedRange(ctl.doc.ws);
  if (!used) return [];
  const sel = ctl.doc.selection;
  const single = sel.ranges.length === 1 && sel.ranges.every((r) => r.r1 === r.r2 && r.c1 === r.c2);
  if (single) return [used];
  const out: Range[] = [];
  for (const r of sel.ranges) {
    const clipped = { r1: Math.max(r.r1, used.r1), c1: Math.max(r.c1, used.c1), r2: Math.min(r.r2, used.r2), c2: Math.min(r.c2, used.c2) };
    if (clipped.r1 <= clipped.r2 && clipped.c1 <= clipped.c2) out.push(clipped);
  }
  return out;
}

/** Merge cell hits on each row into horizontal runs, keeping the selection list short. */
function runs(cells: ReadonlyArray<{ row: number; col: number }>): Range[] {
  const sorted = cells.slice().sort((a, b) => a.row - b.row || a.col - b.col);
  const out: Range[] = [];
  for (const { row, col } of sorted) {
    const last = out.at(-1);
    if (last && last.r1 === row && last.c2 === col - 1) out[out.length - 1] = { ...last, c2: col };
    else out.push({ r1: row, c1: col, r2: row, c2: col });
  }
  return out;
}

function cellsWhere(ctl: EditorController, areas: readonly Range[], test: (value: CellValue, row: number, col: number) => boolean): Range[] {
  const ws = ctl.doc.ws;
  const hits: Array<{ row: number; col: number }> = [];
  for (const area of areas) {
    for (const [row, rowMap] of ws.rows) {
      if (row < area.r1 || row > area.r2) continue;
      for (const [col, cell] of rowMap) if (col >= area.c1 && col <= area.c2 && test(cell.value, row, col)) hits.push({ row, col });
    }
  }
  return runs(hits);
}

function blanks(ctl: EditorController, areas: readonly Range[]): Range[] {
  const ws = ctl.doc.ws;
  const hits: Array<{ row: number; col: number }> = [];
  for (const area of areas) {
    for (let r = area.r1; r <= area.r2; r++) {
      for (let c = area.c1; c <= area.c2; c++) {
        const v = getCellAt(ws, r, c)?.value ?? null;
        if (v === null || v === '') hits.push({ row: r, col: c });
      }
    }
  }
  return runs(hits);
}

function visibleCells(ctl: EditorController, areas: readonly Range[]): Range[] {
  const spans = (lo: number, hi: number, hidden: (i: number) => boolean): Array<[number, number]> => {
    const out: Array<[number, number]> = [];
    for (let i = lo; i <= hi; i++) {
      if (hidden(i)) continue;
      const last = out.at(-1);
      if (last && last[1] === i - 1) last[1] = i;
      else out.push([i, i]);
    }
    return out;
  };
  const out: Range[] = [];
  for (const area of areas) {
    for (const [r1, r2] of spans(area.r1, area.r2, (i) => ctl.doc.rows.isHidden(i))) {
      for (const [c1, c2] of spans(area.c1, area.c2, (i) => ctl.doc.cols.isHidden(i))) out.push({ r1, c1, r2, c2 });
    }
  }
  return out;
}

export function goToSpecialRanges(ctl: EditorController, kind: GoToSpecial, kinds: ValueKinds): Range[] {
  const ws = ctl.doc.ws;
  const areas = searchAreas(ctl);
  switch (kind) {
    case 'currentRegion':
      return [currentRegion(ws, ctl.doc.selection.active)];
    case 'lastCell': {
      const p = lastUsedCell(ws);
      return [{ r1: p.row, c1: p.col, r2: p.row, c2: p.col }];
    }
    case 'blanks':
      return blanks(ctl, areas);
    case 'visible':
      return visibleCells(ctl, areas);
    case 'constants':
      return cellsWhere(ctl, areas, (v) => {
        if (v === null || v === '' || (typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula')) return false;
        const k = kindOf(v);
        return k !== undefined && kinds[k];
      });
    case 'formulas':
      return cellsWhere(ctl, areas, (v) => {
        if (v === null || typeof v !== 'object' || v instanceof Date || v.kind !== 'formula') return false;
        const k = kindOf(v);
        return k !== undefined && kinds[k];
      });
    case 'errors':
      return cellsWhere(ctl, areas, (v) => v !== null && kindOf(v) === 'errors');
    case 'comments': {
      const noted = [...ws.legacyComments, ...(ws.threadedComments ?? [])].map((c) => {
        const p = coordinateFromString(c.ref.replaceAll('$', '').split(':')[0] ?? 'A1');
        return { row: p.row, col: columnIndexFromLetter(p.column) };
      });
      return runs(noted.filter((p) => areas.some((a) => inRange(a, p.row, p.col))));
    }
    case 'conditionalFormats':
      return ws.conditionalFormatting.flatMap((cf) => cf.sqref.ranges.map(fromBoundaries));
    case 'dataValidation':
      return ws.dataValidations.flatMap((dv) => dv.sqref.ranges.map(fromBoundaries));
  }
}

/** Select the ranges (active cell at the first); returns false when nothing matched. */
export function selectRanges(ctl: EditorController, ranges: readonly Range[]): boolean {
  const first = ranges[0];
  if (!first) return false;
  const kept = ranges.slice(0, MAX_RANGES);
  ctl.doc.setSelection({ ...selectRange(first), ranges: kept });
  ctl.reveal(first.r1, first.c1);
  return true;
}
