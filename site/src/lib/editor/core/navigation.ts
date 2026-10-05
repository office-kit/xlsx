// Keyboard navigation semantics that depend on cell contents: Ctrl/Cmd+Arrow
// (jump to the edge of a data block), the current region (Ctrl/Cmd+A,
// Ctrl+Shift+*), and the last used cell (Ctrl/Cmd+End). Also View ▸
// Navigation: each sheet and the elements on it (tables, named ranges,
// PivotTables, charts, pictures, shapes), the way Excel's Navigation pane
// lists them. That list walks object lists only, never cells, so it stays
// cheap on large sheets.

import type { Workbook } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import type { CellPos, Range } from './address.ts';
import { MAX_COL, MAX_ROW, parseRangeAddress } from './address.ts';
import { getCellAt, isBlank } from './cells.ts';
import { tableRange } from './tables.ts';

/** Sorted non-blank indices along one line of the sheet. */
function filledAlong(ws: Worksheet, fixed: number, vertical: boolean): number[] {
  const out: number[] = [];
  if (vertical) {
    for (const [r, rowMap] of ws.rows) if (!isBlank(rowMap.get(fixed))) out.push(r);
  } else {
    for (const [c, cell] of ws.rows.get(fixed) ?? []) if (!isBlank(cell)) out.push(c);
  }
  return out.sort((a, b) => a - b);
}

function lowerBound(sorted: readonly number[], v: number): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if ((sorted[mid] ?? 0) < v) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/**
 * Where Ctrl/Cmd+Arrow lands from `from`. Excel's rule: if the current and the
 * next cell both hold data, run to the last filled cell before a gap;
 * otherwise jump to the next filled cell, or to the sheet edge when none.
 */
export function dataEdge(ws: Worksheet, from: CellPos, dRow: -1 | 0 | 1, dCol: -1 | 0 | 1): CellPos {
  const vertical = dRow !== 0;
  const step = vertical ? dRow : dCol;
  const pos = vertical ? from.row : from.col;
  const limit = vertical ? MAX_ROW : MAX_COL;
  const filled = filledAlong(ws, vertical ? from.col : from.row, vertical);
  const has = (i: number): boolean => {
    const k = lowerBound(filled, i);
    return filled[k] === i;
  };
  const make = (i: number): CellPos => (vertical ? { row: i, col: from.col } : { row: from.row, col: i });
  const next = pos + step;
  if (next < 1 || next > limit) return from;

  if (has(pos) && has(next)) {
    // Walk the contiguous block; `filled` is sorted so consecutive entries differ by 1 inside it.
    let k = lowerBound(filled, next);
    if (step === 1) {
      while (filled[k + 1] === (filled[k] ?? 0) + 1) k++;
    } else {
      while (k > 0 && filled[k - 1] === (filled[k] ?? 0) - 1) k--;
    }
    return make(filled[k] ?? pos);
  }
  if (step === 1) {
    const k = lowerBound(filled, next);
    return make(filled[k] ?? limit);
  }
  const k = lowerBound(filled, next + 1) - 1;
  return make(k >= 0 ? (filled[k] ?? 1) : 1);
}

/**
 * The block of non-blank cells around `pos` bounded by blank rows and columns
 * (Excel's CurrentRegion). Grows the rectangle until every cell on its border
 * ring is blank.
 */
export function currentRegion(ws: Worksheet, pos: CellPos): Range {
  let r1 = pos.row;
  let r2 = pos.row;
  let c1 = pos.col;
  let c2 = pos.col;
  const rowHasData = (r: number, a: number, b: number): boolean => {
    const rowMap = ws.rows.get(r);
    if (!rowMap) return false;
    for (const [c, cell] of rowMap) if (c >= a && c <= b && !isBlank(cell)) return true;
    return false;
  };
  const colHasData = (c: number, a: number, b: number): boolean => {
    for (let r = a; r <= b; r++) if (!isBlank(getCellAt(ws, r, c))) return true;
    return false;
  };
  for (;;) {
    let grew = false;
    if (r1 > 1 && rowHasData(r1 - 1, Math.max(1, c1 - 1), Math.min(MAX_COL, c2 + 1))) {
      r1--;
      grew = true;
    }
    if (r2 < MAX_ROW && rowHasData(r2 + 1, Math.max(1, c1 - 1), Math.min(MAX_COL, c2 + 1))) {
      r2++;
      grew = true;
    }
    if (c1 > 1 && colHasData(c1 - 1, Math.max(1, r1 - 1), Math.min(MAX_ROW, r2 + 1))) {
      c1--;
      grew = true;
    }
    if (c2 < MAX_COL && colHasData(c2 + 1, Math.max(1, r1 - 1), Math.min(MAX_ROW, r2 + 1))) {
      c2++;
      grew = true;
    }
    if (!grew) break;
  }
  return { r1, c1, r2, c2 };
}

/** Ctrl/Cmd+End: the intersection of the last used row and the last used column. */
export function lastUsedCell(ws: Worksheet): CellPos {
  let row = 1;
  let col = 1;
  for (const [r, rowMap] of ws.rows) {
    if (rowMap.size === 0) continue;
    if (r > row) row = r;
    for (const c of rowMap.keys()) if (c > col) col = c;
  }
  return { row, col };
}

export type NavTarget = { kind: 'range'; range: Range } | { kind: 'drawing'; index: number };

export interface NavElement {
  readonly kind: 'table' | 'name' | 'pivot' | 'chart' | 'picture' | 'shape';
  /** Undefined for an unnamed drawing; the pane shows "Chart 2" etc. from `ordinal`. */
  readonly label: string | undefined;
  readonly ordinal: number;
  readonly target: NavTarget;
}

export interface NavSheet {
  readonly index: number;
  readonly title: string;
  readonly hidden: boolean;
  readonly elements: readonly NavElement[];
}

export function navigationTree(wb: Workbook): NavSheet[] {
  const titleIndex = new Map(wb.sheets.map((s, i) => [s.sheet.title.toLowerCase(), i]));
  // Names are attached to the sheet their reference points at.
  const names = new Map<number, NavElement[]>();
  for (const dn of wb.definedNames) {
    // Excel lists user names only, not _xlnm.Print_Area and friends.
    if (dn.name.startsWith('_xlnm.') || dn.hidden) continue;
    const parsed = parseRangeAddress(dn.value.replace(/^=/, ''));
    if (!parsed) continue;
    // An unqualified reference belongs to the sheet the name is scoped to.
    const index = parsed.sheet === undefined ? dn.scope : titleIndex.get(parsed.sheet.toLowerCase());
    if (index === undefined) continue;
    const list = names.get(index) ?? [];
    list.push({ kind: 'name', label: dn.name, ordinal: 0, target: { kind: 'range', range: parsed.range } });
    names.set(index, list);
  }
  return wb.sheets.map((ref, index) => {
    const elements: NavElement[] = [];
    if (ref.kind === 'worksheet') {
      const ws = ref.sheet;
      for (const def of ws.tables) {
        const range = tableRange(def);
        if (range) elements.push({ kind: 'table', label: def.displayName, ordinal: 0, target: { kind: 'range', range } });
      }
      elements.push(...(names.get(index) ?? []));
      for (const pt of ws.pivotTables ?? []) {
        const range = pt.renderedRef ? parseRangeAddress(pt.renderedRef)?.range : undefined;
        if (range) elements.push({ kind: 'pivot', label: pt.name, ordinal: 0, target: { kind: 'range', range } });
      }
      (ws.drawing?.items ?? []).forEach((item, i) => {
        const c = item.content;
        const target: NavTarget = { kind: 'drawing', index: i };
        if (c.kind === 'chart') elements.push({ kind: 'chart', label: undefined, ordinal: i + 1, target });
        else if (c.kind === 'picture') elements.push({ kind: 'picture', label: c.picture.name, ordinal: i + 1, target });
        else if (c.kind === 'shape') elements.push({ kind: 'shape', label: c.shape.name, ordinal: i + 1, target });
      });
    }
    return { index, title: ref.sheet.title, hidden: ref.state !== 'visible', elements };
  });
}
