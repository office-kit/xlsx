// Row/column outlines (Data ▸ Group): the runs of grouped lines per level,
// collapsing/expanding them from the +/− buttons, and the level buttons.
//
// Excel keeps `collapsed` on the summary line (the one after a group when
// summaries are below/right, before it otherwise) and `hidden` on the
// grouped lines; nested groups keep their own state when a parent expands.

import { getColumnDimension, setColumnDimension, type Worksheet } from '@office-kit/xlsx/worksheet';
import { MAX_COL, MAX_ROW } from './address.ts';
import type { SpreadsheetEditor } from './editor.svelte.ts';

export type OutlineAxis = 'row' | 'col';

/** Excel allows eight outline levels (seven nested groups). */
export const MAX_OUTLINE_LEVEL = 7;

export interface OutlineRun {
  readonly level: number;
  readonly start: number;
  readonly end: number;
  /** Line holding the +/− button and the `collapsed` flag; may lie outside the sheet. */
  readonly summary: number;
}

export interface AxisOutline {
  readonly maxLevel: number;
  /** Runs ordered by start, all levels mixed. */
  readonly runs: readonly OutlineRun[];
}

interface LineState {
  level: number;
  hidden: boolean;
  collapsed: boolean;
}

function lineState(ws: Worksheet, axis: OutlineAxis, i: number): LineState {
  const d = axis === 'row' ? ws.rowDimensions.get(i) : getColumnDimension(ws, i);
  return { level: d?.outlineLevel ?? 0, hidden: d?.hidden === true, collapsed: d?.collapsed === true };
}

function summaryAfter(ws: Worksheet, axis: OutlineAxis): boolean {
  const pr = ws.sheetProperties?.outlinePr;
  return axis === 'row' ? pr?.summaryBelow !== false : pr?.summaryRight !== false;
}

/** Grouped lines of one axis, sorted, with their levels. */
function groupedLines(ws: Worksheet, axis: OutlineAxis): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  if (axis === 'row') {
    for (const [r, d] of ws.rowDimensions) if (d.outlineLevel) out.push([r, d.outlineLevel]);
  } else {
    for (const d of ws.columnDimensions.values()) {
      if (!d.outlineLevel) continue;
      for (let c = d.min; c <= Math.min(d.max, MAX_COL); c++) out.push([c, d.outlineLevel]);
    }
  }
  return out.sort((a, b) => a[0] - b[0]);
}

export function buildOutline(ws: Worksheet, axis: OutlineAxis): AxisOutline {
  const lines = groupedLines(ws, axis);
  if (lines.length === 0) return { maxLevel: 0, runs: [] };
  const after = summaryAfter(ws, axis);
  const maxLevel = lines.reduce((m, [, l]) => Math.max(m, l), 0);
  const runs: OutlineRun[] = [];
  for (let level = 1; level <= maxLevel; level++) {
    let start = -1;
    let prev = -1;
    const close = () => {
      if (start > 0) runs.push({ level, start, end: prev, summary: after ? prev + 1 : start - 1 });
    };
    for (const [i, l] of lines) {
      if (l < level) continue;
      if (start > 0 && i === prev + 1) {
        prev = i;
        continue;
      }
      close();
      start = prev = i;
    }
    close();
  }
  runs.sort((a, b) => a.start - b.start || a.level - b.level);
  return { maxLevel, runs };
}

export function isCollapsed(ws: Worksheet, axis: OutlineAxis, run: OutlineRun): boolean {
  for (let i = run.start; i <= run.end; i++) if (!lineState(ws, axis, i).hidden) return false;
  return true;
}

function patchLine(ws: Worksheet, axis: OutlineAxis, i: number, patch: { hidden?: boolean; collapsed?: boolean }): void {
  const limit = axis === 'row' ? MAX_ROW : MAX_COL;
  if (i < 1 || i > limit) return;
  if (axis === 'col') {
    const cur = getColumnDimension(ws, i);
    const next = { hidden: patch.hidden ?? cur?.hidden === true, collapsed: patch.collapsed ?? cur?.collapsed === true };
    setColumnDimension(ws, i, next);
    return;
  }
  const cur = ws.rowDimensions.get(i) ?? {};
  const next: Record<string, unknown> = { ...cur };
  for (const [k, v] of Object.entries(patch)) {
    if (v) next[k] = true;
    else delete next[k];
  }
  if (Object.keys(next).length === 0) ws.rowDimensions.delete(i);
  else ws.rowDimensions.set(i, next);
}

/** Lines of `run` that stay hidden on expand because a nested group is still collapsed. */
function nestedCollapsed(ws: Worksheet, axis: OutlineAxis, outline: AxisOutline, run: OutlineRun): Set<number> {
  const keep = new Set<number>();
  for (const inner of outline.runs) {
    if (inner.level <= run.level || inner.start < run.start || inner.end > run.end) continue;
    if (!lineState(ws, axis, inner.summary).collapsed) continue;
    for (let i = inner.start; i <= inner.end; i++) keep.add(i);
  }
  return keep;
}

/** The +/− button: collapse or expand one group, as one undo step. */
export function toggleRun(doc: SpreadsheetEditor, axis: OutlineAxis, outline: AxisOutline, run: OutlineRun): void {
  const ws = doc.ws;
  const collapse = !isCollapsed(ws, axis, run);
  doc.transact(collapse ? 'Hide Detail' : 'Show Detail', (tx) => {
    tx.sheet(ws, axis === 'row' ? 'rowDimensions' : 'columnDimensions');
    const keep = collapse ? new Set<number>() : nestedCollapsed(ws, axis, outline, run);
    for (let i = run.start; i <= run.end; i++) patchLine(ws, axis, i, { hidden: collapse || keep.has(i) });
    patchLine(ws, axis, run.summary, { collapsed: collapse });
  });
}

/** Level button `n`: show lines with level < n, hide the rest. */
export function showLevel(doc: SpreadsheetEditor, axis: OutlineAxis, outline: AxisOutline, n: number): void {
  const ws = doc.ws;
  doc.transact('Show Level', (tx) => {
    tx.sheet(ws, axis === 'row' ? 'rowDimensions' : 'columnDimensions');
    for (const run of outline.runs) {
      const hide = run.level >= n;
      for (let i = run.start; i <= run.end; i++) {
        if (lineState(ws, axis, i).level === run.level) patchLine(ws, axis, i, { hidden: hide });
      }
      patchLine(ws, axis, run.summary, { collapsed: hide });
    }
  });
}

/** Data ▸ Group: one level deeper than the deepest group already in the span. */
export function groupLines(doc: SpreadsheetEditor, axis: OutlineAxis, from: number, to: number, group: boolean): void {
  const ws = doc.ws;
  doc.transact(group ? 'Group' : 'Ungroup', (tx) => {
    tx.sheet(ws, axis === 'row' ? 'rowDimensions' : 'columnDimensions');
    for (let i = from; i <= to; i++) {
      const level = lineState(ws, axis, i).level;
      const next = group ? Math.min(MAX_OUTLINE_LEVEL, level + 1) : Math.max(0, level - 1);
      if (next === level) continue;
      if (axis === 'col') {
        setColumnDimension(ws, i, { outlineLevel: next });
        continue;
      }
      const { outlineLevel: _old, ...rest } = ws.rowDimensions.get(i) ?? {};
      if (next > 0) ws.rowDimensions.set(i, { ...rest, outlineLevel: next });
      else if (Object.keys(rest).length > 0) ws.rowDimensions.set(i, rest);
      else ws.rowDimensions.delete(i);
    }
  });
}
