// Review ▸ Workbook Statistics: the counts Excel's dialog lists for the
// current sheet and for the whole workbook. One pass over each sheet's
// stored cells.

import type { Workbook } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { cellAddress } from './address.ts';
import { lastUsedCell } from './navigation.ts';

export interface SheetCounts {
  cells: number;
  tables: number;
  pivotTables: number;
  formulas: number;
  charts: number;
  images: number;
  comments: number;
  notes: number;
}

export const COUNT_KEYS: ReadonlyArray<keyof SheetCounts> = ['cells', 'tables', 'pivotTables', 'formulas', 'charts', 'images', 'comments', 'notes'];

function countSheet(ws: Worksheet): SheetCounts {
  const out: SheetCounts = { cells: 0, tables: ws.tables.length, pivotTables: ws.pivotTables?.length ?? 0, formulas: 0, charts: 0, images: 0, comments: 0, notes: ws.legacyComments.length };
  for (const rowMap of ws.rows.values()) {
    for (const cell of rowMap.values()) {
      const v = cell.value;
      if (v === null) continue;
      out.cells++;
      if (typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula') out.formulas++;
    }
  }
  for (const item of ws.drawing?.items ?? []) {
    if (item.content.kind === 'chart') out.charts++;
    else if (item.content.kind === 'picture') out.images++;
  }
  // Replies belong to their thread; Excel counts threads.
  out.comments = (ws.threadedComments ?? []).filter((c) => c.parentId === undefined).length;
  // Each thread also carries a legacy placeholder note, which isn't a note of its own.
  out.notes = Math.max(0, out.notes - out.comments);
  return out;
}

export interface WorkbookStatistics {
  readonly endOfSheet: string;
  readonly sheet: SheetCounts;
  readonly sheets: number;
  readonly workbook: SheetCounts;
}

export function workbookStatistics(wb: Workbook, ws: Worksheet): WorkbookStatistics {
  const total: SheetCounts = { cells: 0, tables: 0, pivotTables: 0, formulas: 0, charts: 0, images: 0, comments: 0, notes: 0 };
  let current: SheetCounts | undefined;
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') {
      total.charts++;
      continue;
    }
    const counts = countSheet(ref.sheet);
    if (ref.sheet === ws) current = counts;
    for (const k of COUNT_KEYS) total[k] += counts[k];
  }
  const end = lastUsedCell(ws);
  return { endOfSheet: cellAddress(end.row, end.col), sheet: current ?? countSheet(ws), sheets: wb.sheets.length, workbook: total };
}
