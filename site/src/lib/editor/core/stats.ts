// Status-bar statistics over the selection, visiting only populated cells
// so selecting whole columns stays cheap.

import type { Cell } from '@office-kit/xlsx/cell';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { forEachCellInRange } from './cells.ts';
import type { Selection } from './selection.ts';

export interface SelectionStats {
  /** Non-empty cells (Excel's "Count"). */
  readonly count: number;
  readonly numCount: number;
  readonly sum: number;
  readonly min: number;
  readonly max: number;
}

export function selectionStats(ws: Worksheet, sel: Selection): SelectionStats {
  let count = 0;
  let numCount = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  const seen = sel.ranges.length > 1 ? new Set<string>() : null;
  const visit = (cell: Cell): void => {
    if (seen) {
      const k = `${cell.row},${cell.col}`;
      if (seen.has(k)) return;
      seen.add(k);
    }
    let v = cell.value;
    if (v === null || v === '') return;
    if (typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula') {
      if (v.cachedValue === undefined || v.cachedValue === '') return;
      v = v.cachedValueType === 'error' ? null : v.cachedValue;
      count++;
    } else count++;
    if (typeof v === 'number' && Number.isFinite(v)) {
      numCount++;
      sum += v;
      if (v < min) min = v;
      if (v > max) max = v;
    }
  };
  for (const range of sel.ranges) forEachCellInRange(ws, range, visit);
  return { count, numCount, sum, min, max };
}
