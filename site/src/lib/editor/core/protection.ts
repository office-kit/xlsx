// Sheet and workbook protection, enforced at the transaction layer: every
// edit declares what it touches before mutating, so one guard covers typing,
// paste, fill, formatting, sorting and structure changes alike.
//
// SheetProtection flags are true when the action is *locked* (Excel's wire
// form), except selectLockedCells / selectUnlockedCells which this guard
// doesn't police.

import type { Cell } from '@office-kit/xlsx/cell';
import { getCellProtection } from '@office-kit/xlsx/styles';
import type { Workbook } from '@office-kit/xlsx/workbook';
import type { SheetProtection, Worksheet } from '@office-kit/xlsx/worksheet';
import { getColumnDimension, getRowDimension } from '@office-kit/xlsx/worksheet';
import type { Range } from './address.ts';
import { forEachCellInRange } from './cells.ts';
import { EditRefusedError, type Part, type TransactionGuard } from './history.ts';

type SheetField = Extract<Part, { kind: 'sheet' }>['fields'][number];

/** Sheet fields a protected sheet still lets change, and the flag that unlocks the rest. */
const ALWAYS_ALLOWED: ReadonlySet<SheetField> = new Set<SheetField>(['sheetProtection', 'views']);
const FIELD_FLAG: Partial<Record<SheetField, keyof SheetProtection>> = {
  columnDimensions: 'formatColumns',
  rowDimensions: 'formatRows',
  drawing: 'objects',
  autoFilter: 'autoFilter',
  hyperlinks: 'insertHyperlinks',
  scenarios: 'scenarios',
};

/**
 * Fields whose only change in a structural step elsewhere (a row insert on
 * another sheet, a sheet rename) is their references being rewritten.
 */
const REFERENCE_FIELDS: ReadonlySet<SheetField> = new Set<SheetField>(['conditionalFormatting', 'dataValidations', 'drawing', 'pivotTables']);

function isProtected(ws: Worksheet): boolean {
  return ws.sheetProtection?.sheet === true;
}

/** The style a blank cell at (row, col) would show: its row's, else its column's, else the default. */
function blankProbe(ws: Worksheet, row: number, col: number): Cell {
  const styleId = getRowDimension(ws, row)?.style ?? getColumnDimension(ws, col)?.style ?? 0;
  return { row, col, value: null, styleId };
}

export function isCellLocked(wb: Workbook, ws: Worksheet, row: number, col: number): boolean {
  const cell = ws.rows.get(row)?.get(col) ?? blankProbe(ws, row, col);
  return getCellProtection(wb, cell).locked !== false;
}

/** True when some cell of `range` on a protected sheet is locked. Sparse: empty cells are judged by one probe. */
function rangeHasLockedCell(wb: Workbook, ws: Worksheet, range: Range): boolean {
  let stored = 0;
  let locked = false;
  forEachCellInRange(ws, range, (cell) => {
    stored++;
    if (!locked && getCellProtection(wb, cell).locked !== false) locked = true;
  });
  if (locked) return true;
  const area = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
  return stored < area && blankLocked(wb, ws, range);
}

function blankLocked(wb: Workbook, ws: Worksheet, range: Range): boolean {
  // Unlocking a whole column / row stores the style on its dimension, so a
  // blank cell is unlocked only if its row or column says so.
  // Probing the first row's columns (capped) keeps a whole-sheet range cheap.
  for (let c = range.c1; c <= Math.min(range.c2, range.c1 + 64); c++) {
    if (getCellProtection(wb, blankProbe(ws, range.r1, c)).locked !== false) return true;
  }
  return false;
}

export function protectionGuard(wb: () => Workbook): TransactionGuard {
  return (part, structural) => {
    switch (part.kind) {
      case 'cells':
        // A structural step only rewrites references here; what it moves is judged by its sheet fields.
        if (!structural && isProtected(part.ws) && rangeHasLockedCell(wb(), part.ws, part.range)) throw new EditRefusedError('protectedCell');
        return;
      case 'sheet': {
        const p = part.ws.sheetProtection;
        if (!p?.sheet) return;
        if (structural && !part.whole && part.fields.every((f) => REFERENCE_FIELDS.has(f))) return;
        for (const field of part.fields) {
          if (ALWAYS_ALLOWED.has(field)) continue;
          const flag = FIELD_FLAG[field];
          if (!flag || p[flag] !== false) throw new EditRefusedError('protectedCell');
        }
        return;
      }
      case 'title':
        if (wb().workbookProtection?.lockStructure) throw new EditRefusedError('protectedStructure');
        return;
      case 'workbook':
        if (wb().workbookProtection?.lockStructure && part.fields.includes('sheets')) throw new EditRefusedError('protectedStructure');
    }
  };
}
