// Applying formatting to ranges.
//
// A range can be whole columns, so formatting never walks every address. It
// restyles the cells that exist, records the format on the row/column
// dimension for whole-row/column selections (how Excel stores "format the
// entire column"), and materialises blank cells only for bounded ranges.
//
// Each distinct starting style maps to one resulting style, so a 10,000-cell
// range with three distinct styles costs three style registrations, not
// 10,000.

import type { Cell } from '@office-kit/xlsx/cell';
import { makeCell } from '@office-kit/xlsx/cell';
import type { Alignment, Border, Fill, FontPatch, Protection, Side } from '@office-kit/xlsx/styles';
import {
  getCellAlignment,
  getCellBorder,
  getCellProtection,
  makeBorder,
  patchCellFont,
  setCellAlignment,
  setCellBorder,
  setCellFill,
  setCellNumberFormat,
  setCellProtection,
} from '@office-kit/xlsx/styles';
import type { Workbook } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { setColumnDimension, setRowDimension, getColumnDimension, getRowDimension } from '@office-kit/xlsx/worksheet';
import type { Range } from './address.ts';
import { isWholeColumns, isWholeRows } from './address.ts';
import { forEachCellInRange } from './cells.ts';

/** Where a cell sits in the range being bordered; selects which sides a border preset sets. */
export interface EdgeInfo {
  readonly top: boolean;
  readonly bottom: boolean;
  readonly left: boolean;
  readonly right: boolean;
}

export interface StylePatch {
  readonly font?: FontPatch;
  /** `null` removes the fill. */
  readonly fill?: Fill | null;
  /** An `undefined` field removes that alignment attribute. */
  readonly alignment?: { readonly [K in keyof Alignment]?: Alignment[K] | undefined };
  readonly numFmt?: string;
  readonly protection?: Partial<Protection>;
  /** Border edit for a cell at the given position inside the range. */
  readonly border?: (current: Border, edge: EdgeInfo) => Border;
}

/** Largest bounded range whose blank cells are materialised to carry a style. */
const MATERIALISE_LIMIT = 200_000;

export function transformStyle(wb: Workbook, styleId: number, patch: StylePatch, edge: EdgeInfo): number {
  const scratch: Cell = makeCell(1, 1, null, styleId);
  if (patch.font) patchCellFont(wb, scratch, patch.font);
  if (patch.fill !== undefined) {
    setCellFill(wb, scratch, patch.fill ?? { kind: 'pattern', patternType: 'none' });
  }
  if (patch.alignment) {
    const merged: Record<string, unknown> = { ...getCellAlignment(wb, scratch) };
    for (const [k, v] of Object.entries(patch.alignment)) {
      if (v === undefined) delete merged[k];
      else merged[k] = v;
    }
    setCellAlignment(wb, scratch, merged as Alignment);
  }
  if (patch.numFmt !== undefined) setCellNumberFormat(wb, scratch, patch.numFmt);
  if (patch.protection) setCellProtection(wb, scratch, { ...getCellProtection(wb, scratch), ...patch.protection });
  if (patch.border) setCellBorder(wb, scratch, patch.border(getCellBorder(wb, scratch), edge));
  return scratch.styleId;
}

function edgeOf(range: Range, row: number, col: number): EdgeInfo {
  return { top: row === range.r1, bottom: row === range.r2, left: col === range.c1, right: col === range.c2 };
}

function edgeKey(e: EdgeInfo): number {
  return (e.top ? 1 : 0) | (e.bottom ? 2 : 0) | (e.left ? 4 : 0) | (e.right ? 8 : 0);
}

/**
 * Apply `patch` across `range`. The caller has declared the range's cells
 * (and, for whole rows/columns, the dimension maps) on the transaction.
 */
export function applyStyle(wb: Workbook, ws: Worksheet, range: Range, patch: StylePatch): void {
  const memo = new Map<string, number>();
  const restyle = (styleId: number, edge: EdgeInfo): number => {
    const key = `${styleId}:${patch.border ? edgeKey(edge) : 0}`;
    let next = memo.get(key);
    if (next === undefined) {
      next = transformStyle(wb, styleId, patch, edge);
      memo.set(key, next);
    }
    return next;
  };

  const wholeCols = isWholeColumns(range);
  const wholeRows = isWholeRows(range);
  if (wholeCols || wholeRows) {
    // Record the format on the dimensions so blank cells in them render and
    // new entries inherit it, then restyle the cells that already exist.
    if (wholeCols) {
      for (let c = range.c1; c <= range.c2; c++) {
        const current = getColumnDimension(ws, c)?.style ?? 0;
        setColumnDimension(ws, c, { style: restyle(current, edgeOf(range, 2, c)) });
      }
    }
    if (wholeRows && !wholeCols) {
      for (let r = range.r1; r <= range.r2; r++) {
        const current = getRowDimension(ws, r)?.style ?? 0;
        setRowDimension(ws, r, { style: restyle(current, edgeOf(range, r, 2)) });
      }
    }
    forEachCellInRange(ws, range, (cell) => {
      cell.styleId = restyle(cell.styleId, edgeOf(range, cell.row, cell.col));
    });
    return;
  }

  const area = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
  if (area > MATERIALISE_LIMIT) {
    forEachCellInRange(ws, range, (cell) => {
      cell.styleId = restyle(cell.styleId, edgeOf(range, cell.row, cell.col));
    });
    return;
  }
  for (let r = range.r1; r <= range.r2; r++) {
    let rowMap = ws.rows.get(r);
    const rowStyle = getRowDimension(ws, r)?.style;
    for (let c = range.c1; c <= range.c2; c++) {
      const edge = edgeOf(range, r, c);
      const existing = rowMap?.get(c);
      if (existing) {
        existing.styleId = restyle(existing.styleId, edge);
        continue;
      }
      const base = rowStyle ?? getColumnDimension(ws, c)?.style ?? 0;
      const styleId = restyle(base, edge);
      if (styleId === 0) continue;
      if (!rowMap) {
        rowMap = new Map();
        ws.rows.set(r, rowMap);
      }
      rowMap.set(c, makeCell(r, c, null, styleId));
    }
  }
}

// ---- border presets (Home ▸ Font ▸ Borders menu) ---------------------------

export type BorderPreset =
  | 'bottom'
  | 'top'
  | 'left'
  | 'right'
  | 'none'
  | 'all'
  | 'outside'
  | 'thickOutside'
  | 'bottomDouble'
  | 'thickBottom'
  | 'topBottom'
  | 'topThickBottom'
  | 'topDoubleBottom'
  | 'insideHorizontal'
  | 'insideVertical'
  | 'inside';

export function borderPatch(preset: BorderPreset, side: Side): NonNullable<StylePatch['border']> {
  const thick: Side = { ...side, style: 'thick' };
  const dbl: Side = { ...side, style: 'double' };
  return (cur, e) => {
    const b: { -readonly [K in keyof Border]: Border[K] } = { ...cur };
    const set = (k: 'top' | 'bottom' | 'left' | 'right', s: Side | undefined) => {
      if (s) b[k] = s;
      else delete b[k];
    };
    switch (preset) {
      case 'none':
        return makeBorder({});
      case 'all':
        set('top', side);
        set('bottom', side);
        set('left', side);
        set('right', side);
        break;
      case 'outside':
      case 'thickOutside': {
        const s = preset === 'outside' ? side : thick;
        if (e.top) set('top', s);
        if (e.bottom) set('bottom', s);
        if (e.left) set('left', s);
        if (e.right) set('right', s);
        break;
      }
      case 'bottom':
        if (e.bottom) set('bottom', side);
        break;
      case 'top':
        if (e.top) set('top', side);
        break;
      case 'left':
        if (e.left) set('left', side);
        break;
      case 'right':
        if (e.right) set('right', side);
        break;
      case 'bottomDouble':
        if (e.bottom) set('bottom', dbl);
        break;
      case 'thickBottom':
        if (e.bottom) set('bottom', thick);
        break;
      case 'topBottom':
        if (e.top) set('top', side);
        if (e.bottom) set('bottom', side);
        break;
      case 'topThickBottom':
        if (e.top) set('top', side);
        if (e.bottom) set('bottom', thick);
        break;
      case 'topDoubleBottom':
        if (e.top) set('top', side);
        if (e.bottom) set('bottom', dbl);
        break;
      case 'insideHorizontal':
        if (!e.top) set('top', side);
        if (!e.bottom) set('bottom', side);
        break;
      case 'insideVertical':
        if (!e.left) set('left', side);
        if (!e.right) set('right', side);
        break;
      case 'inside':
        if (!e.top) set('top', side);
        if (!e.bottom) set('bottom', side);
        if (!e.left) set('left', side);
        if (!e.right) set('right', side);
        break;
    }
    return makeBorder(b);
  };
}
