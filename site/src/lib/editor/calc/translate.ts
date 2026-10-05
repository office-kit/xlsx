// Formula text rewriting for copy / paste / fill, row and column insertion or
// deletion, and sheet renames. Each one tokenizes the formula and splices only
// the reference tokens it changes, so the user's spacing and casing survive.

import type { RefArea, SheetPrefix } from './ast.ts';
import { isValidCol, isValidRow, renderArea, renderPrefix } from './address.ts';
import { type Token, tokenize } from './lexer.ts';
import { MAX_COL, MAX_ROW } from './types.ts';

const REF_ERROR = '#REF!';

type Replacement = { start: number; end: number; text: string };

const splice = (text: string, replacements: readonly Replacement[]): string => {
  let out = '';
  let at = 0;
  for (const r of replacements) {
    out += text.slice(at, r.start) + r.text;
    at = r.end;
  }
  return out + text.slice(at);
};

type AreaRefToken = Extract<Token, { kind: 'ref' }> & { area: RefArea };

/** Rewrite every reference token; `rewrite` returns the new area, `null` for #REF!, or `undefined` to keep it. */
const rewriteAreas = (text: string, rewrite: (t: AreaRefToken) => RefArea | null | undefined): string => {
  const replacements: Replacement[] = [];
  for (const t of tokenize(text, true)) {
    if (t.kind !== 'ref' || t.area === undefined) continue;
    const next = rewrite({ ...t, area: t.area });
    if (next === undefined) continue;
    replacements.push({ start: t.prefixEnd, end: t.end, text: next === null ? REF_ERROR : renderArea(next) });
  }
  return splice(text, replacements);
};

/**
 * Shift the relative parts of every reference by (dRow, dCol), the way a
 * copy / paste or fill-handle drag does. Absolute (`$`) parts stay put; a
 * reference pushed off the grid becomes `#REF!`.
 */
export function translateFormula(text: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return text;
  return rewriteAreas(text, ({ area: a }) => {
    const r1 = a.kind === 'cols' || a.r1Abs ? a.r1 : a.r1 + dRow;
    const r2 = a.kind === 'cols' || a.r2Abs ? a.r2 : a.r2 + dRow;
    const c1 = a.kind === 'rows' || a.c1Abs ? a.c1 : a.c1 + dCol;
    const c2 = a.kind === 'rows' || a.c2Abs ? a.c2 : a.c2 + dCol;
    if (!isValidRow(r1) || !isValidRow(r2) || !isValidCol(c1) || !isValidCol(c2)) return null;
    return { ...a, r1, r2, c1, c2 };
  });
}

export interface StructureEdit {
  readonly sheet: string;
  readonly axis: 'row' | 'col';
  /** First row / column affected (1-based). */
  readonly at: number;
  /** > 0 inserts `count` before `at`; < 0 deletes `-count` starting at `at`. */
  readonly count: number;
  /**
   * Insert / Delete Cells with "shift cells down / up" (axis `row`) or "right /
   * left" (axis `col`): the 1-based span on the *other* axis that moves.
   * Omitted means whole rows / columns.
   */
  readonly band?: { readonly from: number; readonly to: number };
}

/**
 * Move one axis span [lo, hi] through an insertion or deletion. Returns null
 * when every cell of the span was deleted. A span straddling an insertion
 * grows; one straddling a deletion shrinks — Excel's rules.
 */
const adjustSpan = (lo: number, hi: number, at: number, count: number, max: number): [number, number] | null => {
  if (count > 0) {
    const newLo = lo >= at ? lo + count : lo;
    // A span already reaching the grid edge stays pinned there (`A1:A1048576`).
    const newHi = hi >= at ? Math.min(hi + count, max) : hi;
    if (newLo > max) return null;
    return [newLo, newHi];
  }
  const removed = -count;
  const last = at + removed - 1;
  if (lo >= at && hi <= last) return null;
  const newLo = lo < at ? lo : lo > last ? lo - removed : at;
  const newHi = hi < at ? hi : hi > last ? hi - removed : at - 1;
  return [newLo, newHi];
};

/**
 * Rewrite a formula living on `formulaSheet` after rows or columns were
 * inserted into or deleted from `edit.sheet`. References to deleted cells
 * become `#REF!`; ranges grow and shrink like Excel's. Whole-column
 * references ignore row edits and whole-row references ignore column edits.
 */
export function adjustFormulaForStructure(text: string, formulaSheet: string, edit: StructureEdit): string {
  if (edit.count === 0) return text;
  const target = edit.sheet.toLowerCase();
  return rewriteAreas(text, ({ prefix, area: a }) => {
    if (prefix?.sheet2 !== undefined || prefix?.external !== undefined) return undefined;
    if ((prefix?.sheet ?? formulaSheet).toLowerCase() !== target) return undefined;
    // With a band only references lying wholly inside it move; Excel leaves
    // ones that straddle the band's edge alone.
    if (edit.band !== undefined) {
      const lo = edit.axis === 'row' ? Math.min(a.c1, a.c2) : Math.min(a.r1, a.r2);
      const hi = edit.axis === 'row' ? Math.max(a.c1, a.c2) : Math.max(a.r1, a.r2);
      if (lo < edit.band.from || hi > edit.band.to) return undefined;
    }
    if (edit.axis === 'row') {
      if (a.kind === 'cols') return undefined;
      const span = adjustSpan(Math.min(a.r1, a.r2), Math.max(a.r1, a.r2), edit.at, edit.count, MAX_ROW);
      if (span === null) return null;
      const [r1, r2] = span;
      const flip = a.r1 > a.r2;
      return { ...a, r1: flip ? r2 : r1, r2: flip ? r1 : r2 };
    }
    if (a.kind === 'rows') return undefined;
    const span = adjustSpan(Math.min(a.c1, a.c2), Math.max(a.c1, a.c2), edit.at, edit.count, MAX_COL);
    if (span === null) return null;
    const [c1, c2] = span;
    const flip = a.c1 > a.c2;
    return { ...a, c1: flip ? c2 : c1, c2: flip ? c1 : c2 };
  });
}

/** Point every `oldName!` prefix at `newName`, adding or dropping quotes as the new name needs. */
export function renameSheetInFormula(text: string, oldName: string, newName: string): string {
  const old = oldName.toLowerCase();
  const replacements: Replacement[] = [];
  for (const t of tokenize(text, true)) {
    if ((t.kind !== 'ref' && t.kind !== 'name') || t.prefix === undefined || t.prefix.external !== undefined) continue;
    const p = t.prefix;
    const hitFirst = p.sheet.toLowerCase() === old;
    const hitSecond = p.sheet2?.toLowerCase() === old;
    if (!hitFirst && !hitSecond) continue;
    const renamed: SheetPrefix = {
      sheet: hitFirst ? newName : p.sheet,
      ...(p.sheet2 !== undefined ? { sheet2: hitSecond ? newName : p.sheet2 } : {}),
    };
    replacements.push({ start: t.start, end: t.prefixEnd, text: renderPrefix(renamed) });
  }
  return splice(text, replacements);
}

/**
 * Rewrite a formula after sheet `deleted` was removed from `order` (the sheet
 * titles before the deletion). A reference to that sheet becomes #REF!, and a
 * 3-D reference ending on it shrinks to the next sheet inward, as in Excel.
 */
export function deleteSheetInFormula(text: string, deleted: string, order: readonly string[]): string {
  const gone = deleted.toLowerCase();
  const indexOf = (title: string) => order.findIndex((o) => o.toLowerCase() === title.toLowerCase());
  const replacements: Replacement[] = [];
  for (const t of tokenize(text, true)) {
    if ((t.kind !== 'ref' && t.kind !== 'name') || t.prefix === undefined || t.prefix.external !== undefined) continue;
    const p = t.prefix;
    const hitFirst = p.sheet.toLowerCase() === gone;
    const hitSecond = p.sheet2?.toLowerCase() === gone;
    if (!hitFirst && !hitSecond) continue;
    if (p.sheet2 === undefined) {
      replacements.push({ start: t.start, end: t.end, text: REF_ERROR });
      continue;
    }
    const a = indexOf(p.sheet);
    const b = indexOf(p.sheet2);
    if (a < 0 || b < 0) continue;
    const step = a < b ? 1 : -1;
    const first = hitFirst ? a + step : a;
    const last = hitSecond ? b - step : b;
    const sheet = order[first] ?? p.sheet;
    const sheet2 = order[last] ?? p.sheet2;
    replacements.push({ start: t.start, end: t.prefixEnd, text: renderPrefix(first === last ? { sheet } : { sheet, sheet2 }) });
  }
  return splice(text, replacements);
}

export interface MoveEdit {
  /** Sheet the block was cut from. */
  readonly sheet: string;
  readonly range: { readonly r1: number; readonly c1: number; readonly r2: number; readonly c2: number };
  /** Sheet the block was pasted onto (may equal `sheet`). */
  readonly toSheet: string;
  readonly dRow: number;
  readonly dCol: number;
}

interface Rect {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

const contains = (outer: Rect, inner: Rect): boolean =>
  inner.r1 >= outer.r1 && inner.r2 <= outer.r2 && inner.c1 >= outer.c1 && inner.c2 <= outer.c2;

/**
 * Rewrite a formula on `formulaSheet` after a cut / paste (or drag-move) of
 * `move.range`. References lying wholly inside the moved block follow it,
 * anchors and all; references wholly inside the cells the block overwrote
 * become `#REF!`; anything straddling either edge stays as written — Excel's
 * rules. Apply it to every formula in the workbook, the moved ones included
 * (their relative references are not shifted, unlike a copy).
 */
export function adjustFormulaForMove(text: string, formulaSheet: string, move: MoveEdit): string {
  const source: Rect = {
    r1: Math.min(move.range.r1, move.range.r2),
    r2: Math.max(move.range.r1, move.range.r2),
    c1: Math.min(move.range.c1, move.range.c2),
    c2: Math.max(move.range.c1, move.range.c2),
  };
  const dest: Rect = { r1: source.r1 + move.dRow, r2: source.r2 + move.dRow, c1: source.c1 + move.dCol, c2: source.c2 + move.dCol };
  const from = move.sheet.toLowerCase();
  const to = move.toSheet.toLowerCase();
  const replacements: Replacement[] = [];
  for (const t of tokenize(text, true)) {
    if (t.kind !== 'ref' || t.area === undefined) continue;
    const p = t.prefix;
    if (p?.sheet2 !== undefined || p?.external !== undefined) continue;
    const sheet = (p?.sheet ?? formulaSheet).toLowerCase();
    const a = t.area;
    const rect: Rect = { r1: Math.min(a.r1, a.r2), r2: Math.max(a.r1, a.r2), c1: Math.min(a.c1, a.c2), c2: Math.max(a.c1, a.c2) };
    if (sheet === from && contains(source, rect)) {
      const moved: RefArea = {
        ...a,
        r1: a.kind === 'cols' ? a.r1 : a.r1 + move.dRow,
        r2: a.kind === 'cols' ? a.r2 : a.r2 + move.dRow,
        c1: a.kind === 'rows' ? a.c1 : a.c1 + move.dCol,
        c2: a.kind === 'rows' ? a.c2 : a.c2 + move.dCol,
      };
      const crossSheet = move.toSheet.toLowerCase() !== formulaSheet.toLowerCase();
      const prefix = p !== undefined || crossSheet ? renderPrefix({ sheet: move.toSheet }) : '';
      replacements.push({ start: t.start, end: t.end, text: prefix + renderArea(moved) });
    } else if (sheet === to && contains(dest, rect)) {
      replacements.push({ start: t.prefixEnd, end: t.end, text: REF_ERROR });
    }
  }
  return splice(text, replacements);
}
