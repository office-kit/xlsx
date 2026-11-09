// Formula rewriting for whole-row insertion and deletion. Unlike the shared
// formula translator next door (copy semantics: relative parts shift, `$`
// anchors stay), a structural edit moves the cells themselves, so every
// reference to a moved cell follows it whatever its anchors, a range
// straddling the edit grows or shrinks, and a reference whose every cell was
// deleted becomes `#REF!` — Excel's rules.

import { columnIndexFromLetter, MAX_COL, MAX_ROW } from '../utils/coordinate.js';
import { OPERAND, RANGE, TokenizerError, tokenize, WSPACE } from './tokenizer.js';

const REF_ERROR = '#REF!';

export interface RowEdit {
  /** Title of the sheet whose rows move. */
  readonly sheet: string;
  /** First row affected, 1-based. */
  readonly at: number;
  /** > 0 inserts `count` rows before `at`; < 0 deletes `-count` rows starting at `at`. */
  readonly count: number;
  /** The edited sheet's tables, for {@link adjustArea}. */
  readonly tables?: ReadonlyArray<TableBand>;
}

/** Where a table's rows sit, 1-based. */
export interface TableBand {
  readonly header?: number;
  readonly firstData: number;
  readonly lastData: number;
  readonly totals?: number;
  readonly minCol: number;
  readonly maxCol: number;
}

/**
 * Move the row span [lo, hi] through `edit`. Returns `undefined` when every
 * row of the span was deleted, or pushed off the bottom of the grid.
 */
export function adjustRowSpan(lo: number, hi: number, edit: RowEdit): [number, number] | undefined {
  const { at, count } = edit;
  if (count > 0) {
    const newLo = lo >= at ? lo + count : lo;
    if (newLo > MAX_ROW) return undefined;
    // A span already reaching the last row stays pinned there (`A5:A1048576`).
    return [newLo, hi >= at ? Math.min(hi + count, MAX_ROW) : hi];
  }
  const last = at - count - 1;
  if (lo >= at && hi <= last) return undefined;
  const newLo = lo < at ? lo : lo > last ? lo + count : at;
  // Pinned to the last row on deletion too, as Excel keeps `A5:A1048576`.
  const newHi = hi === MAX_ROW ? MAX_ROW : hi < at ? hi : hi > last ? hi + count : at - 1;
  return [newLo, newHi];
}

/**
 * {@link adjustRowSpan} for an area within columns [minCol, maxCol]. Excel
 * reads an area covering a table's whole data body (with or without its
 * header or totals row) as a reference to the table's columns, so rows
 * inserted anywhere into the body — at the first data row or at the totals
 * row included, where a plain range would shift or stay — widen it.
 */
export function adjustArea(lo: number, hi: number, minCol: number, maxCol: number, edit: RowEdit): [number, number] | undefined {
  if (edit.count > 0) {
    for (const t of edit.tables ?? []) {
      if (minCol < t.minCol || maxCol > t.maxCol) continue;
      const atStart = edit.at === t.firstData && t.header !== undefined;
      const atTotals = edit.at === t.totals;
      if (!atStart && !atTotals) continue;
      const fromTop = lo === t.firstData || (lo === t.header && t.header !== undefined);
      const toBottom = hi === t.lastData || (hi === t.totals && atStart);
      if (fromTop && toBottom) return [lo, Math.min(hi + edit.count, MAX_ROW)];
    }
  }
  return adjustRowSpan(lo, hi, edit);
}

// `Sheet!`, `'My sheet'!`. A `:` or `[` in the prefix is a 3-D or an external
// reference, which a row edit on one sheet of this workbook never touches.
const PREFIX_RE = /^(?:'((?:[^']|'')+)'|([^'![\]:]+))!/;
const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;
const ROWS_RE = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})$/;

interface RowPart {
  readonly abs: string;
  readonly row: number;
}

interface ParsedArea {
  /** Column text preceding each row part (`$B`), empty for a row range. */
  readonly cols: readonly [string, string];
  readonly rows: readonly [RowPart, RowPart];
  readonly single: boolean;
}

const onGrid = (row: number): boolean => row >= 1 && row <= MAX_ROW;

// Compared as text so a past-the-grid column never reaches the index parser.
const LAST_COL = 'XFD';
const colOnGrid = (letters: string): boolean => letters.length < LAST_COL.length || letters <= LAST_COL;
const colIndex = (text: string): number => columnIndexFromLetter(text.replace('$', '').toUpperCase());

function parseCell(text: string): { col: string; part: RowPart } | undefined {
  const m = CELL_RE.exec(text);
  if (m === null) return undefined;
  const col = (m[2] as string).toUpperCase();
  const row = Number(m[4]);
  // `XFE1` or `A2000000` read as a name to Excel, not a reference.
  if (!colOnGrid(col) || !onGrid(row)) return undefined;
  return { col: `${m[1]}${m[2]}`, part: { abs: m[3] as string, row } };
}

function parseArea(text: string): ParsedArea | undefined {
  const rows = ROWS_RE.exec(text);
  if (rows !== null) {
    const r1 = Number(rows[2]);
    const r2 = Number(rows[4]);
    if (!onGrid(r1) || !onGrid(r2)) return undefined;
    return {
      cols: ['', ''],
      rows: [
        { abs: rows[1] as string, row: r1 },
        { abs: rows[3] as string, row: r2 },
      ],
      single: false,
    };
  }
  const colon = text.indexOf(':');
  if (colon === -1) {
    const cell = parseCell(text);
    return cell === undefined ? undefined : { cols: [cell.col, cell.col], rows: [cell.part, cell.part], single: true };
  }
  const a = parseCell(text.slice(0, colon));
  const b = parseCell(text.slice(colon + 1));
  if (a === undefined || b === undefined) return undefined;
  return { cols: [a.col, b.col], rows: [a.part, b.part], single: false };
}

/**
 * Rewrite one reference operand (`A1`, `Sheet!$A$1:$B$9`, `3:5`) written on
 * `hostSheet`. Returns the operand unchanged when it does not point at
 * `edit.sheet`, or is not a plain reference (a name, a whole-column range, a
 * 3-D or external reference, a structured reference).
 */
function adjustReferenceForRowEdit(operand: string, hostSheet: string | undefined, edit: RowEdit): string {
  const prefixMatch = PREFIX_RE.exec(operand);
  let sheet: string | undefined;
  let prefix = '';
  if (prefixMatch !== null) {
    prefix = prefixMatch[0];
    sheet = prefixMatch[1] !== undefined ? prefixMatch[1].replaceAll("''", "'") : prefixMatch[2];
  } else {
    sheet = hostSheet;
  }
  if (sheet === undefined || sheet.toLowerCase() !== edit.sheet.toLowerCase()) return operand;
  const area = parseArea(operand.slice(prefix.length));
  if (area === undefined) return operand;
  const [p1, p2] = area.rows;
  const flip = p1.row > p2.row;
  const [c1, c2] = area.cols[0] !== '' ? [colIndex(area.cols[0]), colIndex(area.cols[1])] : [1, MAX_COL];
  const span = adjustArea(Math.min(p1.row, p2.row), Math.max(p1.row, p2.row), Math.min(c1, c2), Math.max(c1, c2), edit);
  if (span === undefined) return `${prefix}${REF_ERROR}`;
  const [lo, hi] = span;
  const r1 = flip ? hi : lo;
  const r2 = flip ? lo : hi;
  if (area.single) return `${prefix}${area.cols[0]}${p1.abs}${r1}`;
  return `${prefix}${area.cols[0]}${p1.abs}${r1}:${area.cols[1]}${p2.abs}${r2}`;
}

/**
 * Rewrite every reference in `formula` (stored without its leading `=`, as
 * `<f>` holds it) that points at rows of `edit.sheet`. `hostSheet` is the
 * sheet unprefixed references resolve against; `undefined` for a
 * workbook-scope defined name, whose references are always qualified.
 * Returns the input string itself when nothing changed.
 */
export function adjustFormulaForRowEdit(formula: string, hostSheet: string | undefined, edit: RowEdit): string {
  // Only a reference to the edited sheet can change. One that names it
  // carries its title (quotes doubled inside a quoted name); an unprefixed
  // one needs the formula to live there.
  const onEditedSheet = hostSheet !== undefined && hostSheet.toLowerCase() === edit.sheet.toLowerCase();
  const lower = formula.toLowerCase();
  const title = edit.sheet.toLowerCase();
  if (!onEditedSheet && !lower.includes(title) && !lower.includes(title.replaceAll("'", "''"))) return formula;
  // The tokenizer folds runs of whitespace, so the rewrite splices the
  // original text at each operand rather than rendering the tokens back.
  let out = '';
  let cursor = 0;
  for (const t of tokenize(`=${formula}`)) {
    if (t.type === WSPACE) continue;
    const start = formula.indexOf(t.value, cursor);
    if (start === -1) throw new TokenizerError(`cannot locate token "${t.value}" in formula "${formula}"`);
    const end = start + t.value.length;
    if (t.type === OPERAND && t.subtype === RANGE) {
      const next = adjustReferenceForRowEdit(t.value, hostSheet, edit);
      if (next !== t.value) {
        out += formula.slice(cursor, start) + next;
        cursor = end;
        continue;
      }
    }
    out += formula.slice(cursor, end);
    cursor = end;
  }
  out += formula.slice(cursor);
  return out === formula ? formula : out;
}

/**
 * The relative (un-`$`) row parts of `formula`'s references to `sheet`, read
 * from `hostSheet` as in {@link adjustFormulaForRowEdit}, each with the
 * column its reference starts at (`undefined` for a row range). They are
 * what a copy of the formula moves, so where they land against an edit
 * decides whether copies at different rows still agree.
 */
export function relativeRows(formula: string, hostSheet: string, sheet: string): Array<{ row: number; col: number | undefined }> {
  const out: Array<{ row: number; col: number | undefined }> = [];
  for (const t of tokenize(`=${formula}`)) {
    if (t.type !== OPERAND || t.subtype !== RANGE) continue;
    const prefixMatch = PREFIX_RE.exec(t.value);
    const named = prefixMatch === null ? hostSheet : (prefixMatch[1]?.replaceAll("''", "'") ?? prefixMatch[2]);
    if (named === undefined || named.toLowerCase() !== sheet.toLowerCase()) continue;
    const area = parseArea(t.value.slice(prefixMatch?.[0].length ?? 0));
    if (area === undefined) continue;
    const col = area.cols[0] === '' ? undefined : colIndex(area.cols[0]);
    for (const part of area.rows) if (part.abs === '') out.push({ row: part.row, col });
  }
  return out;
}
