// Data ▸ Text to Columns: split each cell of a one-column selection into
// several cells, by delimiters or at fixed character positions, the way
// Excel's Convert Text to Columns Wizard does. Pieces are entered as if typed,
// so "12" becomes a number and "1/2/2024" a date.

import { makeCell } from '@office-kit/xlsx/cell';
import { getCellDisplayText } from '@office-kit/xlsx/styles';
import type { CellPos, Range } from './address.ts';
import { MAX_COL } from './address.ts';
import { getCellAt } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';
import { parseInput } from './input.ts';

export interface DelimitedOptions {
  readonly kind: 'delimited';
  readonly tab: boolean;
  readonly semicolon: boolean;
  readonly comma: boolean;
  readonly space: boolean;
  /** Any one extra delimiter character ('' for none). */
  readonly other: string;
  readonly consecutiveAsOne: boolean;
  /** Text qualifier: delimiters inside a qualified run are literal; a doubled qualifier is one literal. */
  readonly qualifier: '"' | "'" | null;
}

export interface FixedWidthOptions {
  readonly kind: 'fixed';
  /** Character positions a new column starts at (break lines in the wizard's ruler). */
  readonly breaks: readonly number[];
}

export type SplitOptions = DelimitedOptions | FixedWidthOptions;

function delimiterSet(opts: DelimitedOptions): Set<string> {
  const set = new Set<string>();
  if (opts.tab) set.add('\t');
  if (opts.semicolon) set.add(';');
  if (opts.comma) set.add(',');
  if (opts.space) set.add(' ');
  const other = opts.other.charAt(0);
  if (other) set.add(other);
  return set;
}

function splitDelimited(text: string, opts: DelimitedOptions): string[] {
  const delims = delimiterSet(opts);
  if (delims.size === 0) return [text];
  const out: string[] = [];
  let field = '';
  let quoted = false;
  let afterDelimiter = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charAt(i);
    if (opts.qualifier !== null && ch === opts.qualifier) {
      if (quoted && text.charAt(i + 1) === opts.qualifier) {
        field += ch;
        i++;
      } else {
        quoted = !quoted;
      }
      afterDelimiter = false;
      continue;
    }
    if (!quoted && delims.has(ch)) {
      if (opts.consecutiveAsOne && afterDelimiter) continue;
      out.push(field);
      field = '';
      afterDelimiter = true;
      continue;
    }
    field += ch;
    afterDelimiter = false;
  }
  out.push(field);
  return out;
}

function splitFixed(text: string, breaks: readonly number[]): string[] {
  const cuts = [...new Set(breaks)].filter((b) => b > 0 && b < text.length).sort((a, b) => a - b);
  const out: string[] = [];
  let start = 0;
  for (const cut of cuts) {
    out.push(text.slice(start, cut));
    start = cut;
  }
  out.push(text.slice(start));
  return out;
}

export function splitText(text: string, opts: SplitOptions): string[] {
  return opts.kind === 'delimited' ? splitDelimited(text, opts) : splitFixed(text, opts.breaks);
}

/** Parse the wizard's break-line field: "5, 12" → [5, 12]; undefined when malformed. */
export function parseBreaks(input: string): number[] | undefined {
  const parts = input.split(/[\s,]+/).filter((p) => p !== '');
  const out: number[] = [];
  for (const p of parts) {
    if (!/^\d+$/.test(p)) return undefined;
    const n = Number(p);
    if (n <= 0) return undefined;
    out.push(n);
  }
  return [...new Set(out)].sort((a, b) => a - b);
}

/**
 * Split the first column of `range` into cells starting at `dest` (same row
 * offsets). Returns the number of destination columns written.
 */
export function textToColumns(ctl: EditorController, range: Range, opts: SplitOptions, dest: CellPos): number {
  const doc = ctl.doc;
  const ws = doc.ws;
  const lines: Array<{ row: number; parts: string[] }> = [];
  let width = 0;
  for (let r = range.r1; r <= range.r2; r++) {
    const cell = getCellAt(ws, r, range.c1);
    if (!cell || cell.value === null) continue;
    const parts = splitText(getCellDisplayText(doc.wb, cell), opts);
    width = Math.max(width, parts.length);
    lines.push({ row: r, parts });
  }
  if (lines.length === 0) return 0;
  const c2 = Math.min(MAX_COL, dest.col + width - 1);
  const parseOpts = { dateOrder: ctl.dateOrder(), date1904: doc.wb.date1904 };
  doc.transact('Text to Columns', (tx) => {
    const top = dest.row;
    const bottom = dest.row + (range.r2 - range.r1);
    tx.cells(ws, { r1: top, c1: dest.col, r2: bottom, c2 });
    for (const { row, parts } of lines) {
      const targetRow = dest.row + (row - range.r1);
      parts.forEach((part, i) => {
        const col = dest.col + i;
        if (col > MAX_COL) return;
        const { value } = parseInput(part, parseOpts);
        let rowMap = ws.rows.get(targetRow);
        const existing = rowMap?.get(col);
        if (existing) {
          existing.value = value;
          return;
        }
        if (value === null) return;
        if (!rowMap) {
          rowMap = new Map();
          ws.rows.set(targetRow, rowMap);
        }
        rowMap.set(col, makeCell(targetRow, col, value, ctl.defaultStyleAt(targetRow, col)));
      });
    }
  });
  return c2 - dest.col + 1;
}
