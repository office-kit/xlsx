// Data ▸ Advanced (Advanced Filter): filter a list by a criteria range — one
// condition row per alternative, the conditions of a row all applying — in
// place, or copy the matching records to another location.

import { getCellDisplayText } from '@office-kit/xlsx/styles';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import type { CalcScalar } from '../calc/index.ts';
import type { MessageKey } from '../i18n/i18n.svelte.ts';
import { rangesIntersect, type Range } from './address.ts';
import { getCellAt, isBlank, setValueAt } from './cells.ts';
import type { EditorController } from './controller.svelte.ts';

type Test = (value: CalcScalar) => boolean;

const OP_RE = /^(<=|>=|<>|=|<|>)(.*)$/s;

/** Excel wildcards (* ? and ~ as escape) as an anchored, case-insensitive pattern. */
function wildcard(pattern: string, prefix: boolean): RegExp {
  let src = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern.charAt(i);
    if (ch === '~' && i + 1 < pattern.length) {
      src += pattern.charAt(++i).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    } else if (ch === '*') src += '.*';
    else if (ch === '?') src += '.';
    else src += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${src}${prefix ? '' : '$'}`, 'is');
}

function asText(v: CalcScalar): string {
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'object') return v.code;
  return String(v);
}

/**
 * One criteria cell as a test. Plain text matches values that begin with it,
 * a number or boolean matches equal values, and a leading comparison operator
 * compares (`=` alone matches blanks, `<>` alone non-blanks).
 */
export function criterionTest(criterion: CalcScalar): Test | undefined {
  if (criterion === null || criterion === '') return undefined;
  if (typeof criterion === 'number' || typeof criterion === 'boolean') return (v) => v === criterion;
  if (typeof criterion === 'object') return (v) => typeof v === 'object' && v !== null && v.code === criterion.code;
  const m = OP_RE.exec(criterion);
  if (!m) {
    const re = wildcard(criterion, true);
    return (v) => typeof v === 'string' && re.test(v);
  }
  const op = m[1] ?? '=';
  const operand = m[2] ?? '';
  if (operand === '') return op === '=' ? (v) => v === null || v === '' : op === '<>' ? (v) => v !== null && v !== '' : undefined;
  const num = Number(operand);
  const numeric = operand.trim() !== '' && Number.isFinite(num);
  const upper = operand.toUpperCase();
  const bool = upper === 'TRUE' ? true : upper === 'FALSE' ? false : undefined;
  const compare = (v: CalcScalar): number | undefined => {
    if (numeric) return typeof v === 'number' ? v - num : undefined;
    if (bool !== undefined) return typeof v === 'boolean' ? Number(v) - Number(bool) : undefined;
    if (typeof v !== 'string') return undefined;
    return v.localeCompare(operand, undefined, { sensitivity: 'base' });
  };
  if (op === '=' || op === '<>') {
    const re = numeric || bool !== undefined ? undefined : wildcard(operand, false);
    const equal: Test = (v) => (re ? typeof v === 'string' && re.test(v) : compare(v) === 0);
    return op === '=' ? equal : (v) => !equal(v);
  }
  return (v) => {
    const d = compare(v);
    if (d === undefined) return false;
    return op === '<' ? d < 0 : op === '<=' ? d <= 0 : op === '>' ? d > 0 : d >= 0;
  };
}

/** Rows of a criteria range as AND-lists of (list column index, test); undefined when a label is unknown. */
export function compileCriteria(listHeaders: readonly string[], criteria: ReadonlyArray<ReadonlyArray<CalcScalar>>): Array<Array<[number, Test]>> | undefined {
  const [header = [], ...rows] = criteria;
  const lower = listHeaders.map((h) => h.trim().toLowerCase());
  const columns = header.map((h) => lower.indexOf(asText(h).trim().toLowerCase()));
  if (columns.some((c, i) => c < 0 && asText(header[i] ?? null).trim() !== '')) return undefined;
  return rows.map((row) =>
    row.flatMap((value, i): Array<[number, Test]> => {
      const col = columns[i];
      const test = criterionTest(value);
      return col !== undefined && col >= 0 && test ? [[col, test]] : [];
    }),
  );
}

/** Whether a record passes: any criteria row whose tests all pass (an empty row passes everything). */
export function matches(record: readonly CalcScalar[], compiled: ReadonlyArray<ReadonlyArray<[number, Test]>>): boolean {
  if (compiled.length === 0) return true;
  return compiled.some((row) => row.every(([col, test]) => test(record[col] ?? null)));
}

export interface AdvancedFilterOptions {
  readonly list: Range;
  readonly criteria: Range | undefined;
  /** Top-left (or label row) of the copy destination; undefined filters in place. */
  readonly copyTo: Range | undefined;
  readonly unique: boolean;
}

/** List ranges filtered in place, per sheet, so Data ▸ Clear can show their rows again. */
const inPlace = new WeakMap<Worksheet, Range>();

export function hasAdvancedFilter(ws: Worksheet): boolean {
  return inPlace.has(ws);
}

function values(ctl: EditorController, r: Range): CalcScalar[][] {
  const sheet = ctl.doc.ws.title;
  const out: CalcScalar[][] = [];
  for (let row = r.r1; row <= r.r2; row++) {
    const line: CalcScalar[] = [];
    for (let col = r.c1; col <= r.c2; col++) line.push(ctl.doc.calc.cellValue(sheet, row, col));
    out.push(line);
  }
  return out;
}

function setRowHidden(ws: Worksheet, row: number, hidden: boolean): void {
  const dim = ws.rowDimensions.get(row);
  if (hidden) {
    ws.rowDimensions.set(row, { ...dim, hidden: true });
    return;
  }
  if (!dim?.hidden) return;
  const { hidden: _h, ...rest } = dim;
  if (Object.keys(rest).length === 0) ws.rowDimensions.delete(row);
  else ws.rowDimensions.set(row, rest);
}

export function runAdvancedFilter(ctl: EditorController, opts: AdvancedFilterOptions): MessageKey | undefined {
  const doc = ctl.doc;
  const ws = doc.ws;
  const { list } = opts;
  if (list.r2 <= list.r1) return 'dtAfListTooSmall';
  const data = values(ctl, list);
  const headers = (data[0] ?? []).map(asText);
  const compiled = opts.criteria ? compileCriteria(headers, values(ctl, opts.criteria)) : [];
  if (!compiled) return 'dtAfBadCriteria';

  // Destination columns: labels already in the copy-to row pick and order them, as in Excel.
  let outCols = headers.map((_, i) => i);
  if (opts.copyTo) {
    if (rangesIntersect(opts.copyTo, list)) return 'dtAfCopyOverlap';
    const labels: number[] = [];
    for (let c = opts.copyTo.c1; c <= opts.copyTo.c2; c++) {
      const cell = getCellAt(ws, opts.copyTo.r1, c);
      if (!cell || isBlank(cell)) continue;
      const i = headers.findIndex((h) => h.trim().toLowerCase() === getCellDisplayText(doc.wb, cell).trim().toLowerCase());
      if (i < 0) return 'dtAfBadCopyLabels';
      labels.push(i);
    }
    if (labels.length > 0) outCols = labels;
  }

  const seen = new Set<string>();
  const keep: number[] = [];
  for (let i = 1; i < data.length; i++) {
    const record = data[i] ?? [];
    if (!matches(record, compiled)) continue;
    if (opts.unique) {
      const key = outCols.map((c) => asText(record[c] ?? null).toLowerCase()).join('\u0000');
      if (seen.has(key)) continue;
      seen.add(key);
    }
    keep.push(list.r1 + i);
  }

  if (!opts.copyTo) {
    doc.transact('Advanced Filter', (tx) => {
      tx.sheet(ws, 'rowDimensions');
      const shown = new Set(keep);
      for (let r = list.r1 + 1; r <= list.r2; r++) setRowHidden(ws, r, !shown.has(r));
    });
    inPlace.set(ws, list);
    return undefined;
  }

  const dest = opts.copyTo;
  const out: Range = { r1: dest.r1, c1: dest.c1, r2: dest.r1 + keep.length, c2: dest.c1 + outCols.length - 1 };
  doc.transact('Advanced Filter', (tx) => {
    tx.cells(ws, out);
    const copy = (fromRow: number, toRow: number) =>
      outCols.forEach((c, j) => {
        const src = getCellAt(ws, fromRow, list.c1 + c);
        const value = data[fromRow - list.r1]?.[c] ?? null;
        // Records are copied as values with their formats, as Excel does.
        setValueAt(ws, toRow, dest.c1 + j, value !== null && typeof value === 'object' ? { kind: 'error', code: value.code } : value, src?.styleId ?? 0);
      });
    copy(list.r1, dest.r1);
    keep.forEach((r, k) => copy(r, dest.r1 + 1 + k));
  });
  return undefined;
}

/** Data ▸ Clear for an in-place Advanced Filter: show its rows again. */
export function clearAdvancedFilter(ctl: EditorController): void {
  const ws = ctl.doc.ws;
  const list = inPlace.get(ws);
  if (!list) return;
  inPlace.delete(ws);
  ctl.doc.transact('Clear Filter', (tx) => {
    tx.sheet(ws, 'rowDimensions');
    for (let r = list.r1 + 1; r <= list.r2; r++) setRowHidden(ws, r, false);
  });
}
