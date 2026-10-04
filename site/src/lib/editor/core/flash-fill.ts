// Flash Fill (Data ▸ Flash Fill, Cmd/Ctrl+E): infer how the examples typed in
// a column were derived from the neighbouring columns and fill the rest.
//
// The search space is small and deliberate: each output is either one token
// of a source cell (split on common separators, counted from the start or
// the end, with a case change), a fixed-length prefix/suffix, or two such
// pieces joined by a literal. That covers the everyday cases — first/last
// names, e-mail domains, initials, codes — and fails visibly otherwise.

import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { getCellAt, isBlank } from './cells.ts';

type Case = 'same' | 'upper' | 'lower' | 'proper';

interface Piece {
  readonly col: number;
  readonly sep: string;
  /** Token index; negative counts from the end. `null` = the whole cell. */
  readonly index: number | null;
  readonly kase: Case;
  /** Take only the first n characters of the token. */
  readonly take?: number;
}

const SEPARATORS = [' ', ',', '@', '-', '/', '.', '_', ':', ';'];

function applyCase(s: string, k: Case): string {
  switch (k) {
    case 'upper':
      return s.toUpperCase();
    case 'lower':
      return s.toLowerCase();
    case 'proper':
      return s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, a: string, b: string) => a + b.toUpperCase());
    default:
      return s;
  }
}

function evalPiece(p: Piece, source: (col: number) => string): string | undefined {
  const text = source(p.col);
  let token: string | undefined;
  if (p.index === null) token = text;
  else {
    const parts = text.split(p.sep).filter((x) => x !== '');
    token = p.index >= 0 ? parts[p.index] : parts[parts.length + p.index];
  }
  if (token === undefined) return undefined;
  if (p.take !== undefined) token = token.slice(0, p.take);
  return applyCase(token.trim(), p.kase);
}

function candidates(cols: readonly number[], sample: (col: number) => string): Piece[] {
  const out: Piece[] = [];
  const kases: Case[] = ['same', 'upper', 'lower', 'proper'];
  for (const col of cols) {
    const text = sample(col);
    for (const kase of kases) {
      out.push({ col, sep: '', index: null, kase });
      for (const take of [1, 2, 3]) out.push({ col, sep: '', index: null, kase, take });
      for (const sep of SEPARATORS) {
        const n = text.split(sep).filter((x) => x !== '').length;
        if (n < 2) continue;
        for (let i = 0; i < Math.min(n, 6); i++) {
          out.push({ col, sep, index: i, kase });
          out.push({ col, sep, index: -(i + 1), kase });
          out.push({ col, sep, index: i, kase, take: 1 });
        }
      }
    }
  }
  return out;
}

type Program = { pieces: Piece[]; joiner: string };

function run(p: Program, source: (col: number) => string): string | undefined {
  const parts: string[] = [];
  for (const piece of p.pieces) {
    const v = evalPiece(piece, source);
    if (v === undefined) return undefined;
    parts.push(v);
  }
  return parts.join(p.joiner);
}

/**
 * Fill blanks in `target` column between `r1` and `r2` from the examples
 * already present there. Returns the produced values by row, or undefined
 * when no program explains every example.
 */
export function flashFill(ws: Worksheet, target: number, r1: number, r2: number): Map<number, string> | undefined {
  const sourceCols: number[] = [];
  for (let c = target - 1; c >= Math.max(1, target - 4); c--) if (!isBlank(getCellAt(ws, r1, c)) || !isBlank(getCellAt(ws, r1 + 1, c))) sourceCols.push(c);
  for (let c = target + 1; c <= target + 2; c++) if (!isBlank(getCellAt(ws, r1, c))) sourceCols.push(c);
  if (sourceCols.length === 0) return undefined;
  const text = (row: number, col: number): string => {
    const v = getCellAt(ws, row, col)?.value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object' && !(v instanceof Date)) {
      if (v.kind === 'formula') return v.cachedValue === undefined ? '' : String(v.cachedValue);
      if (v.kind === 'rich-text') return v.runs.map((r) => r.text).join('');
      return '';
    }
    return String(v);
  };
  const examples: Array<{ row: number; out: string }> = [];
  for (let r = r1; r <= r2; r++) {
    const out = text(r, target);
    if (out !== '') examples.push({ row: r, out });
  }
  const first = examples[0];
  if (!first) return undefined;
  const pieces = candidates(sourceCols, (col) => text(first.row, col));
  const fits = (p: Program) => examples.every((ex) => run(p, (col) => text(ex.row, col)) === ex.out);

  let found: Program | undefined;
  for (const piece of pieces) {
    const p = { pieces: [piece], joiner: '' };
    if (fits(p)) {
      found = p;
      break;
    }
  }
  const pair = (): Program | undefined => {
    for (const joiner of ['', ' ', ', ', '.', '-', '_', '@', '/']) {
      for (const a of pieces) {
        const va = evalPiece(a, (col) => text(first.row, col));
        if (va === undefined || !first.out.startsWith(va + joiner)) continue;
        for (const b of pieces) {
          const p = { pieces: [a, b], joiner };
          if (fits(p)) return p;
        }
      }
    }
    return undefined;
  };
  found ??= pair();
  if (!found) return undefined;
  const program = found;
  const result = new Map<number, string>();
  for (let r = r1; r <= r2; r++) {
    if (text(r, target) !== '') continue;
    if (sourceCols.every((c) => text(r, c) === '')) continue;
    const v = run(program, (col) => text(r, col));
    if (v !== undefined) result.set(r, v);
  }
  return result;
}
