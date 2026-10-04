// Reference-level text utilities for the formula bar: coloured reference
// highlighting and F4 anchor cycling. Both work on half-typed formulas, so they
// tokenize leniently and never throw.

import type { RefArea } from './ast.ts';
import { renderArea } from './address.ts';
import { type Token, tokenize } from './lexer.ts';

export interface FormulaReferenceSpan {
  /** Offsets into the text passed in, `end` exclusive, prefix (`Sheet1!`) included. */
  readonly start: number;
  readonly end: number;
  readonly sheet?: string;
  /** Normalised rectangle; whole columns / rows span the full grid. */
  readonly range: { readonly r1: number; readonly c1: number; readonly r2: number; readonly c2: number };
  readonly text: string;
}

type RefToken = Extract<Token, { kind: 'ref' }>;

const isRefToken = (t: Token): t is RefToken & { area: RefArea } => t.kind === 'ref' && t.area !== undefined;

/** Every A1 reference in `text`, in order, for coloured highlighting while editing. */
export function formulaReferences(text: string): FormulaReferenceSpan[] {
  const spans: FormulaReferenceSpan[] = [];
  for (const t of tokenize(text, true)) {
    if (!isRefToken(t)) continue;
    const a = t.area;
    spans.push({
      start: t.start,
      end: t.end,
      ...(t.prefix !== undefined ? { sheet: t.prefix.sheet } : {}),
      range: { r1: Math.min(a.r1, a.r2), c1: Math.min(a.c1, a.c2), r2: Math.max(a.r1, a.r2), c2: Math.max(a.c1, a.c2) },
      text: text.slice(t.start, t.end),
    });
  }
  return spans;
}

// Excel's F4 order: relative → absolute → row-absolute → column-absolute → relative.
const nextAnchors = (rAbs: boolean, cAbs: boolean): { rAbs: boolean; cAbs: boolean } => {
  if (!rAbs && !cAbs) return { rAbs: true, cAbs: true };
  if (rAbs && cAbs) return { rAbs: true, cAbs: false };
  if (rAbs) return { rAbs: false, cAbs: true };
  return { rAbs: false, cAbs: false };
};

const cycle = (a: RefArea): RefArea => {
  switch (a.kind) {
    case 'cols':
      return { ...a, c1Abs: !a.c1Abs, c2Abs: !a.c1Abs };
    case 'rows':
      return { ...a, r1Abs: !a.r1Abs, r2Abs: !a.r1Abs };
    default: {
      // The first corner decides the state; the whole reference moves together.
      const { rAbs, cAbs } = nextAnchors(a.r1Abs, a.c1Abs);
      return { ...a, r1Abs: rAbs, r2Abs: rAbs, c1Abs: cAbs, c2Abs: cAbs };
    }
  }
};

/**
 * F4: cycle the `$` anchors of the reference the caret touches
 * (`A1 → $A$1 → A$1 → $A1 → A1`). The caret lands after the rewritten
 * reference. Text without a reference at the caret comes back unchanged.
 */
export function toggleReferenceAt(text: string, caret: number): { text: string; caret: number } {
  const target = tokenize(text, true).find((t) => isRefToken(t) && t.start <= caret && caret <= t.end);
  if (target === undefined || !isRefToken(target)) return { text, caret };
  const rendered = renderArea(cycle(target.area));
  return {
    text: text.slice(0, target.prefixEnd) + rendered + text.slice(target.end),
    caret: target.prefixEnd + rendered.length,
  };
}
