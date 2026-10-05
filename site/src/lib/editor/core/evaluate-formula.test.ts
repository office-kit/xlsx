import { describe, expect, it } from 'vitest';
import type { CalcScalar } from '../calc/index.ts';
import { startEvaluation, stepEvaluation } from './evaluate-formula.ts';

// The sub-formulas the steps hand out, with their values.
const results: Record<string, CalcScalar> = { A1: 2, A2: 3, '3*2': 6, '2+6': 8 };
const evaluate = (f: string): CalcScalar => results[f] ?? { kind: 'error', code: '#NAME?' };

describe('Evaluate Formula steps', () => {
  it('reduces references first, then operations, leftmost innermost', () => {
    let s = startEvaluation('A1+A2*2');
    expect(s.text.slice(s.start, s.end)).toBe('A1');
    s = stepEvaluation(s, evaluate);
    expect(s.text).toBe('2+A2*2');
    expect(s.text.slice(s.start, s.end)).toBe('A2');
    s = stepEvaluation(s, evaluate);
    expect(s.text.slice(s.start, s.end)).toBe('3*2');
    s = stepEvaluation(s, evaluate);
    s = stepEvaluation(s, evaluate);
    expect(s).toMatchObject({ text: '8', done: true });
  });

  it('keeps parentheses the tree needs and quotes strings', () => {
    expect(startEvaluation('(A1+A2)*2').text).toBe('(A1+A2)*2');
    expect(startEvaluation('B1&"a""b"').text).toBe('B1&"a""b"');
  });

  it('evaluates whole functions over big ranges at once', () => {
    const s = startEvaluation('SUM(A:A)');
    expect(s.text.slice(s.start, s.end)).toBe('SUM(A:A)');
  });
});
