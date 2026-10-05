import { describe, expect, it } from 'vitest';
import { applyCompletion, argumentHintAt, completionAt, currentArgIndex, insertArgumentNames, rewriteReference, syntaxParts } from './formula-assist.ts';

describe('completionAt', () => {
  it('lists functions starting with the typed word', () => {
    const c = completionAt('=SU', 3, []);
    expect(c?.start).toBe(1);
    expect(c?.items.map((i) => i.name)).toContain('SUM');
    expect(c?.items.every((i) => i.name.startsWith('SU'))).toBe(true);
  });
  it('includes defined names', () => {
    expect(completionAt('=1+Sal', 6, ['Sales'])?.items.some((i) => i.name === 'Sales' && i.kind === 'name')).toBe(true);
  });
  it('ignores words inside strings, references and non-formulas', () => {
    expect(completionAt('="SU', 4, [])).toBeUndefined();
    expect(completionAt('=Sheet1!SU', 10, [])).toBeUndefined();
    expect(completionAt('SU', 2, [])).toBeUndefined();
  });
  it('inserts the function with its parenthesis', () => {
    const text = '=su';
    const c = completionAt(text, 3, []);
    const sum = c?.items.find((i) => i.name === 'SUM');
    expect(c && sum && applyCompletion(text, c, sum)).toEqual({ text: '=SUM(', caret: 5 });
  });
});

describe('argumentHintAt', () => {
  it('tracks the innermost call and argument', () => {
    expect(argumentHintAt('=SUM(A1,', 8)).toMatchObject({ fn: { name: 'SUM' }, argIndex: 1 });
    expect(argumentHintAt('=IF(A1>0,ROUND(B1', 17)).toMatchObject({ fn: { name: 'ROUND' }, argIndex: 0 });
    expect(argumentHintAt('=IF(A1>0,ROUND(B1,2),', 21)).toMatchObject({ fn: { name: 'IF' }, argIndex: 2 });
  });
  it('skips commas in strings and array constants', () => {
    expect(argumentHintAt('=SUM({1,2},"a,b",', 17)).toMatchObject({ argIndex: 2 });
  });
});

describe('syntaxParts', () => {
  it('keeps the repeating argument current past the end', () => {
    const { args } = syntaxParts('SUM(number1, [number2], ...)');
    expect(args).toEqual(['number1', '[number2]', '...']);
    expect(currentArgIndex(args, 5)).toBe(1);
    expect(currentArgIndex(syntaxParts('ROUND(number, num_digits)').args, 3)).toBe(-1);
  });
});

describe('insertArgumentNames', () => {
  it('spells out the arguments after a function name', () => {
    expect(insertArgumentNames('=ROUND(', 7)).toEqual({ text: '=ROUND(number, num_digits)', caret: 26 });
    expect(insertArgumentNames('=round', 6)?.text).toBe('=round(number, num_digits)');
    expect(insertArgumentNames('=A1+', 4)).toBeUndefined();
  });
});

describe('rewriteReference', () => {
  it('keeps $ anchoring and the sheet prefix', () => {
    expect(rewriteReference('$A$1:B2', { r1: 2, c1: 2, r2: 4, c2: 3 })).toBe('$B$2:C4');
    expect(rewriteReference("'My Sheet'!A1", { r1: 3, c1: 27, r2: 3, c2: 27 })).toBe("'My Sheet'!AA3");
    expect(rewriteReference('A1', { r1: 1, c1: 1, r2: 2, c2: 2 })).toBe('A1:B2');
  });
  it('leaves whole-row and whole-column references alone', () => {
    expect(rewriteReference('A:A', { r1: 1, c1: 2, r2: 1, c2: 2 })).toBeUndefined();
  });
});
