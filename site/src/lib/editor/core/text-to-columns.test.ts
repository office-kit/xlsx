import { describe, expect, it } from 'vitest';
import { parseBreaks, splitText, type DelimitedOptions } from './text-to-columns.ts';

const delimited = (over: Partial<DelimitedOptions> = {}): DelimitedOptions => ({
  kind: 'delimited',
  tab: false,
  semicolon: false,
  comma: false,
  space: false,
  other: '',
  consecutiveAsOne: false,
  qualifier: '"',
  ...over,
});

describe('splitText (delimited)', () => {
  it('splits on each chosen delimiter', () => {
    expect(splitText('a,b;c', delimited({ comma: true, semicolon: true }))).toEqual(['a', 'b', 'c']);
    expect(splitText('a\tb', delimited({ tab: true }))).toEqual(['a', 'b']);
    expect(splitText('a|b|c', delimited({ other: '|' }))).toEqual(['a', 'b', 'c']);
  });

  it('keeps empty fields unless consecutive delimiters are treated as one', () => {
    expect(splitText('a,,b', delimited({ comma: true }))).toEqual(['a', '', 'b']);
    expect(splitText('a,,b', delimited({ comma: true, consecutiveAsOne: true }))).toEqual(['a', 'b']);
    expect(splitText('a  b', delimited({ space: true, consecutiveAsOne: true }))).toEqual(['a', 'b']);
  });

  it('treats delimiters inside the text qualifier as literal text', () => {
    expect(splitText('"Smith, John",42', delimited({ comma: true }))).toEqual(['Smith, John', '42']);
    expect(splitText('"say ""hi""",x', delimited({ comma: true }))).toEqual(['say "hi"', 'x']);
    expect(splitText('"a,b"', delimited({ comma: true, qualifier: null }))).toEqual(['"a', 'b"']);
  });

  it('returns the text unchanged when no delimiter is chosen', () => {
    expect(splitText('a,b', delimited())).toEqual(['a,b']);
  });
});

describe('splitText (fixed width)', () => {
  it('cuts at each break position', () => {
    expect(splitText('2024ABC123', { kind: 'fixed', breaks: [4, 7] })).toEqual(['2024', 'ABC', '123']);
  });

  it('ignores breaks beyond the text and duplicates', () => {
    expect(splitText('abc', { kind: 'fixed', breaks: [2, 2, 10] })).toEqual(['ab', 'c']);
  });
});

describe('parseBreaks', () => {
  it('parses comma- or space-separated positions, sorted and unique', () => {
    expect(parseBreaks('10, 5 5')).toEqual([5, 10]);
    expect(parseBreaks('')).toEqual([]);
  });

  it('rejects non-numeric or zero positions', () => {
    expect(parseBreaks('a')).toBeUndefined();
    expect(parseBreaks('0')).toBeUndefined();
  });
});
