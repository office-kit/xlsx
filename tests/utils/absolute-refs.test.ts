// Tests for the $-marker options on tupleToCoordinate / boundariesToRangeString.
//
// These are what a caller tracking integer coordinates uses to build formula
// text, instead of concatenating a column letter by hand.

import { describe, expect, it } from 'vitest';
import { boundariesToRangeString, tupleToCoordinate } from '../../src/utils/coordinate.js';

describe('tupleToCoordinate absolute markers', () => {
  it('defaults to a relative ref', () => {
    expect(tupleToCoordinate(2, 5)).toBe('B5');
    expect(tupleToCoordinate(2, 5, {})).toBe('B5');
  });

  it('places $ on each axis independently', () => {
    expect(tupleToCoordinate(2, 5, { absoluteCol: true })).toBe('$B5');
    expect(tupleToCoordinate(2, 5, { absoluteRow: true })).toBe('B$5');
    expect(tupleToCoordinate(2, 5, { absoluteCol: true, absoluteRow: true })).toBe('$B$5');
  });

  it('works past the single-letter columns', () => {
    expect(tupleToCoordinate(27, 1, { absoluteCol: true, absoluteRow: true })).toBe('$AA$1');
  });
});

describe('boundariesToRangeString absolute markers', () => {
  const bounds = { minCol: 1, minRow: 4, maxCol: 8, maxRow: 20 };

  it('formats a rectangle relative by default', () => {
    expect(boundariesToRangeString(bounds)).toBe('A4:H20');
  });

  it('forwards the markers to both corners', () => {
    expect(boundariesToRangeString(bounds, { absoluteCol: true, absoluteRow: true })).toBe('$A$4:$H$20');
    expect(boundariesToRangeString(bounds, { absoluteCol: true })).toBe('$A4:$H20');
  });

  it('collapses a single cell and still marks it', () => {
    const one = { minCol: 2, minRow: 5, maxCol: 2, maxRow: 5 };
    expect(boundariesToRangeString(one, { absoluteCol: true, absoluteRow: true })).toBe('$B$5');
  });
});
