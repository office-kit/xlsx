import { describe, expect, it } from 'vitest';
import { dateToSerial } from './input.ts';
import { seriesValues, type SeriesOptions } from './series.ts';

const opts = (over: Partial<SeriesOptions>): SeriesOptions => ({ type: 'linear', step: 1, dateUnit: 'day', ...over });

describe('seriesValues', () => {
  it('steps a linear series and stops at the stop value', () => {
    expect(seriesValues(1, 4, opts({ step: 2 }))).toEqual([3, 5, 7, 9]);
    expect(seriesValues(1, 10, opts({ step: 2, stop: 6 }))).toEqual([3, 5]);
    expect(seriesValues(10, 10, opts({ step: -3, stop: 2 }))).toEqual([7, 4]);
  });

  it('avoids floating-point drift', () => {
    expect(seriesValues(0, 3, opts({ step: 0.1 }))).toEqual([0.1, 0.2, 0.3]);
  });

  it('multiplies for a growth series', () => {
    expect(seriesValues(2, 3, opts({ type: 'growth', step: 3 }))).toEqual([6, 18, 54]);
    expect(seriesValues(100, 5, opts({ type: 'growth', step: 0.5, stop: 20 }))).toEqual([50, 25]);
  });

  it('adds days, weekdays, months and years to dates', () => {
    const fri = dateToSerial(2024, 3, 1); // a Friday
    expect(seriesValues(fri, 2, opts({ type: 'date', dateUnit: 'day' }))).toEqual([fri + 1, fri + 2]);
    expect(seriesValues(fri, 2, opts({ type: 'date', dateUnit: 'weekday' }))).toEqual([fri + 3, fri + 4]);
    const jan31 = dateToSerial(2024, 1, 31);
    expect(seriesValues(jan31, 2, opts({ type: 'date', dateUnit: 'month' }))).toEqual([dateToSerial(2024, 2, 29), dateToSerial(2024, 3, 31)]);
    expect(seriesValues(jan31, 1, opts({ type: 'date', dateUnit: 'year', step: 2 }))).toEqual([dateToSerial(2026, 1, 31)]);
  });
});
