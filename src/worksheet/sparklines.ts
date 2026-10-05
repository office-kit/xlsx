// Sparklines (Excel 2010+).
//
// Sparklines have no slot in the ECMA-376 worksheet schema; Excel stores them
// as `<x14:sparklineGroups>` inside the worksheet `<extLst>`. A group shares
// one type, colour set and flag set across every sparkline it holds, which is
// why Excel's Sparkline tab edits the whole group at once.

import type { Color } from '../styles/colors.js';

/**
 * `line` draws a polyline, `column` bars, and `stacked` Excel's Win/Loss
 * bars. Excel draws any other `type` value as a line, so the reader treats it
 * as absent.
 */
export type SparklineType = 'line' | 'column' | 'stacked';
export type SparklineEmptyCells = 'gap' | 'zero' | 'span';
export type SparklineAxisType = 'individual' | 'group' | 'custom';

/** One sparkline: the cell it draws in and the data it plots. */
export interface Sparkline {
  /** The data range as a formula reference, e.g. `Sheet1!A2:E2`. */
  formula: string;
  /** The single cell the sparkline is drawn in, e.g. `F2`. */
  location: string;
}

export interface SparklineGroup {
  /** Defaults to `line` when absent. */
  type?: SparklineType;
  sparklines: Sparkline[];
  colorSeries?: Color;
  colorNegative?: Color;
  colorAxis?: Color;
  colorMarkers?: Color;
  colorFirst?: Color;
  colorLast?: Color;
  colorHigh?: Color;
  colorLow?: Color;
  /** Line sparklines only: a marker on every point. */
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  displayXAxis?: boolean;
  displayHidden?: boolean;
  rightToLeft?: boolean;
  dateAxis?: boolean;
  /** Range of dates for a date axis (`<xm:f>` directly under the group). */
  dateFormula?: string;
  displayEmptyCellsAs?: SparklineEmptyCells;
  /** Line weight in points; Excel's default is 0.75. */
  lineWeight?: number;
  minAxisType?: SparklineAxisType;
  maxAxisType?: SparklineAxisType;
  manualMin?: number;
  manualMax?: number;
}

/**
 * A sparkline group with Excel's Insert ▸ Sparklines defaults: the theme's
 * dark-blue series and red negatives / markers.
 */
export function makeSparklineGroup(opts: Partial<SparklineGroup> & { sparklines: Sparkline[] }): SparklineGroup {
  return {
    colorSeries: { theme: 4, tint: -0.499984740745262 },
    colorNegative: { theme: 5 },
    colorAxis: { rgb: 'FF000000' },
    colorMarkers: { theme: 4, tint: -0.499984740745262 },
    colorFirst: { theme: 4, tint: 0.39997558519241921 },
    colorLast: { theme: 4, tint: 0.39997558519241921 },
    colorHigh: { theme: 4 },
    colorLow: { theme: 4 },
    displayEmptyCellsAs: 'gap',
    ...opts,
  };
}
