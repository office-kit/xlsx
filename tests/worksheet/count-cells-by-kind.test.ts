// Tests for countCellsByKind — value-kind histogram of populated cells.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  countCellsByKind,
  ensureCell,
  setCell,
} from '../../src/worksheet/worksheet.js';

const zeros = {
  null: 0,
  string: 0,
  number: 0,
  boolean: 0,
  date: 0,
  duration: 0,
  error: 0,
  'rich-text': 0,
  formula: 0,
};

describe('countCellsByKind', () => {
  it('returns all-zero buckets for an empty worksheet', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(countCellsByKind(ws)).toEqual(zeros);
  });

  it('counts each kind for one example cell apiece', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'str');
    setCell(ws, 2, 1, 42);
    setCell(ws, 3, 1, true);
    setCell(ws, 4, 1, new Date('2025-01-01'));
    setCell(ws, 5, 1, { kind: 'duration', ms: 1000 });
    setCell(ws, 6, 1, { kind: 'error', code: '#REF!' });
    setCell(ws, 7, 1, { kind: 'rich-text', runs: [] });
    const f = ensureCell(ws, 8, 1);
    setFormula(f, 'A1+1');
    setCell(ws, 9, 1, null);
    expect(countCellsByKind(ws)).toEqual({
      ...zeros,
      string: 1,
      number: 1,
      boolean: 1,
      date: 1,
      duration: 1,
      error: 1,
      'rich-text': 1,
      formula: 1,
      null: 1,
    });
  });

  it('aggregates across multiple cells of the same kind', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 1, 'b');
    setCell(ws, 3, 1, 'c');
    setCell(ws, 1, 2, 1);
    setCell(ws, 2, 2, 2);
    expect(countCellsByKind(ws)).toEqual({ ...zeros, string: 3, number: 2 });
  });

  it('walks every populated row (sparse layout works)', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 100, 100, 'far');
    setCell(ws, 1, 1, 1);
    expect(countCellsByKind(ws)).toEqual({ ...zeros, string: 1, number: 1 });
  });

  it('does not count cells that point to absent rows after wipe', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setCell(ws, 1, 1, 'a');
    setCell(ws, 2, 1, 'b');
    ws.rows.delete(2);
    expect(countCellsByKind(ws)).toEqual({ ...zeros, string: 1 });
  });
});
