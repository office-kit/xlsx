import { type CellValue, makeArrayFormula, makeFormula, makeSharedFormula } from '@office-kit/xlsx/cell';
import { addDefinedName, addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { addTable, deleteCell, getCell, makeTableColumn, makeTableDefinition, setCell, type Worksheet } from '@office-kit/xlsx/worksheet';
import { describe, expect, test } from 'vitest';
import { CalcEngine, type CellRef } from './index.ts';

/** The value a cell shows: a formula's cached result, or the literal. */
const shown = (ws: Worksheet, row: number, col: number): CellValue | undefined => {
  const v = getCell(ws, row, col)?.value;
  return v !== null && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? v.cachedValue : v;
};

const setup = (cells: ReadonlyArray<[number, number, CellValue]>) => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  for (const [r, c, v] of cells) setCell(ws, r, c, v);
  const engine = new CalcEngine(wb, { now: () => new Date(2024, 0, 15, 12), random: () => 0.5 });
  engine.recalculateAll();
  return { wb, ws, engine };
};

const at = (row: number, col: number): CellRef => ({ sheet: 'S', row, col });

describe('recalculation', () => {
  test('update recomputes dependents and reports changed cells', () => {
    const { ws, engine } = setup([
      [1, 1, 2],
      [2, 1, makeFormula('A1*10')],
      [3, 1, makeFormula('SUM(A1:A2)')],
      [4, 1, makeFormula('B1+1')],
    ]);
    expect(shown(ws, 3, 1)).toBe(22);
    setCell(ws, 1, 1, 3);
    const changed = engine.update([at(1, 1)]);
    expect(shown(ws, 3, 1)).toBe(33);
    expect(changed).toEqual(expect.arrayContaining([at(2, 1), at(3, 1)]));
    expect(changed).not.toContainEqual(at(4, 1));
  });

  test('a cell entering a watched range triggers its readers', () => {
    const { ws, engine } = setup([[1, 2, makeFormula('SUM(A:A)')]]);
    setCell(ws, 500, 1, 7);
    engine.update([at(500, 1)]);
    expect(shown(ws, 1, 2)).toBe(7);
  });

  test('a long dependency chain does not overflow the stack', () => {
    const cells: Array<[number, number, CellValue]> = [[1, 1, 1]];
    for (let r = 2; r <= 20_000; r++) cells.push([r, 1, makeFormula(`A${r - 1}+1`)]);
    const { ws } = setup(cells);
    expect(shown(ws, 20_000, 1)).toBe(20_000);
  });

  test('circular references compute to 0 and are reported', () => {
    const { ws, engine } = setup([
      [1, 1, makeFormula('B1+1')],
      [1, 2, makeFormula('A1+1')],
      [1, 3, makeFormula('A1+5')],
    ]);
    expect(engine.circularRefs).toEqual(expect.arrayContaining([at(1, 1), at(1, 2)]));
    expect(engine.circularRefs).not.toContainEqual(at(1, 3));
    expect(typeof shown(ws, 1, 3)).toBe('number');
  });

  test('volatile formulas recompute on every update', () => {
    let n = 0;
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, makeFormula('RAND()'));
    setCell(ws, 2, 1, 1);
    const engine = new CalcEngine(wb, { random: () => ++n / 10 });
    engine.recalculateAll();
    expect(shown(ws, 1, 1)).toBe(0.1);
    engine.update([at(2, 1)]);
    expect(shown(ws, 1, 1)).toBe(0.2);
  });

  test('shared formula followers reuse the master with their offset', () => {
    const { ws } = setup([
      [1, 1, 1],
      [2, 1, 2],
      [3, 1, 3],
      [1, 2, makeSharedFormula(0, 'A1*2', 'B1:B3')],
      [2, 2, makeSharedFormula(0)],
      [3, 2, makeSharedFormula(0)],
    ]);
    expect([shown(ws, 1, 2), shown(ws, 2, 2), shown(ws, 3, 2)]).toEqual([2, 4, 6]);
  });

  test('defined names and structured references', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    setCell(ws, 1, 1, 'Item');
    setCell(ws, 1, 2, 'Qty');
    setCell(ws, 2, 1, 'a');
    setCell(ws, 2, 2, 3);
    setCell(ws, 3, 1, 'b');
    setCell(ws, 3, 2, 4);
    addTable(ws, makeTableDefinition({ id: 1, displayName: 'T', ref: 'A1:B3', columns: [makeTableColumn({ id: 1, name: 'Item' }), makeTableColumn({ id: 2, name: 'Qty' })] }));
    addDefinedName(wb, { name: 'Rate', value: '10' });
    setCell(ws, 1, 4, makeFormula('SUM(T[Qty])*Rate'));
    setCell(ws, 2, 3, makeFormula('T[@Qty]+1'));
    const engine = new CalcEngine(wb);
    engine.recalculateAll();
    expect(shown(ws, 1, 4)).toBe(70);
    expect(shown(ws, 2, 3)).toBe(4);
  });

  test('evaluate runs a formula as if it sat in a cell', () => {
    const { engine } = setup([[1, 1, 5]]);
    expect(engine.evaluate('A1>3', 'S', 9, 9)).toBe(true);
    expect(engine.evaluate('B1', 'S', 1, 1)).toBe(0);
    expect(engine.evaluate('SUM(', 'S', 1, 1)).toEqual({ kind: 'error', code: '#NAME?' });
  });
});

describe('dynamic arrays', () => {
  test('a multi-cell result spills and its anchor becomes an array formula', () => {
    const { ws } = setup([[1, 1, makeFormula('SEQUENCE(3)')]]);
    expect(getCell(ws, 1, 1)?.value).toMatchObject({ kind: 'formula', t: 'array', ref: 'A1:A3', cachedValue: 1 });
    expect([getCell(ws, 2, 1)?.value, getCell(ws, 3, 1)?.value]).toEqual([2, 3]);
  });

  test('a blocked spill shows #SPILL! and recovers when the blocker leaves', () => {
    const { ws, engine } = setup([
      [1, 1, makeFormula('SEQUENCE(3)')],
      [3, 1, 'x'],
    ]);
    expect(getCell(ws, 1, 1)?.value).toMatchObject({ cachedValue: '#SPILL!', cachedValueType: 'error' });
    deleteCell(ws, 3, 1);
    engine.update([at(3, 1)]);
    expect([shown(ws, 1, 1), getCell(ws, 3, 1)?.value]).toEqual([1, 3]);
  });

  test('typing into a spilled cell blocks the spill', () => {
    const { ws, engine } = setup([[1, 1, makeFormula('SEQUENCE(3)')]]);
    setCell(ws, 2, 1, 'mine');
    engine.update([at(2, 1)]);
    expect(getCell(ws, 1, 1)?.value).toMatchObject({ cachedValue: '#SPILL!', cachedValueType: 'error' });
    expect(getCell(ws, 2, 1)?.value).toBe('mine');
  });

  test('a shrinking spill clears the cells it no longer covers', () => {
    const { ws, engine } = setup([
      [1, 2, 3],
      [1, 1, makeFormula('SEQUENCE(B1)')],
    ]);
    setCell(ws, 1, 2, 1);
    engine.update([at(1, 2)]);
    expect(getCell(ws, 1, 1)?.value).toMatchObject({ t: 'normal', cachedValue: 1 });
    expect(getCell(ws, 2, 1)).toBeUndefined();
  });

  test('# reads the whole spill and follows its growth', () => {
    const { ws, engine } = setup([
      [1, 2, 2],
      [1, 1, makeFormula('SEQUENCE(B1)')],
      [1, 3, makeFormula('SUM(A1#)')],
    ]);
    expect(shown(ws, 1, 3)).toBe(3);
    setCell(ws, 1, 2, 4);
    engine.update([at(1, 2)]);
    expect(shown(ws, 1, 3)).toBe(10);
  });

  test('a CSE array formula loaded from a file keeps its range', () => {
    const { ws } = setup([[1, 1, makeArrayFormula('A1:A2', '{1;2}*10')]]);
    expect([shown(ws, 1, 1), getCell(ws, 2, 1)?.value]).toEqual([10, 20]);
  });
});
