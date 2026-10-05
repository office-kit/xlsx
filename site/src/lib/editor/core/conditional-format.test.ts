import { makeFormula } from '@office-kit/xlsx/cell';
import type { Workbook } from '@office-kit/xlsx/workbook';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import type { ConditionalFormattingRule, Worksheet } from '@office-kit/xlsx/worksheet';
import { makeCfRule, makeConditionalFormatting, setCell } from '@office-kit/xlsx/worksheet';
import { beforeEach, describe, expect, it } from 'vitest';
import { CalcEngine } from '../calc/index.ts';
import {
  conditionalOverlay,
  defaultRuleForm,
  formFromRule,
  inTimePeriod,
  operandFormula,
  operandText,
  ruleFromForm,
  type RuleForm,
  parseAppliesTo,
  parseVisualRule,
  registerCfFormat,
  subtractRange,
  visualRuleXml,
  type OverlayFn,
} from './conditional-format.ts';
import { StyleResolver } from './render-style.ts';

let wb: Workbook;
let ws: Worksheet;

beforeEach(() => {
  wb = createWorkbook();
  ws = addWorksheet(wb, 'Sheet1');
});

function column(values: Array<number | string | null>, col = 1): void {
  values.forEach((v, i) => setCell(ws, i + 1, col, v));
}

function pushRule(sqref: string, opts: Parameters<typeof makeCfRule>[0]): ConditionalFormattingRule {
  const made = makeCfRule(opts);
  ws.conditionalFormatting.push(makeConditionalFormatting({ sqref, rules: [made] }));
  return made;
}

function overlay(): OverlayFn {
  const calc = new CalcEngine(wb);
  calc.recalculateAll();
  const fn = conditionalOverlay(wb, ws, calc, new StyleResolver(wb));
  if (!fn) throw new Error('expected rules');
  return fn;
}

const RED = { fill: '#FFC7CE', color: '#9C0006' };

describe('conditionalOverlay', () => {
  it('returns undefined for a sheet without rules', () => {
    expect(conditionalOverlay(wb, ws, new CalcEngine(wb), new StyleResolver(wb))).toBeUndefined();
  });

  it('applies cellIs operators with the rule DXF', () => {
    column([1, 5, 10]);
    const dxfId = registerCfFormat(wb, RED);
    pushRule('A1:A3', { type: 'cellIs', priority: 1, operator: 'greaterThan', formulas: ['4'], dxfId });
    const at = overlay();
    expect(at(1, 1, undefined)).toBeUndefined();
    expect(at(2, 1, undefined)).toMatchObject({ fill: '#FFC7CE', color: '#9C0006' });
    expect(at(3, 1, undefined)?.fill).toBe('#FFC7CE');
    expect(at(2, 2, undefined)).toBeUndefined();
  });

  it('handles between in either bound order', () => {
    column([1, 5, 10]);
    const dxfId = registerCfFormat(wb, RED);
    pushRule('A1:A3', { type: 'cellIs', priority: 1, operator: 'between', formulas: ['6', '2'], dxfId });
    const at = overlay();
    expect([1, 2, 3].map((r) => at(r, 1, undefined) !== undefined)).toEqual([false, true, false]);
  });

  it('shifts relative references in expression rules', () => {
    column([1, 2, 3, 4]);
    column([0, 2, 0, 4], 2);
    const dxfId = registerCfFormat(wb, { bold: true });
    pushRule('A1:A4', { type: 'expression', priority: 1, formulas: ['A1=$B1'], dxfId });
    const at = overlay();
    expect([1, 2, 3, 4].map((r) => at(r, 1, undefined)?.bold === true)).toEqual([false, true, false, true]);
  });

  it('respects priority and stopIfTrue', () => {
    column([10]);
    const red = registerCfFormat(wb, { fill: '#FF0000', bold: true });
    const blue = registerCfFormat(wb, { fill: '#0000FF', italic: true });
    pushRule('A1', { type: 'cellIs', priority: 2, operator: 'greaterThan', formulas: ['1'], dxfId: blue });
    const first = pushRule('A1', { type: 'cellIs', priority: 1, operator: 'greaterThan', formulas: ['1'], dxfId: red });
    expect(overlay()(1, 1, undefined)).toMatchObject({ fill: '#FF0000', bold: true, italic: true });
    first.stopIfTrue = true;
    expect(overlay()(1, 1, undefined)?.italic).toBeUndefined();
  });

  it('interpolates a three-colour scale', () => {
    column([0, 50, 100]);
    pushRule('A1:A3', {
      type: 'colorScale',
      priority: 1,
      innerXml: '<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale>',
    });
    const at = overlay();
    expect(at(1, 1, undefined)?.fill).toBe('#F8696B');
    expect(at(2, 1, undefined)?.fill).toBe('#FFEB84');
    expect(at(3, 1, undefined)?.fill).toBe('#63BE7B');
  });

  it('sizes data bars between min and max lengths and hides values on request', () => {
    column([0, 10]);
    pushRule('A1:A2', { type: 'dataBar', priority: 1, innerXml: '<dataBar showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>' });
    const at = overlay();
    expect(at(1, 1, undefined)?.bar).toEqual({ start: 0, end: 0.1, color: '#638EC6' });
    expect(at(2, 1, undefined)?.bar?.end).toBeCloseTo(0.9);
    expect(at(2, 1, undefined)?.hideValue).toBe(true);
  });

  it('picks icons by percent thresholds and honours reverse', () => {
    column([0, 50, 100]);
    const icons = pushRule('A1:A3', { type: 'iconSet', priority: 1, innerXml: '<iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet>' });
    let at = overlay();
    expect([1, 2, 3].map((r) => at(r, 1, undefined)?.icon?.glyph)).toEqual(['↓', '→', '↑']);
    icons.innerXml = '<iconSet iconSet="3Arrows" reverse="1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet>';
    at = overlay();
    expect(at(1, 1, undefined)?.icon?.glyph).toBe('↑');
  });

  it('highlights top / bottom N and percent', () => {
    column([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    const dxfId = registerCfFormat(wb, RED);
    pushRule('A1:A10', { type: 'top10', priority: 1, rank: 3, dxfId });
    pushRule('B1:B1', { type: 'top10', priority: 2, rank: 20, percent: true, bottom: true, dxfId });
    const at = overlay();
    const hits = Array.from({ length: 10 }, (_, i) => at(i + 1, 1, undefined) !== undefined);
    expect(hits).toEqual([false, false, false, false, false, false, false, true, true, true]);
  });

  it('compares with the average and standard deviations', () => {
    column([1, 2, 3, 4, 5]);
    const dxfId = registerCfFormat(wb, RED);
    pushRule('A1:A5', { type: 'aboveAverage', priority: 1, aboveAverage: false, dxfId });
    const at = overlay();
    expect([1, 2, 3, 4, 5].map((r) => at(r, 1, undefined) !== undefined)).toEqual([true, true, false, false, false]);
  });

  it('marks duplicate and unique values case-insensitively', () => {
    column(['a', 'A', 'b', null]);
    const dxfId = registerCfFormat(wb, RED);
    pushRule('A1:A4', { type: 'duplicateValues', priority: 1, dxfId });
    pushRule('B1:B4', { type: 'uniqueValues', priority: 2, dxfId });
    const at = overlay();
    expect([1, 2, 3, 4].map((r) => at(r, 1, undefined) !== undefined)).toEqual([true, true, false, false]);
  });

  it('evaluates the text, blank and error families', () => {
    column(['Apple pie', 'banana', null]);
    setCell(ws, 4, 1, makeFormula('1/0'));
    const dxfId = registerCfFormat(wb, RED);
    pushRule('A1:A4', { type: 'containsText', priority: 1, operator: 'containsText', text: 'PIE', dxfId });
    pushRule('B1:B4', { type: 'beginsWith', priority: 2, operator: 'beginsWith', text: 'ban', dxfId });
    pushRule('C1:C4', { type: 'containsBlanks', priority: 3, dxfId });
    pushRule('D1:D4', { type: 'containsErrors', priority: 4, dxfId });
    pushRule('A3', { type: 'containsBlanks', priority: 5, dxfId });
    pushRule('A4', { type: 'containsErrors', priority: 6, dxfId });
    const at = overlay();
    expect(at(1, 1, undefined)).toBeDefined();
    expect(at(2, 1, undefined)).toBeUndefined();
    expect(at(3, 1, undefined)).toBeDefined(); // blank
    expect(at(4, 1, undefined)).toBeDefined(); // #DIV/0!
  });
});

describe('inTimePeriod', () => {
  // 45201 is Monday 2023-10-02 in the 1900 date system.
  const today = 45201;
  it('recognises days, weeks and months', () => {
    expect(inTimePeriod('today', today, today, false)).toBe(true);
    expect(inTimePeriod('yesterday', today - 1, today, false)).toBe(true);
    expect(inTimePeriod('last7Days', today - 6, today, false)).toBe(true);
    expect(inTimePeriod('last7Days', today - 7, today, false)).toBe(false);
    expect(inTimePeriod('thisWeek', today - 1, today, false)).toBe(true); // Sunday
    expect(inTimePeriod('lastWeek', today - 2, today, false)).toBe(true); // previous Saturday
    expect(inTimePeriod('thisMonth', today - 1, today, false)).toBe(true);
    expect(inTimePeriod('lastMonth', today - 2, today, false)).toBe(true); // 2023-09-30
  });
});

describe('visual rule XML', () => {
  it('round-trips through parse and serialise', () => {
    const xml = '<dataBar showValue="0"><cfvo type="num" val="0"/><cfvo type="formula" val="MAX(&quot;a&quot;)"/><color rgb="FF638EC6"/></dataBar>';
    const v = parseVisualRule(makeCfRule({ type: 'dataBar', priority: 1, innerXml: xml }));
    expect(v).toMatchObject({ kind: 'dataBar', showValue: false, cfvos: [{ type: 'num', val: '0' }, { type: 'formula', val: 'MAX("a")' }] });
    if (!v) throw new Error('expected a data bar');
    expect(visualRuleXml(v)).toBe(xml);
  });

  it('treats a data bar with an x14 extension as an Excel 2010 bar growing from zero', () => {
    column([5, 10]);
    pushRule('A1:A2', { type: 'dataBar', priority: 1, innerXml: '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{1}</x14:id></ext></extLst>' });
    const at = overlay();
    expect(at(1, 1, undefined)?.bar?.end).toBeCloseTo(0.5);
    expect(at(2, 1, undefined)?.bar?.end).toBeCloseTo(1);
  });

  it('draws an Excel 2010 bar over only negatives leftward from an axis at the right edge', () => {
    column([-10, -5, -1]);
    pushRule('A1:A3', { type: 'dataBar', priority: 1, innerXml: '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{1}</x14:id></ext></extLst>' });
    const at = overlay();
    // The most negative value gets the longest red bar, all ending at the axis.
    expect([1, 2, 3].map((r) => at(r, 1, undefined)?.bar)).toEqual([
      { start: 0, end: 1, color: '#FF0000' },
      { start: 0.5, end: 1, color: '#FF0000' },
      { start: 0.9, end: 1, color: '#FF0000' },
    ]);
  });
});

describe('range helpers', () => {
  it('subtracts a hole into disjoint rectangles', () => {
    const parts = subtractRange({ r1: 1, c1: 1, r2: 5, c2: 5 }, { r1: 2, c1: 2, r2: 3, c2: 3 });
    const cells = parts.reduce((n, r) => n + (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1), 0);
    expect(cells).toBe(25 - 4);
  });

  it('parses Applies to text', () => {
    expect(parseAppliesTo('=$A$1:$B$2,$D:$D')).toEqual([
      { r1: 1, c1: 1, r2: 2, c2: 2 },
      { r1: 1, c1: 4, r2: 1_048_576, c2: 4 },
    ]);
    expect(parseAppliesTo('=nope!!')).toBeUndefined();
  });
});

describe('rule forms', () => {
  it('turns comparison input into operand formulas and back', () => {
    expect(operandFormula('5')).toBe('5');
    expect(operandFormula('=$B$1')).toBe('$B$1');
    expect(operandFormula('say "hi"')).toBe('"say ""hi"""');
    expect(operandText('"say ""hi"""')).toBe('say "hi"');
    expect(operandText('$B$1')).toBe('=$B$1');
  });

  it('round-trips every rule kind through the editor form', () => {
    const palette = new StyleResolver(wb).palette;
    const forms: Array<Partial<RuleForm>> = [
      { kind: 'values', style: 'scale3' },
      { kind: 'values', style: 'dataBar', barOnly: true },
      { kind: 'values', style: 'iconSet', iconSet: '5Arrows', reverse: true },
      { kind: 'contains', contains: 'cellValue', operator: 'notBetween', value1: '1', value2: '=$C$1' },
      { kind: 'contains', contains: 'text', textOp: 'endsWith', text: 'x' },
      { kind: 'contains', contains: 'date', period: 'lastMonth' },
      { kind: 'contains', contains: 'noErrors' },
      { kind: 'topBottom', bottom: true, percent: true, rank: 5 },
      { kind: 'average', average: 'std2Below' },
      { kind: 'average', average: 'equalBelow' },
      { kind: 'unique', unique: true },
      { kind: 'formula', formula: '=MOD(ROW(),2)=0' },
    ];
    for (const patch of forms) {
      const form = { ...defaultRuleForm(), ...patch };
      const made = makeCfRule({ ...ruleFromForm(form, 'B2'), priority: 1 });
      expect(formFromRule(made, palette)).toEqual(form);
    }
  });

  it('writes the companion formula Excel reads for text rules', () => {
    const made = ruleFromForm({ ...defaultRuleForm(), kind: 'contains', contains: 'text', textOp: 'containsText', text: 'a"b' }, 'C3');
    expect(made.formulas).toEqual(['NOT(ISERROR(SEARCH("a""b",C3)))']);
  });
});
