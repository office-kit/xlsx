import { describe, expect, expectTypeOf, it } from 'vitest';
import type {
  ColorMod,
  DmlColor,
  DmlColorWithMods,
  Fill,
  ParagraphProperties,
  RunProperties,
  SchemeColorName,
  TextBody,
  TextParagraph,
  TextRun,
} from '../../src/drawing';
import {
  makeBreak,
  makeColor,
  makeGradientFill,
  makeNoFill,
  makeParagraph,
  makePatternFill,
  makeRun,
  makeRunProperties,
  makeSchemeColor,
  makeSimpleTextBody,
  makeSolidFill,
  makeSrgbColor,
  makeTextBody,
  PRESET_PATTERN_NAMES,
  SCHEME_COLOR_NAMES,
} from '../../src/drawing';

describe('issues #55 + #56 — DML colour + fill + text helpers re-export from @office-kit/xlsx/drawing', () => {
  it('exposes makeSrgbColor / makeSchemeColor / makeColor as values', () => {
    const srgb = makeSrgbColor('FF0000');
    expect(srgb).toEqual({ kind: 'srgb', value: 'FF0000' });
    const scheme = makeSchemeColor('accent1');
    expect(scheme).toEqual({ kind: 'schemeClr', value: 'accent1' });
    const wrapped = makeColor(srgb);
    expect(wrapped.base).toEqual(srgb);
    expect(wrapped.mods).toEqual([]);
  });

  it('exposes the colour type aliases', () => {
    expectTypeOf<DmlColor>().toMatchTypeOf<{ kind: string }>();
    expectTypeOf<SchemeColorName>().toMatchTypeOf<string>();
    const mod: ColorMod = { kind: 'tint', val: 50000 };
    const wrapped: DmlColorWithMods = { base: makeSrgbColor('00FF00'), mods: [mod] };
    expect(wrapped.mods[0]?.kind).toBe('tint');
  });

  it('exposes Fill helpers and shape', () => {
    const solid = makeSolidFill(makeColor(makeSrgbColor('0000FF')));
    expect(solid.kind).toBe('solidFill');
    const none = makeNoFill();
    expect(none.kind).toBe('noFill');
    const grad = makeGradientFill({
      stops: [
        { pos: 0, color: makeColor(makeSrgbColor('FFFFFF')) },
        { pos: 100000, color: makeColor(makeSrgbColor('000000')) },
      ],
    });
    expect(grad.kind).toBe('gradFill');
    const patt = makePatternFill({ preset: 'pct50' });
    expect(patt.kind).toBe('pattFill');
    // Type-only assertion: Fill is the discriminated union of fill kinds.
    const _f: Fill = solid;
    expect(_f).toBe(solid);
  });

  it('exposes the preset / scheme constants', () => {
    expect(SCHEME_COLOR_NAMES.length).toBeGreaterThan(0);
    expect(PRESET_PATTERN_NAMES.length).toBeGreaterThan(0);
  });

  it('exposes TextBody / Paragraph / Run helpers and types', () => {
    const rPr: RunProperties = makeRunProperties({ sz: 900, b: true });
    expect(rPr.sz).toBe(900);
    const run: TextRun = makeRun('hello', rPr);
    expect(run.kind === 'r' && run.t).toBe('hello');
    const br: TextRun = makeBreak();
    expect(br.kind).toBe('br');
    const pPr: ParagraphProperties = { defRPr: rPr };
    const paragraph: TextParagraph = makeParagraph([run], pPr);
    expect(paragraph.runs.length).toBe(1);
    const body: TextBody = makeTextBody([paragraph]);
    expect(body.paragraphs.length).toBe(1);
    const simple = makeSimpleTextBody('axis label', rPr);
    expect(simple.paragraphs[0]?.runs.length).toBe(1);
  });
});
