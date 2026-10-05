import { expect, it } from 'vitest';
import { cssFontFamily } from './render-style.ts';

it.each([
  ['游明朝', 'serif'],
  ['ＭＳ 明朝', 'serif'],
  ['ＭＳ ゴシック', 'monospace'],
  ['Consolas', 'monospace'],
  ['メイリオ', 'sans-serif'],
])('%s falls back to a font of the same kind (%s)', (name, generic) => {
  expect(cssFontFamily(name).endsWith(generic)).toBe(true);
});

it('the Japanese and English names of a font share its substitutes', () => {
  expect(cssFontFamily('游明朝')).toBe(cssFontFamily('Yu Mincho'));
  expect(cssFontFamily('メイリオ')).toBe(cssFontFamily('Meiryo'));
});
