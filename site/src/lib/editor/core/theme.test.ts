import { describe, expect, it } from 'vitest';
import { themeFontsFromXml } from './theme.ts';

describe('themeFontsFromXml', () => {
  it("reads the theme's Latin heading and body fonts", () => {
    const xml = new TextEncoder().encode(
      '<a:theme><a:themeElements><a:fontScheme name="Office">' +
        '<a:majorFont><a:latin typeface="Aptos Display" panose=""/><a:ea typeface=""/></a:majorFont>' +
        '<a:minorFont><a:latin typeface="Aptos Narrow" panose=""/><a:ea typeface=""/></a:minorFont>' +
        '</a:fontScheme></a:themeElements></a:theme>',
    );
    expect(themeFontsFromXml(xml)).toEqual({ major: 'Aptos Display', minor: 'Aptos Narrow' });
  });

  it('falls back to the Office theme fonts without a theme part', () => {
    expect(themeFontsFromXml(undefined)).toEqual({ major: 'Calibri Light', minor: 'Calibri' });
  });
});
