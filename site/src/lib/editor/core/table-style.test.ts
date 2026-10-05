import type { TableDefinition } from '@office-kit/xlsx/worksheet';
import { describe, expect, it } from 'vitest';
import { BUILT_IN_TABLE_STYLES, tableLooks } from './table-style.ts';

// Office 2023 theme, the palette Excel rendered the reference values with.
const PALETTE = ['FFFFFF', '000000', 'E8E8E8', '0E2841', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E', '467886', '96607D'];

function table(style: string, extra: Partial<TableDefinition> = {}): TableDefinition {
  return { id: 1, displayName: 'T', ref: 'B2:F8', columns: [], styleInfo: { name: style, showRowStripes: true }, ...extra };
}

describe('tableLooks', () => {
  it('lists the 60 built-in styles', () => {
    expect(BUILT_IN_TABLE_STYLES).toHaveLength(60);
    expect(BUILT_IN_TABLE_STYLES[22]).toBe('TableStyleMedium2');
  });

  it('paints TableStyleMedium2 like Excel: accent header, banded body', () => {
    const look = tableLooks([table('TableStyleMedium2')], PALETTE);
    expect(look?.(2, 3)).toMatchObject({ fill: '#156082', color: '#FFFFFF', bold: true });
    // Excel shows C0E6F5 for accent 1 lightened 80%.
    expect(look?.(3, 3)?.fill).toMatch(/^#C[01]E[56]F5$/);
    expect(look?.(4, 3)?.fill).toBeUndefined();
    expect(look?.(9, 3)).toBeUndefined();
  });

  it('turns banding off with the Banded Rows option', () => {
    const look = tableLooks([table('TableStyleMedium2', { styleInfo: { name: 'TableStyleMedium2', showRowStripes: false } })], PALETTE);
    expect(look?.(3, 3)?.fill).toBeUndefined();
  });

  it('uses the neutral shades for the first style of a run', () => {
    const look = tableLooks([table('TableStyleLight1')], PALETTE);
    expect(look?.(3, 3)?.fill).toBe('#D9D9D9');
    expect(look?.(2, 3)?.bold).toBe(true);
  });

  it('formats the total row and edge columns', () => {
    const def = table('TableStyleDark2', { totalsRowCount: 1, styleInfo: { name: 'TableStyleDark2', showFirstColumn: true } });
    const look = tableLooks([def], PALETTE);
    expect(look?.(8, 4)).toMatchObject({ bold: true, color: '#FFFFFF' });
    expect(look?.(8, 4)?.top?.width).toBe(2);
    expect(look?.(5, 2)).toMatchObject({ bold: true });
  });

  it('ignores tables without a known style', () => {
    expect(tableLooks([table('MyCustomStyle')], PALETTE)).toBeUndefined();
  });
});
