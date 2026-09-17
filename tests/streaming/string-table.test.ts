import { describe, expect, it } from 'vitest';
import { createWriteOnlyStringTable } from '../../src/streaming/string-table.js';
import { parseSharedStringsXml } from '../../src/workbook/shared-strings.js';

describe('bounded write-only string table', () => {
  it('honours the byte budget, then permanently falls back while retaining existing IDs', () => {
    const table = createWriteOnlyStringTable();
    // A plain value costs two UTF-16 copies plus the seven-character <t> wrapper.
    // Fill most of the budget with a valid Excel-length value, then the remainder.
    let used = 0;
    let last = '';
    for (let i = 0; i < 65; i++) {
      last = `${i}`.padEnd(32_000, 'x');
      expect(table.serialize(last).type).toBe('s');
      used += 4 * last.length + 14;
    }
    // One odd-length wrapper leaves two bytes unused, which cannot fit any entry.
    const remaining = Math.floor((8 * 1024 * 1024 - used - 14) / 4);
    const tail = 'y'.repeat(remaining);
    expect(table.serialize(tail).type).toBe('s');
    expect(table.serialize('').type).toBe('inlineStr');
    expect(table.serialize(last)).toEqual({ type: 's', xml: '<v>64</v>' });
    expect(table.serialize(tail)).toEqual({ type: 's', xml: '<v>65</v>' });
    expect(table.serialize('later').type).toBe('inlineStr');
    expect(table.size).toBe(66);
  });

  it('counts rich text against the same payload budget and snapshots mutable input', async () => {
    const table = createWriteOnlyStringTable();
    const run = { text: 'original', font: { b: true } };
    const runs = [run];
    expect(table.serialize({ kind: 'rich-text', runs }).type).toBe('s');
    run.text = 'changed';
    run.font.b = false;
    const chunks: Uint8Array[] = [];
    await table.write({ write: (chunk) => { chunks.push(chunk); }, end: async () => {} });
    const xml = chunks.map((chunk) => new TextDecoder().decode(chunk)).join('');
    expect(xml).toContain('<t>original</t>');
    expect(xml).not.toContain('changed');
    for (let i = 0; i < 200; i++) {
      table.serialize({ kind: 'rich-text', runs: [{ text: `${i}`.padEnd(32_000, 'x') }] });
    }
    expect(table.size).toBeLessThan(200);
    expect(table.serialize('plain after rich overflow').type).toBe('inlineStr');
    expect(table.serialize({ kind: 'rich-text', runs: [{ text: 'original', font: { b: true } }] }))
      .toEqual({ type: 's', xml: '<v>0</v>' });
  });

  it('does not retain a single oversized value or reopen admission for smaller values', () => {
    const table = createWriteOnlyStringTable();
    expect(table.serialize('x'.repeat(4 * 1024 * 1024)).type).toBe('inlineStr');
    expect(table.serialize('small').type).toBe('inlineStr');
    expect(table.size).toBe(0);
  });

  it('streams the final table in small UTF-8 chunks without splitting surrogate pairs', async () => {
    const table = createWriteOnlyStringTable();
    const value = '😀あ'.repeat(10_000);
    table.serialize(value);
    // Cell text escapes surrogates before chunking; font attributes do not.
    // Shifting the prefix covers both alignments of a surrogate pair.
    const rich = [0, 1].map((n) => ({ kind: 'rich-text' as const,
      runs: [{ text: 'font', font: { name: 'x'.repeat(n) + '😀'.repeat(10_000) } }],
    }));
    for (const entry of rich) table.serialize(entry);
    const chunks: Uint8Array[] = [];
    await table.write({ write: (chunk) => { chunks.push(chunk); }, end: async () => {} });
    expect(chunks.length).toBeGreaterThan(1);
    expect(Math.max(...chunks.map((c) => c.length))).toBeLessThanOrEqual(64 * 1024);
    const xml = chunks.map((chunk) => new TextDecoder('utf-8', { fatal: true }).decode(chunk)).join('');
    expect(parseSharedStringsXml(xml).entries).toEqual([value, ...rich]);
  });
});
