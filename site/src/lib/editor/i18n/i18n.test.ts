import { readFileSync, readdirSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

// The message tables are spread together, so a key defined in two files is
// silently overridden by whichever comes later. Keep every key in one place.
describe('message tables', () => {
  for (const lang of ['en', 'ja']) {
    it(`define each ${lang} key once`, () => {
      const dir = new URL('.', import.meta.url);
      const files = readdirSync(dir).filter((f) => f === `${lang}.ts` || f.endsWith(`.${lang}.ts`));
      const owner = new Map<string, string>();
      const clashes: string[] = [];
      for (const f of files) {
        for (const m of readFileSync(new URL(f, dir), 'utf8').matchAll(/^ {2}([A-Za-z0-9_]+):/gm)) {
          const key = m[1] ?? '';
          const prev = owner.get(key);
          if (prev) clashes.push(`${key} (${prev}, ${f})`);
          else owner.set(key, f);
        }
      }
      expect(clashes).toEqual([]);
    });
  }
});
