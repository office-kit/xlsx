import { expect, test } from '@playwright/test';
import { corpus } from '../conformance/corpus.js';
declare global {
  interface Window {
    qa: { exercise(url: string, adapter: string, streaming: boolean): Promise<{ values: unknown[]; font?: Record<string, unknown>; border?: string; edited?: string }> };
  }
}
for (const c of corpus) {
  test(c.id, async ({ page }) => {
    await page.goto('/');
    await page.waitForFunction(() => !!window.qa);
    for (const adapter of ['arrayBuffer', 'blob', 'response', 'stream']) {
      for (const streaming of [false, true]) {
        const result = await page.evaluate(async ([id, method, mode]) => window.qa.exercise(`/cases/${id}.xlsx`, method, mode), [c.id, adapter, streaming] as const);
        expect(result.values).toEqual([c.expectedValue ?? 'audit']);
        if (!streaming) {
          expect(result.edited).toBe('browser edit');
          if (c.expectedFont) expect(result.font).toMatchObject(c.expectedFont);
          for (const key of c.expectedMissingFont ?? []) expect(result.font?.[key]).toBeUndefined();
          if (c.expectedBorder) expect(result.border).toBe(c.expectedBorder);
        }
      }
    }
  });
}
