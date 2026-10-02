import { writeFileSync } from 'node:fs';
import { adapters, outputDirectory } from './output-setup.js';
import { required } from '../conformance/required.js';
import { validateXlsx } from '../conformance/validate.js';
import { expect, test } from '@playwright/test';
import { corpus } from '../conformance/corpus.js';
declare global {
  interface Window {
    qa: { exercise(url: string, adapter: string, streaming: boolean, sheetName: string): Promise<{ values: unknown[]; bytes?: number[]; font?: Record<string, unknown>; border?: string; edited?: string; date?: string; names?: unknown[]; date1904?: boolean }> };
  }
}
for (const c of corpus) {
  test(c.id, async ({ page }, testInfo) => {
    await page.goto('/');
    await page.waitForFunction(() => !!window.qa);
    for (const adapter of adapters) {
      for (const streaming of [false, true]) {
        const result = await page.evaluate(async ([id, method, mode, sheet]) => window.qa.exercise(`/cases/${id}.xlsx`, method, mode, sheet), [c.id, adapter, streaming, c.sheetName?.replace('&apos;', "'") ?? 'Audit'] as const);
        expect(result.values).toEqual((streaming ? c.expectedStreamValues : undefined) ?? c.expectedValues ?? [c.expectedValue ?? 'audit']);
        if (c.expectedDate1904 !== undefined) expect(result.date1904).toBe(c.expectedDate1904);
        if (!streaming) {
          const bytes = Uint8Array.from(required(result.bytes));
          const id = `${c.id}.${testInfo.project.name}.${adapter}`;
          writeFileSync(`${outputDirectory}/${id}.output.xlsx`, bytes);
          const validation = await validateXlsx(bytes);
          expect(validation.status, JSON.stringify(validation)).toBe('valid');
          if (c.expectedDate) expect(result.date).toBe(c.expectedDate);
          if (c.expectedNames) expect(result.names).toEqual(c.expectedNames);
          expect(result.edited).toBe('browser edit');
          if (c.expectedFont) expect(result.font).toMatchObject(c.expectedFont);
          for (const key of c.expectedMissingFont ?? []) expect(result.font?.[key]).toBeUndefined();
          if (c.expectedBorder) expect(result.border).toBe(c.expectedBorder);
        }
      }
    }
  });
}
