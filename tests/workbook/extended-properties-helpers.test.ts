// Tests for docProps/app.xml extended-property ergonomic helpers.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import {
  setWorkbookAppVersion,
  setWorkbookApplication,
  setWorkbookCompany,
  setWorkbookHyperlinkBase,
  setWorkbookManager,
} from '../../src/packaging/extended';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

describe('extended-properties helpers', () => {
  it('lazily allocate wb.appProperties and write each field', () => {
    const wb = createWorkbook();
    expect(wb.appProperties).toBeUndefined();
    setWorkbookCompany(wb, 'Anthropic');
    setWorkbookManager(wb, 'Alice');
    setWorkbookApplication(wb, '@office-kit/xlsx');
    setWorkbookAppVersion(wb, '0.1.0');
    setWorkbookHyperlinkBase(wb, 'https://docs.example.com/');
    expect(wb.appProperties).toEqual({
      company: 'Anthropic',
      manager: 'Alice',
      application: '@office-kit/xlsx',
      appVersion: '0.1.0',
      hyperlinkBase: 'https://docs.example.com/',
    });
  });

  it('subsequent calls overwrite', () => {
    const wb = createWorkbook();
    setWorkbookCompany(wb, 'Old');
    setWorkbookCompany(wb, 'New');
    expect(wb.appProperties?.company).toBe('New');
  });

  it('round-trips through saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setWorkbookCompany(wb, 'Anthropic');
    setWorkbookManager(wb, 'Alice');
    setWorkbookApplication(wb, '@office-kit/xlsx');
    setWorkbookHyperlinkBase(wb, 'https://example.com/');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.appProperties?.company).toBe('Anthropic');
    expect(wb2.appProperties?.manager).toBe('Alice');
    expect(wb2.appProperties?.application).toBe('@office-kit/xlsx');
    expect(wb2.appProperties?.hyperlinkBase).toBe('https://example.com/');
  });

  it('non-ASCII company / manager round-trip', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setWorkbookCompany(wb, '株式会社サンプル');
    setWorkbookManager(wb, '田中花子 🌸');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.appProperties?.company).toBe('株式会社サンプル');
    expect(wb2.appProperties?.manager).toBe('田中花子 🌸');
  });
});
