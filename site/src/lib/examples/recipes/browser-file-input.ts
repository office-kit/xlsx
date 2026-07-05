// Browser: parse the xlsx the user just selected via <input type="file">.
// fromBlob is streaming, so the workbook starts parsing while the file
// is still being read.

import { fromBlob, loadWorkbook } from '@office-kit/xlsx/io';

export async function loadFromInput(input: HTMLInputElement) {
  const file = input.files?.[0];
  if (!file) return null;
  const wb = await loadWorkbook(fromBlob(file));
  return wb;
}
