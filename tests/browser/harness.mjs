// Exercise the built distribution in a real browser, including its dependencies.
import { fromArrayBuffer, fromBlob, fromResponse, fromStream, loadWorkbook, workbookToBytes } from '../../dist/io.mjs';
import { loadWorkbookStream } from '../../dist/streaming.mjs';
import { getSheet } from '../../dist/workbook.mjs';
import { getCell, setCell } from '../../dist/worksheet.mjs';
import { getCellFont, getCellBorder, getCellDate } from '../../dist/styles.mjs';
export async function exercise(url, adapter, streaming, sheetName = 'Audit') {
  const response = await fetch(url);
  const bytes = new Uint8Array(await response.arrayBuffer());
  let offset = 0;
  const chunked = () => new ReadableStream({ pull(c) {
    if (offset >= bytes.length) { c.close(); return; }
    c.enqueue(bytes.slice(offset, offset + 7)); offset += 7;
  } });
  const source = adapter === 'blob' ? fromBlob(new Blob([bytes]))
    : adapter === 'response' ? fromResponse(new Response(chunked()))
    : adapter === 'stream' ? fromStream(chunked()) : fromArrayBuffer(bytes);
  if (streaming) {
    const wb = await loadWorkbookStream(source);
    const values = [];
    try { for await (const row of wb.openWorksheet(sheetName).iterRows()) values.push(canonical(row[0]?.value ?? null)); }
    finally { await wb.close(); }
    return { values, date1904: wb.date1904 };
  }
  const wb = await loadWorkbook(source);
  setCell(getSheet(wb, sheetName), 2, 2, 'browser edit');
  const saved = await workbookToBytes(wb);
  const reloaded = await loadWorkbook(fromArrayBuffer(saved));
  const ws = getSheet(reloaded, sheetName);
  const cell = getCell(ws, 1, 1);
  return { bytes: Array.from(saved), values: Array.from({ length: getCell(ws, 2, 1) ? 2 : 1 }, (_, i) => canonical(getCell(ws, i + 1, 1)?.value ?? null)), date: getCellDate(reloaded, cell)?.toISOString(), names: reloaded.definedNames, date1904: reloaded.date1904, font: getCellFont(reloaded, cell), border: getCellBorder(reloaded, cell).left?.style, edited: getCell(ws, 2, 2)?.value };
}

function canonical(value) { return value instanceof Date ? { date: value.toISOString() } : value; }
