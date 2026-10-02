// Bounded adversarial inputs are hand-authored; no library writer defines them.
import { strToU8, zipSync } from 'fflate';
import { loadWorkbook } from '../dist/io.mjs';
import { fromBuffer } from '../dist/node.mjs';
import { fromStream } from '../dist/io.mjs';
import { loadWorkbookStream } from '../dist/streaming.mjs';
import { OpenXmlContentLimitError, OpenXmlDecompressionBombError, OpenXmlSchemaError } from '../dist/utils.mjs';
import { getCell } from '../dist/worksheet.mjs';
const [id, mode] = process.argv.slice(2);
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const options = {};
let rows = '<row r="1"><c r="A1"><v>7</v></c></row>';
let before = ''; let padding = ''; let dimension = 'A1'; let ErrorClass;
if (id === 'inflation') { padding = `<!--${'x'.repeat(8 * 1024 * 1024)}-->`; options.decompressionLimits = { maxEntryUncompressedBytes: 1024 * 1024 }; ErrorClass = OpenXmlDecompressionBombError; }
else if (id === 'rows') { rows = Array.from({ length: 100_000 }, (_, i) => `<row r="${i + 1}"/>`).join(''); dimension = 'A1:A100000'; options.contentLimits = { maxRows: 20 }; ErrorClass = OpenXmlContentLimitError; }
else if (id === 'cells') { rows = Array.from({ length: 50_000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>1</v></c><c r="B${i + 1}"><v>2</v></c></row>`).join(''); dimension = 'A1:B50000'; options.contentLimits = { maxCells: 20 }; ErrorClass = OpenXmlContentLimitError; }
else if (id === 'doctype') { before = '<!DOCTYPE worksheet [<!ENTITY a "blocked">]>'; ErrorClass = OpenXmlSchemaError; }
else if (id === 'dimension') { dimension = 'A1:XFD1048576'; options.contentLimits = { maxCells: 1, maxRows: 1 }; }
else throw new Error(`Unknown resource case: ${id}`);
if (!['model', 'stream'].includes(mode)) throw new Error(`Unknown loader: ${mode}`);
const xml = {
  '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
  '_rels/.rels': `<Relationships xmlns="${REL}"><Relationship Id="office" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  'xl/workbook.xml': `<workbook xmlns="${SML}" xmlns:r="${R}"><sheets><sheet name="Audit" sheetId="1" r:id="sheet"/></sheets></workbook>`,
  'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${REL}"><Relationship Id="sheet" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  'xl/worksheets/sheet1.xml': `${before}<worksheet xmlns="${SML}"><dimension ref="${dimension}"/><sheetData>${rows}</sheetData>${padding}</worksheet>`,
};
const bytes = zipSync(Object.fromEntries(Object.entries(xml).map(([path, text]) => [path, strToU8(text)])));
let offset = 0;
const source = mode === 'model' ? fromBuffer(bytes) : fromStream(new ReadableStream({ pull(controller) {
  if (offset === bytes.length) controller.close();
  else { controller.enqueue(bytes.subarray(offset, offset + 7)); offset = Math.min(bytes.length, offset + 7); }
} }));
let outcome;
try {
  if (mode === 'model') {
    const wb = await loadWorkbook(source, options);
    const ref = wb.sheets[0];
    if (ref?.kind !== 'worksheet' || getCell(ref.sheet, 1, 1)?.value !== 7) throw new Error('Wrong positive-control value');
  } else {
    const wb = await loadWorkbookStream(source, options);
    try {
      const values = [];
      for await (const row of wb.openWorksheet('Audit').iterValues()) values.push(...row);
      if (JSON.stringify(values) !== '[7]') throw new Error('Wrong positive-control stream');
    } finally { await wb.close(); }
  }
  if (ErrorClass) throw new Error(`Expected ${ErrorClass.name}, input was accepted`);
  outcome = 'accepted';
} catch (error) {
  if (!ErrorClass || !(error instanceof ErrorClass)) throw error;
  outcome = error.name;
}
// maxRSS is KiB on supported Node platforms, per process.resourceUsage documentation.
console.info(JSON.stringify({ id, mode, outcome, inputBytes: bytes.length, peakRssBytes: process.resourceUsage().maxRSS * 1024 }));
