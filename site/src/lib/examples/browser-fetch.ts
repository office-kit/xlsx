// Browser: pipe a fetch Response straight into the loader. fromResponse is
// streaming, so the workbook starts parsing before the download is done.

import { fromResponse, loadWorkbook } from '@office-kit/xlsx/io';

const response = await fetch('/sheet.xlsx');
const wb = await loadWorkbook(fromResponse(response));
const ref = wb.sheets[0];
if (ref?.kind === 'worksheet') {
  console.log(ref.sheet.title);
}
