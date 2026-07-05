// One-shot read + save direct from / to disk via the @office-kit/xlsx/node
// helpers, no manual fs glue needed.

import { loadWorkbook, saveWorkbook } from '@office-kit/xlsx/io';
import { fromFile, toFile } from '@office-kit/xlsx/node';

const wb = await loadWorkbook(fromFile('input.xlsx'));
// ...mutate wb...
await saveWorkbook(wb, toFile('output.xlsx'));
