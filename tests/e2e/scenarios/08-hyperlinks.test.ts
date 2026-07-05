// Scenario 08: external + internal hyperlinks.
// Output: 08-hyperlinks.xlsx
//
// What to verify in Excel:
// - A1 is clickable, opens https://github.com/office-kit/xlsx
// - A2 is an in-workbook hyperlink jumping to "Target!A1"
// - A3 has a tooltip ("Hover me") visible on mouse-over
// - The "Target" sheet has the destination cell at A1.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../../src/workbook/index';
import { setCell } from '../../../src/worksheet/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 08 — hyperlinks', () => {
  it('writes 08-hyperlinks.xlsx', async () => {
    const wb = createWorkbook();
    const main = addWorksheet(wb, 'Main');
    const target = addWorksheet(wb, 'Target');
    setCell(target, 1, 1, '↑ jumped here from Main!A2');

    setCell(main, 1, 1, '@office-kit/xlsx on GitHub');
    setCell(main, 2, 1, 'jump to Target sheet');
    setCell(main, 3, 1, 'with tooltip');

    main.hyperlinks.push({
      ref: 'A1',
      target: 'https://github.com/office-kit/xlsx',
      tooltip: 'Click to open the project page',
    });
    main.hyperlinks.push({
      ref: 'A2',
      location: 'Target!A1',
      display: 'Target!A1',
    });
    main.hyperlinks.push({
      ref: 'A3',
      target: 'https://example.com',
      tooltip: 'Hover me',
    });

    const result = await writeWorkbook('08-hyperlinks.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
