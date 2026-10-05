---
'@office-kit/xlsx': patch
---

fix: table column names with a line break and table column formulas survived load and save. Excel writes a line break in a header as `_x000a_`, and the name was kept with that sequence, so it no longer matched the header cell. Calculated-column (`<calculatedColumnFormula>`) and custom total-row (`<totalsRowFormula>`) formulas were dropped on load and never written.
