# Genuine Excel Strict fixtures

These files are unmodified Apache POI test data, introduced as Excel 2013
Strict workbooks in commit
[b5b76b5a12db548701a57a47adfaef42bba8d97d](https://github.com/apache/poi/commit/b5b76b5a12db548701a57a47adfaef42bba8d97d).
They are distributed under the accompanying Apache 2.0 license and POI notice.

- [sample.strict.xlsx](https://github.com/apache/poi/blob/b5b76b5a12db548701a57a47adfaef42bba8d97d/test-data/spreadsheet/sample.strict.xlsx): three sheets, shared and rich strings, a SUM formula, styles, theme, calc chain and printer settings.
- [SimpleStrict.xlsx](https://github.com/apache/poi/blob/b5b76b5a12db548701a57a47adfaef42bba8d97d/test-data/spreadsheet/SimpleStrict.xlsx): two sheets, formulas and an ISO date (`1990-01-01`, sheet 2 A4).

Additional date-system and malformed-input cases mutate copies in memory;
the committed fixtures remain byte-identical to upstream. Chart/drawing tests
use explicitly synthetic namespace variants, not claimed Excel Strict output.
