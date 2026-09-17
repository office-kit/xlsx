---
"@office-kit/xlsx": minor
---

Load the cell shapes a current Excel writes. A `<c t="e">` carrying one of the
nine error tokens Excel added since 2018 (`#SPILL!`, `#CALC!`, `#FIELD!`,
`#BLOCKED!`, `#CONNECT!`, `#BUSY!`, `#UNKNOWN!`, `#PYTHON!`, `#EXTERNAL!`) threw
`unknown error code`, and a `<c t="d">` holding an ISO 8601 date threw
`unknown cell type t="d"`. One such cell anywhere in a workbook aborted the
whole load, so a file with thousands of good rows failed over a single spilled
formula.

`ERROR_CODES` now lists all seventeen tokens, and a `t="e"` token outside that
set is kept verbatim rather than dropped: `t="e"` is the file declaring the cell
an error, Excel keeps adding tokens, and a load then save re-emits the token
unchanged. A `t="d"` cell reads as a `Date`, or as a duration for the time-only
and `PT…` forms, in UTC like every other `Date` in the model. Writing is
unchanged: a `Date` still saves as a serial number under the workbook epoch, not
as `t="d"`.

**This changes `ExcelErrorCode`** from a union of eight literals to
`` `#${string}` ``. Code that passes the type around, or into `makeErrorValue`,
is unaffected; an exhaustive `switch` over the eight literals in your own code
no longer type-checks as exhaustive.

`loadWorkbookStream` answers all of these the way `loadWorkbook` does. It
previously read an unlisted error token as an empty cell and ran a `t` outside
`ST_CellType` through `Number.parseFloat`, so the same file could load as empty
cells in one reader and throw in the other.
