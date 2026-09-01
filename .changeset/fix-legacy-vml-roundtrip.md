---
'@office-kit/xlsx': patch
---

fix: round-tripping a sheet with form controls produced a package Excel refused (#105)

Loading a worksheet that carries a form control (`<legacyDrawing>` + the
x14-gated `<controls>` block Excel 2010+ writes) and saving it again wrote a
file Excel reported as corrupt:

- `<legacyDrawing r:id>` and its `vmlDrawing` relationship were dropped on
  sheets without comments — only the comment VML was ever re-linked — leaving
  the control's VML shape orphaned;
- Excel's `<mc:AlternateContent><mc:Choice Requires="x14">` wrapper around
  `<controls>` / `<oleObjects>` was passed through as an opaque node, landing
  before `<drawing>` (out of `CT_Worksheet` order) with the `x14` prefix
  undeclared.

`Worksheet` now exposes `legacyDrawingRId` next to `legacyDrawingHFRId`; any
VML that is not a comment overlay keeps its relationship and part across a
round-trip, and the `Requires="x14"` wrapper is read into the typed
`controls` / `oleObjects` model and written back in schema order with
`xmlns:x14` declared on the worksheet root.
