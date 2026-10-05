---
'@office-kit/xlsx': minor
---

feat: model drawing shapes, text boxes and connectors, and worksheet sparklines

- `DrawingItem.content` gains a `shape` kind for `<xdr:sp>` (including text boxes, `textBox: true`) and `<xdr:cxnSp>` (`connector: true`): preset or custom geometry, fill, outline, the theme `style` reference and the text body all round-trip, and `makeShapeDrawingItem` builds one. Shapes the model can't carry faithfully — with a macro, a hyperlink, shape locks or `<xdr:clientData>` flags — are still kept verbatim as `unsupported`. Code that switches on `content.kind` should handle the new `shape` variant; such shapes were previously reported as `unsupported`.
- `Worksheet.sparklineGroups` reads and writes Excel 2010 sparklines (`<x14:sparklineGroups>`): line / column / win-loss (`stacked`) type, the high / low / negative / first / last / markers flags, colours, axis settings and the data-range ↔ location pairs. `makeSparklineGroup` applies Excel's default colours. The field is optional (`addWorksheet` sets it to `[]`), so older `Worksheet` literals still compile. Other worksheet `<extLst>` entries are preserved, and the `<extLst>` is now written as the last worksheet child, where the schema requires it.
- Breaking for exhaustive matches: a `switch` on `DrawingItem.content.kind` that ends in a `never` check stops compiling until it adds `case 'shape':`. Nothing else in the public types was removed or narrowed.
- `@office-kit/xlsx/drawing` now also exports `makePresetGeometry`, `PRESET_SHAPE_NAMES`, `makeLine` and the geometry / line types.
