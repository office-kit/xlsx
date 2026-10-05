---
'@office-kit/xlsx': patch
---

fix: Excel no longer asks to repair workbooks with histogram, Pareto, box & whisker, waterfall, funnel, treemap or sunburst charts

- Each `cx:` chart is now saved with its chart style and color parts. Excel refused a chartex chart without them and removed it on repair.
- Histogram and Pareto bin settings are written the way the schema defines them. `binCount` / `binSize` become child elements, and automatic binning writes nothing. They used to be written as attributes, which Excel rejected.
- The box & whisker quartile method is written as `<cx:statistics quartileMethod="…"/>`. It used to be `<cx:quartileMethod val="…"/>`, which Excel rejected.
- Re-saving a workbook that Excel wrote with one of these charts no longer drops the `xmlns:cx1` declaration that `<mc:Choice Requires="cx1">` refers to. The same applies to any prefix named in `Requires`, `mc:Ignorable`, `mc:MustUnderstand` or `mc:ProcessContent` inside content kept verbatim.
