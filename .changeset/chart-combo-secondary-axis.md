---
'@office-kit/xlsx': minor
---

feat(chart): write and read combo charts. `PlotArea.additionalCharts` holds further chart groups drawn in the same plot area (columns plus a line, or a series on a secondary value axis), and `additionalCatAx` / `additionalValAx` define the axes they use. Loading a combo chart made in Excel now keeps every chart group instead of only the first one. `chartToBytes` throws `OpenXmlSchemaError` when an additional chart group names an axis the plot area doesn't define.

feat(chart): `Legend.legendEntry` hides (`{ idx, delete: true }`) or restyles (`{ idx, txPr }`) the legend row of a single series, and survives load and save.
