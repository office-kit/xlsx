---
'@office-kit/xlsx': minor
---

fix: a rotated pie chart (Angle of first slice) lost its rotation on load/save. `PieChart.firstSliceAng` now round-trips, as it already did for doughnuts.
