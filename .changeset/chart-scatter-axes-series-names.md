---
'@office-kit/xlsx': minor
---

fix: charts kept their series names and scatter / bubble charts kept both value axes through load and save. Series names (`<c:tx>`) were dropped on load, so legends showed "Series1". The X value axis of a scatter chart was lost and replaced by a default one sharing the Y axis's id. It is now read into the new `PlotArea.xValAx`. Typed series names are written as `<c:v>` instead of an empty-formula reference.
