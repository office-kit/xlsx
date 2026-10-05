---
'@office-kit/xlsx': patch
---

perf: `mergeCells` and `unmergeCells` compare ranges as numbers instead of building an A1 string for every existing merge, so adding many merges to a sheet is several times faster (10,000 merges: 2.5 s → 0.6 s).
