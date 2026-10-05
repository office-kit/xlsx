---
'@office-kit/xlsx': minor
---

fix: AutoFilter custom conditions, Top 10 and dynamic (e.g. above average) filters were dropped on load/save; colour and icon filters are now kept verbatim. Breaking: `FilterColumn` gains `custom`, `top10`, `dynamic` and `raw` variants, so code reading `values` must first check `kind === 'filters'`.
