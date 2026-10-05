---
'@office-kit/xlsx': minor
---

fix: AutoFilter custom conditions, Top 10 and dynamic (e.g. above average) filters were dropped on load/save; colour and icon filters are now kept verbatim. The sort Excel stores inside `<autoFilter>` and in table parts is kept too (`AutoFilter.sortState`, `TableDefinition.sortState`). Breaking: `FilterColumn` gains `custom`, `top10`, `dynamic` and `raw` variants, so code reading `values` must first check `kind === 'filters'`.
