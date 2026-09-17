---
"@office-kit/xlsx": minor
---

Add `contentLimits` to `loadWorkbook` and `loadWorkbookStream`: a cap on the
cells and rows a read will model. `decompressionLimits` bounds the bytes an
archive inflates to, which is not the quantity that decides what a read costs.
A 2 MB upload can inflate to a few hundred MB of `<sheetData>` while staying
inside every byte default, and the model used to be built for every cell of it
before the caller could look at anything.

```ts
const wb = await loadWorkbook(source, {
  contentLimits: { maxCells: 1_000_000, maxRows: 100_000 },
});
```

Enforced inside the cell loop, so a workbook past the cap is refused in the
time it takes to read the cells up to it. Exceeding either cap throws the new
`OpenXmlContentLimitError`, exported from `@office-kit/xlsx/utils`, which names
the cap and the cell or row that reached it. `ContentLimits` is exported from
`@office-kit/xlsx/worksheet`.

Unlimited by default, so nothing changes for a caller that does not ask. The
counts cover one pass over the content: every worksheet of the workbook for
`loadWorkbook`, and one row-iteration for `loadWorkbookStream`, which holds a
row at a time and can be iterated again.

`SECURITY.md` carries the recommended ingestion profile.
