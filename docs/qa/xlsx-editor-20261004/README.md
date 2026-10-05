# XLSX editor QA screenshots, 2026-10-04

Taken with Playwright (Chromium, 1280×800) against `vite dev` of `site/`;
no native Excel involved.

| File | Shows | Source it was taken from (sha256, first 12) |
| --- | --- | --- |
| `remove-duplicates-before.png` | table `Sales` A1:B6 with East / West repeated, sentinels at D3 and A9 | `site/src/lib/editor/core/data.ts` dc8447acc12c, `dialogs/RemoveDuplicatesDialog.svelte` b2e11b07b9f4 |
| `remove-duplicates-after.png` | Table ▸ Remove Duplicates on Region: three rows left, the table shrunk to A1:B4, sentinels unchanged | same |
| `fill-handle-before.png` | A1:A2 = 1, 2 selected; the fill handle found by the crosshair cursor | `site/src/lib/editor/grid/GridView.svelte` be7050e23214 |
| `fill-handle-after.png` | after dragging the fill handle down: A3:A5 = 3, 4, 5 (Undo empties them again) | same |
