---
'@office-kit/xlsx': minor
---

fix: a split or frozen sheet lost all but its first `<selection>`, so after a round-trip Excel reopened on the frozen corner instead of the selected cell. Breaking: `SheetView.selection` is now `SheetView.selections` (one per pane); `activeSelection(view)` returns the active pane's.
