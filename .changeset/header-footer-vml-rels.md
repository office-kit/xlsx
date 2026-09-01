---
'@office-kit/xlsx': patch
---

fix: round-tripping a sheet with a header/footer picture (`<legacyDrawingHF>`) produced a file Excel refused to open (#104)

`loadWorkbook` → `saveWorkbook` re-emitted `<legacyDrawingHF r:id="…"/>` but dropped the `vmlDrawing` relationship it points at, the VML part's own `.rels`, and the image behind it, and the surviving `.vml` part had no content type. Parts referenced by relationships the writer re-emits verbatim are now carried over together with their own relationships (transitively), and `<Default>` content types from the source manifest are preserved for such parts.
