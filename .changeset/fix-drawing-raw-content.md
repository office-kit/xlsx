---
'@office-kit/xlsx': patch
---

fix: drawing shapes were rewritten as chart graphicFrames with an empty r:id (#110)

An anchor holding anything the model doesn't cover — a shape, a group, a
connector — was written back as a chart `<xdr:graphicFrame>` carrying
`<c:chart r:id=""/>`, so a rectangle came back as "Chart 1" with a dangling
reference, its geometry gone and no drawing rels part written at all. Anchors
sitting inside Excel's `<mc:AlternateContent>` wrapper were dropped outright,
leaving an empty `<xdr:wsDr/>`. Together those are what still broke worksheets
carrying form controls after #105: the worksheet side was fixed, but the shapes
that draw the controls live in `drawing1.xml`.

Unmodeled drawing content is now kept as the verbatim source XML and written
back untouched — the whole anchor element, so `editAs` and the
`<xdr:clientData fLocksWithSheet="0">` flags form controls rely on survive too
— and `<mc:AlternateContent>` wrappers keep their place in document order,
which is z-order. Relationships those nodes reference keep their original ids
and their target parts, and the writer allocates ids for modeled charts and
pictures around them.
