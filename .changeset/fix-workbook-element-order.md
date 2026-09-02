---
'@office-kit/xlsx': patch
---

fix: workbook.xml was written out of CT_Workbook order (#112)

`CT_Workbook` (ECMA-376 §18.2.27) is an `xsd:sequence`, so the children of
`<workbook>` have a normative order. Saving a workbook wrote `<definedNames>`
straight after `<sheets>` — ahead of `<functionGroups>` and
`<externalReferences>` — and `<pivotCaches>` ahead of `<calcPr>`. A workbook
carrying both an external reference and a defined name therefore came back as
a package Excel refuses to open.

`workbook.xml` is now emitted in schema order: `functionGroups`,
`externalReferences`, `definedNames`, `calcPr`, `oleSize`,
`customWorkbookViews`, `pivotCaches`, `smartTagPr`, `smartTagTypes`,
`fileRecoveryPr`, then any unmodeled tail (`webPublishing`,
`webPublishObjects`, `extLst`) that was carried over from the loaded file.
