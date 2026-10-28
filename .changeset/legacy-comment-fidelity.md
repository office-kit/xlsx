---
"@office-kit/xlsx": minor
---

fix: notes (legacy comments) lost their formatting and their box on save (#307). A loaded note now keeps its rich-text runs in the new `LegacyComment.richText`, and its box's `anchor` (position and size) and `visible` (shown, not only on hover) from the VML drawing; all three are written back. `text` stays the note's content: runs are written only while they join to exactly `text`, and `editCommentText` drops them. `makeLegacyComment`, `setComment` and `setComments` accept the new fields. `makeLegacyComment` throws when `richText` does not join to `text`, and saving throws when an `anchor` value is not a non-negative integer. Comment text now decodes and encodes `_xHHHH_` escapes as cell strings already do, so a note's line breaks are written as `_x000A_` (Excel and LibreOffice read them back as line breaks).
