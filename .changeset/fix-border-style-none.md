---
'@office-kit/xlsx': patch
---

fix: `loadWorkbook` rejected borders with `style="none"` (#99)

Files written by OnlyOffice (and any producer that emits the explicit no-border
value) set border sides to `style="none"`, which is the first value of
ECMA-376's `ST_BorderStyle` enumeration. Loading such a file failed with
`expected one of [thin, medium, ...]; got "none"`. `none` is now accepted and
round-trips faithfully; it draws no stroke in HTML/SVG preview.
