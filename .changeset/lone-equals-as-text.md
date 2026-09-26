---
"@office-kit/xlsx": patch
---

`bindValue(cell, '=')` stores the text `=` instead of throwing, and `inferCellType('=')` returns `'s'`. Excel keeps a lone `=` typed into a cell as text; every other string that starts with `=` still takes the formula path, and the ones Excel refuses (`'= '`, `'==A1'`) still throw.

Correction to the 0.22.0 notes, which said reading was unaffected: a `<formula>`, `<formula1>`, `<formula2>` or `<definedName>` stored as `==$A$1>0` loaded as `=$A$1>0` before 0.22.0 and loads as `$A$1>0` since, and an `<f>` with three or more leading `=` no longer keeps one. The rejection added in 0.22.0 names the constructor that rejected the text, so `setFormula` and `bindValue` report `makeFormula`.
