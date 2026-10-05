---
'@office-kit/xlsx': minor
---

feat: `getCellDisplayText` renders Japanese era dates and Japanese-locale names the way Excel does: `[$-411]ggge"年"m"月"d"日"` shows `令和5年3月15日`, and `[$-411]` sections print `aaa` / `ddd` as `水`, `mmm` as `3月` and `AM/PM` as `午前` / `午後`. `isDateFormat` now treats `g` and `aaa` codes as dates.
