---
'@office-kit/xlsx': minor
---

Renamed the package from `xlsx-kit` to `@office-kit/xlsx`. The old `xlsx-kit`
package on npm is deprecated and will receive no further releases — install
`@office-kit/xlsx` and update every subpath import (`xlsx-kit/io` →
`@office-kit/xlsx/io`, and likewise for `/streaming`, `/workbook`, `/worksheet`,
`/cell`, `/styles`, `/chart`, `/drawing`, `/node`, and the low-level `/xml`,
`/zip`, `/packaging`, `/schema` escape hatches). No runtime behaviour changed.
