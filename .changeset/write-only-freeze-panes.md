---
'@office-kit/xlsx': minor
---

feat(streaming): a write-only worksheet can freeze panes ([#320](https://github.com/office-kit/xlsx/issues/320)). `ws.setFreezePanes(topLeft)` takes the same `"A2"` ref or `{ rows, cols }` counts as the modelled `setFreezePanes` and rejects the same inputs. Like `setColumnWidth`, call it before the first `appendRow`; a later call throws `OpenXmlIoError`. A sheet that never calls it is written as before.
