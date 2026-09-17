---
"@office-kit/xlsx": minor
---

Bound write-only shared-string retention for exports with many distinct plain
or rich-text values. The writer now retains at most 100,000 shared strings
within an 8 MiB string/XML payload budget, writes new values inline after the
budget is exhausted, and streams the final shared-string table in chunks.

Compatibility note: public APIs and cell values are unchanged, but write-only
ZIP bytes and shared-string IDs may change, and consumers inspecting raw XML
must support both shared and inline strings. Normal model-based saves are
unaffected. Streaming reads now decode OOXML text escapes in inline strings,
including rich-text runs.
