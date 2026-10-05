---
'@office-kit/xlsx': patch
---

fix: drop the `fast-xml-parser` runtime dependency, shrinking `@office-kit/xlsx/io`

- `parseXml` now uses a small built-in lexer instead of `fast-xml-parser`, which removes about 74 kB minified (about 20 kB brotli) from bundles that load workbooks. Parse results are unchanged: the new lexer is checked against `fast-xml-parser` on every XML part of the test fixtures.
- A payload whose only top-level content is text followed by a stray closing tag (for example `x</q>`) used to throw a raw `RangeError` (stack overflow). It now throws `OpenXmlSchemaError` like any other unreadable payload.
