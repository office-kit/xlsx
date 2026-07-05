# Third Party Notices

`@office-kit/xlsx` is a TypeScript port of, and incorporates work derived from, the
following projects.

## openpyxl (MIT)

Source: <https://foss.heptapod.net/openpyxl/openpyxl> (canonical, Mercurial),
mirrored at <https://github.com/quintagroup/openpyxl> (Git, used as this
repository's `reference/openpyxl` submodule for porting reference only).

```
Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

## Runtime dependencies

The following npm packages ship as part of `@office-kit/xlsx` at runtime. Each is
licensed under its own terms — links go to the upstream repository where the
license text lives.

- **fflate** (MIT) — <https://github.com/101arrowz/fflate>
  ZIP / DEFLATE I/O. Used by `src/zip/reader.ts` (`unzipSync` fallback +
  `inflateSync`) and `src/zip/writer.ts` (`Zip` + `ZipDeflate` /
  `ZipPassThrough` streaming-deflate writer).
- **saxes** (ISC) — <https://github.com/lddubeau/saxes>
  Streaming XML parser. Drives the SAX iter API in `src/xml/iterparse.ts`
  which the streaming read-only path in `src/streaming/read-only.ts` uses.
- **fast-xml-parser** (MIT) —
  <https://github.com/NaturalIntelligence/fast-xml-parser>
  DOM-style XML parser used by `src/xml/parser.ts` for the eager
  load path; the streaming read-only path uses saxes instead.

## Dev dependencies

Dev-only dependencies declared in `package.json` (vitest, oxlint, tsdown,
typescript, fast-check, size-limit, …) are subject to their upstream
licenses; consult `pnpm-lock.yaml` for resolved versions.
