# OOXML quality assurance

This directory is the starting point for a test262-style, specification-linked
suite. It tests a **bounded Transitional SpreadsheetML profile**, not every
feature of ECMA-376. Passing it is not an OOXML certification.

## Why these gates

Recent regressions around optional booleans, explicit `none` values, XML text,
and browser input paths require more than writer → reader round trips. A reader
and writer can share the same mistake. Likewise, an XSD-valid document can lose
formatting or contain a dangling relationship or invalid style index.

The reference standard is **Office Open XML (OOXML), ECMA-376**, rather than
OpenOffice's OpenDocument format. LibreOffice is an independent consumer and
compatibility target; its behavior does not override normative expectations.

| Contract | Independent evidence | Failure policy |
| --- | --- | --- |
| XML well-formedness | namespace-aware SAX tree, before MC preprocessing | reject malformed XML, including unknown parts |
| OPC package graph | separate content-type/relationship traversal | reject missing targets, duplicate IDs, wrong selected relationship types |
| Schema structure | vendored ECMA XSDs through `xmllint --nonet` | missing/crashed runner is inconclusive, never a pass |
| Cross-part semantics | actual child counts and coordinate/index checks | reject invalid references even without optional declared counts |
| Reader meaning and edit preservation | hand-authored inputs and independent expected values | compare meaning before and after an unrelated edit |
| Browser distribution | built package in Chromium, Firefox, WebKit | compare expected meaning through four adapters and two reading modes |
| Second schema implementation | pinned Open XML SDK | reject validation errors and require negative calibration to fail |
| Application compatibility | headless LibreOffice plus Python XML/ZIP inspection | record exact known differences; reject new or changed differences |

The oracle uses `fflate` rather than the production ZIP parser, `saxes` rather
than the production XML model parser, and no production manifest/relationship
parsing or metadata enums. Some dependencies remain shared; independence is
strengthened by libxml2, .NET Open XML SDK, and Python/LibreOffice. This is not
complete implementation independence.

## Corpus and calibration

`corpus/manifest.json` records the specification edition, provenance, license,
stable case ID, schema type/clause, input fragments, and expected semantics.
The 82 cases cover:

- Seven font boolean properties: absent, implicit true, `0`, `1`, `false`, `true`.
- Five underline and fourteen border values, including explicit `none`.
- Four cell boolean spellings, Unicode character references, CDATA, optional row index.
- Numeric, boolean, text, empty and error formula caches; sheet-qualified names
  in shared formulas, including the translated follower.
- Inline/shared rich text, quoted/scoped defined names and leap-day/date-epoch
  semantics. Dates use numeric serials plus number formats in this Transitional profile.

Streaming intentionally checks cached formula results: its current public contract
does not expose the ordinary reader’s formula objects. LibreOffice recalculates
formulas, so its independent expected values differ from an empty input cache.

Inputs are assembled independently of the library writer. Each case validates
the input, checks the loaded model, edits B2, validates and reloads the output,
and checks streamed rows. Additional tests cover 1/7/257-byte adapters and
equivalent namespace prefixes and single-quoted attributes. Absent font
properties are asserted explicitly; empty streams cannot pass vacuously.

`validator-calibration.test.ts` mutates one fault at a time, including missing
package relationships/content types, wrong relationship type or target type,
prefixed row/cell mismatches, false style counts, missing counts with invalid
indices, missing shared strings, excluded dimensions, and malformed unknown
XML. Positive controls cover namespace-equivalent XML and MC selection.
Skipping XSD checks or encountering unsupported XML returns `incomplete`;
an unavailable XSD executable returns `inconclusive`. Only `valid` sets `ok`.

MC preprocessing understands the explicitly listed namespace profile,
requires every `Choice/@Requires` namespace to be understood, preserves
namespace scope when promoting children, handles `ProcessContent`, and rejects
unsupported `MustUnderstand`. Raw XML is always checked first. It is not a
complete implementation of every MC rule; for example, full AlternateContent
ordering and preservation directives need dedicated expansion.

## Run locally

Use the repository's pinned pnpm and a supported Node version. Install
`xmllint` (libxml2-utils on Ubuntu), then:

```sh
pnpm install --frozen-lockfile
pnpm exec vitest run tests/conformance
pnpm exec playwright install chromium firefox webkit
pnpm test:browser
```

For independent validators, install .NET SDK 10, LibreOffice Calc, and Python 3:

```sh
pnpm qa:corpus
dotnet restore tests/conformance/sdk --locked-mode
dotnet run --project tests/conformance/sdk --no-restore -- .qa/corpus
pnpm qa:office
```

`SOFFICE` can select an executable. The office runner uses its own profile and
clears its own output before conversion so stale files cannot satisfy a check.
The SDK requires exactly the output set in the manifest. Both refuse an empty
corpus. Generated inputs/outputs, diagnostics, and compatibility results are
under `.qa/`; these are build artifacts, not golden files from the writer.

For reproducible generation:

```sh
QA_FUZZ_RUNS=1000 QA_FUZZ_SEED=376262 pnpm exec vitest run tests/conformance/property.test.ts
```

The generator explores bounded cells, styles, merges, names, dimensions, and
panes. It does not generate every XML grammar or hostile ZIP input. Failures
retain fast-check seed, shrink path, counterexample, and validator diagnostics.
Convert a minimized failure into a permanent corpus or calibration case.

## CI and interpretation

Existing Node/OS, package, size, and performance gates remain. PR CI adds the
three-browser corpus and the SDK/LibreOffice validators. A nightly job runs
1,000 generated workbooks with a recorded seed. Diagnostic artifacts are
uploaded even on failure. These workflow changes have been checked locally;
GitHub-hosted runs occur after the branch is submitted.

Open XML SDK checks schema validity using the Office2016 profile. It does not
prove Excel rendering, formula calculation, or the absence of Excel repair
dialogs. Browser tests check model and save/reload semantics; they do not
independently run libxml2 on every browser-produced byte stream.

The observed LibreOffice 26.2.4.2 baseline has **74 matching cases and 8 known
differences**: enabled `condense`/`extend` are dropped (six lexical variants),
and accounting underlines become ordinary single/double underlines (two).
`corpus/libreoffice-profile.json` records exact expected/actual pairs and their
reasons. The report marks these as `known-compatibility-difference`, not pass.
The gate permits this explicit baseline but rejects any new difference and
also rejects disappearance of an old difference until the baseline is reviewed.
Different LibreOffice versions may therefore require a deliberate review.

## Expanding toward test262

Add cases by normative requirement, not by production method count. Each new
feature needs an independently constructed valid input and expected model or
preservation contract; add a single-fault negative control when the validator
is meant to reject it. Keep equivalent encodings together (prefixes, quoting,
optional/default values, element presence), and exercise input adapters and
streaming where applicable. Do not turn a regenerated writer output into its
own expected answer, or broadly allowlist validation failures.

The next coverage priorities are Strict OOXML, unknown extension/VBA/custom
parts preservation, charts/tables/drawings and their relationship combinations,
formula/date/number-format semantics, Excel-produced fixture provenance, and
Excel open/save/repair checks. Duplicate ZIP entries, full OPC URI rules,
resource exhaustion, full MC compliance, visual rendering, calculation, and
exhaustive feature interactions are not guaranteed by the current suite.

Track coverage by specification requirement × read/write/preserve × adapter ×
runtime, with explicit supported/unsupported/untested outcomes. A growing case
count alone must never be presented as a percentage of OOXML conformance.

## Local verification, 2026-10-02

On macOS with Node 26, the complete unit suite passed (361 files, 3,366 tests),
as did 1,000 generated workbooks with seed 376262. All 204 browser cases passed
(68 per engine, four adapters × two reading modes each). Open XML SDK 3.5.1
accepted all 68 outputs and rejected its negative calibration. LibreOffice
26.2.4.2 produced the 60/8 compatibility baseline described above.

Frozen-lockfile installation, lint (no lint warnings), TypeScript, build, knip,
size limits, published type resolution, real consumer compilation/execution,
and the enabled performance gate passed. The lower-level package checks cover
node16, nodenext, and bundler resolution. Existing pnpm configuration and root
Svelte configuration notices are unrelated tool notices, not test failures.

This is local evidence. The new Linux CI jobs, all nine existing Node/OS matrix
combinations, and desktop Excel itself have not been executed in this run.

References: [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/),
[Open XML SDK validator](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.validation.openxmlvalidator),
[test262](https://github.com/tc39/test262),
[Playwright browsers](https://playwright.dev/docs/browsers).

## Package preservation contracts

`preservation.test.ts` reads pinned openpyxl fixtures with `fflate`/`saxes` and
compares the serialized bytes/content types of VBA, controls, printer settings
and external-link parts after an unrelated T100 edit. It separately compares
relationship type, source, mode and resolved target, plus XML bindings, so a
new relationship ID is permitted but a lost connection is not. The legacy
drawing fixture's contract covers the form-control sheet; the modeled comment
sheet regenerates its VML with a new path and is outside byte preservation.
Hand-authored extensions cover unknown worksheet/style subtrees and empty
style pools. Single-fault controls change bytes, a content type, a target and
an extension subtree; each must be detected. Existing reference fixture
provenance/license remain those of the pinned openpyxl submodule.

These contracts exposed and fix two save errors: the VBA relationship must use
Microsoft's macro relationship namespace, and VBA must be typed by a part
override rather than replacing every unrelated `.bin` default.
