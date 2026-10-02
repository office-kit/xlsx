# OOXML quality assurance

This directory is the starting point for a test262-style, specification-linked
suite. It tests **bounded Transitional and Strict SpreadsheetML profiles**, not every
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
| Browser distribution | built package in Chromium, Firefox, WebKit | compare expected meaning through four adapters/two reading modes; validate actual saves with XSD and SDK |
| Second schema implementation | pinned Open XML SDK | reject validation errors and require negative calibration to fail |
| Application compatibility | headless LibreOffice plus Python XML/ZIP inspection | record exact known differences; reject new or changed differences |
| Desktop Excel | separately recorded native Mac Open/Save As | require no repair, provenance, exact cases/hashes and independent semantic comparison; manual evidence only |
| Coverage inventory | requirement × operation × adapter/runtime profile | fail missing cases/evidence or report drift; expose untested/unsupported guarantees |

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
complete implementation of every MC rule. Full Choice/Fallback ordering is
calibrated, including after a supported Choice; preservation directives and
other rules remain outside the bounded oracle profile.

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

The valid generator explores bounded cells, styles, merges, names, dimensions,
and panes. A separate invalid generator and single-fault calibration cover grid,
URI, ZIP and MC faults. Neither explores every XML grammar or hostile ZIP input. Failures
retain fast-check seed, shrink path, counterexample, and validator diagnostics.
Convert a minimized failure into a permanent corpus or calibration case.

## CI and interpretation

Existing Node/OS, package, size, and performance gates remain. PR CI adds the
three-browser corpus, SDK/LibreOffice validators, curated regression mutations,
manual Excel protocol calibration, and the coverage inventory. A nightly job
runs the valid and invalid generators with recorded seeds and 1,000 runs.
Diagnostic artifacts are uploaded even on failure.

Open XML SDK checks schema validity using the Office2016 profile. It does not
prove Excel rendering, formula calculation, or the absence of Excel repair
dialogs. Browser tests check save/reload semantics and independently validate
every saved byte stream with ZIP/XML/OPC/XSD and Open XML SDK. Streaming read
paths check row values/caches and date epoch flags, not saved output or all
ordinary-reader metadata.

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

The current additions establish bounded formula/cache, rich-text, name and date
contracts, selected serialized-part preservation, ZIP/URI/resource and MC
calibration, independent browser-save validation, curated mutation detection,
and a nine-case native Mac Excel profile. Further work remains for independent
full Strict conformance, arbitrary custom-part and feature interactions, full OPC URI
and MC compliance, hostile ZIP grammars beyond the bounded profile, formula
calculation, visual rendering, and Windows desktop Excel execution. Existing
production tests outside this inventory do not imply these broader guarantees.

Track coverage by specification requirement × read/write/preserve × adapter ×
runtime, with explicit supported/unsupported/untested outcomes. A growing case
count alone must never be presented as a percentage of OOXML conformance.

## Verification checkpoint, 2026-10-02

On macOS with Node 26, the complete unit suite passed (3,401 tests), as did
1,000 valid generated workbooks with seed 376262. All 246 browser tests passed
(82 per engine, four adapters × two reading modes); all 984 saved files passed
independent XSD and Open XML SDK checks. Open XML SDK accepted all 82 Node
outputs and rejected its negative calibration. LibreOffice 26.2.4.2 produced
the 74/8 compatibility baseline. Five curated regression mutants were killed
by assertions, not compiler failures. Nine actual desktop Mac Excel cases
opened without repair and passed independent saved-file comparison.

Frozen-lockfile installation, lint, TypeScript, build, knip, size limits,
published type resolution, real consumer compilation/execution, and performance
checks passed. Submitted QA PRs require the GitHub Node/OS, browser, independent
conformance and packaging checks before merge. Desktop Excel is manual and
its nine observations cannot substitute for Windows or continuous desktop CI.

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

### Invalid-input calibration

`invalid-generation.test.ts` starts from hand-authored packages, checks the valid
XFD1048576 boundary, and generates out-of-grid references with a reproducible
fast-check seed and shrinking. Single faults cover invalid internal targets,
duplicate ZIP entries, inflation budgets, and the complete MC Choice/Fallback
sequence even after an earlier supported Choice. Invalid xml:space is rejected
before normalization. Nightly runs exercise both valid and invalid generators.

The independent ZIP preflight is a bounded UTF-8 ZIP32 profile: at most 10,000
entries, 16 MiB per inflated part and 64 MiB total. These are oracle resource
limits, not production API limits. ZIP64 is reported incomplete; it is not
misrepresented as fully validated. This suite verifies the oracle's rejection
behavior, not that every production reader rejects every malformed package.

### Curated regression mutations

`pnpm qa:mutations` copies sources/tests into a temporary directory, verifies a
green baseline, then reintroduces five documented regressions one at a time.
Every regression must fail a test assertion. A surviving mutant, changed source
anchor, empty run, timeout, compile failure or missing report fails the gate.
The working sources are never mutated. CI retains per-mutant diagnostics and
the exact failing test names in `.qa/mutations/`. This proves detection of this
curated set; it is not a general mutation score or a conformance percentage.
The ten selected regressions include local-name ownership, active-tab identity,
rich-text numeric coercion and public cell/row budgets. Each is first checked
against a passing baseline; runner failures cannot count as detection.

### Independent browser output validation

Each real Chromium/Firefox/WebKit test returns the actual saved bytes to the
Node host. Every case/adapter output must pass independent ZIP/XML/OPC/XSD and
semantic validation before the test passes. All 82 × 3 × 4 outputs are retained
and independently checked by Open XML SDK (Office2016 profile). The expected
output manifest is declared before execution, the export directory is cleared
first, and the SDK requires the exact output set. Missing cases, stale outputs
and a missing saved byte array cannot produce a vacuous pass. Streaming read
paths have semantic tests; they do not claim to produce a saved workbook.

### Native desktop Excel manual profile

No licensed Windows self-hosted runner is currently available. Desktop Excel
is a separately recorded manual profile, not an automatically passing CI job.
Prepare the current corpus, then use a fresh directory:

```sh
pnpm qa:corpus
pnpm qa:excel --prepare .qa/excel-session-2026-10-02
```

Open each of the nine `input/*.output.xlsx` files in licensed desktop Excel.
A repair dialog is a failure: decline repair and investigate. Save As the same
basename in `output/`, close all tested workbooks (including their lock files),
and record Excel's complete version/build, platform, observation timestamp,
`openedWithoutRepair: true`, and saved-file SHA256 in `observations.json`.
On macOS `shasum -a 256 FILE` computes the hash; on Windows use
`Get-FileHash -Algorithm SHA256 FILE`. Preparation already records input hashes.
Never set a no-repair observation for a workbook that was not actually opened.

```sh
pnpm qa:excel --check .qa/excel-session-2026-10-02
```

The checker requires the exact current normative case manifest, exact input and
saved output sets, all nine observations, and matching hashes. Failed reruns
remove stale passing reports. Python independently compares values, edit
preservation, name scope/reference, rich-text runs, date epoch and number format.
Keep the entire session directory with `results.json` and `check.log` as evidence.
A report alone cannot prove a human observation; this is a manual protocol.
`qa:excel-calibration` tests evidence rejection in Linux CI using synthetic
observations; it neither launches Excel nor counts as application compatibility.

`corpus/excel-mac-observation.json` records the actual 2026-10-02 native macOS
Excel 16.113.3 (16.113.26092714) run: nine opened without repair, saved, and passed
the independent comparison. The report covers shared formula caches, inline/SST
rich text, quoted/scoped names and both date systems. It is not a calculation
engine or visual-rendering certification, and is not evidence for Windows Excel.

The original scoped-name corpus used `$A$1` without a sheet reference. Both XSD
and Open XML SDK accepted it, but native Excel requested repair. The corpus now
uses `Audit!$A$1` with `localSheetId="0"`; this opened without repair. Name scope
and the target sheet reference are separate concepts. This corrects our test
input, not a demonstrated production serializer defect. Schema acceptance alone
must not be used as evidence that a workbook opens without Excel repair.

Date inputs also explicitly apply their number format (`applyNumberFormat="1"`).
The independent office comparison now checks this format as well as the serial
and epoch; a numerically unchanged date rendered as General is a failure.

Built-in 14 is locale-dependent. The bounded comparator accepts the equivalent
custom date spellings observed on macOS (`mm/dd/yyyy`) and Linux (`m/d/yyyy`),
as well as the listed two-digit-year forms; it still rejects General or an
explicitly disabled format. This is date-category preservation, not identical
locale-specific rendered text.

### Machine-readable requirement coverage

[`coverage.json`](coverage.json) inventories 22 bounded requirements and nine
profiles with subject (library/oracle/consumer), runtime, adapter, operation
(read/write/preserve/reject), status, exact corpus cases and evidence anchors.
[`coverage.md`](coverage.md) is generated from it. Unlisted combinations expand
to **untested**; **unsupported** means outside that named QA profile, not that a
library feature is necessarily unsupported. **Tested** means the stated bounded
assertions exist, not every input works. Consumer write describes import/export
of library output; oracle rejection does not assert production-reader rejection.

```sh
pnpm qa:matrix --write # after deliberate inventory changes
pnpm qa:matrix         # references, complete corpus assignment and report drift
pnpm qa:matrix-calibration
```

CI rejects missing/duplicate corpus assignments, missing test/file anchors,
unknown profiles/statuses, overlapping claims, incomplete tested case sets,
Windows marked tested without an available runner, and Mac claims outside the
recorded bounded profiles. Calibration supplies deliberately false inventories
and observations. The inventory checks metadata, not test execution or a human
observation; the referenced CI/manual gates supply that separate evidence.
Adding a corpus case requires an explicit requirement assignment. No count in
this inventory is a percentage of OOXML conformance.

## Editing sequences

`sheet-operations.test.ts` compares generated move/swap/rename/copy/remove
sequences against an independent tab/name model and inspects saved XML. It
checks local name ownership, active tabs, numeric cells, formula caches, and
unique table identifiers. Formula expressions remain verbatim: automatic
rewriting of renamed/deleted sheet references is outside this API contract.

PR CI retains the fixed 30-run, 10-operation regression seed. Nightly CI also
explores 250 sequences of up to 50 operations with its run number as the seed.
Failures are shrunk and retained as `.qa/sheet-operations-counterexample.json`
with the operation list, seed and shrink path; these are evidence for adding a
small permanent regression, not an automatically approved corpus case.
Replay the same generator and bounds with:

```sh
QA_EDIT_RUNS=250 QA_EDIT_SEED=123 QA_EDIT_MAX_OPERATIONS=50 QA_EDIT_PATH='0:1' pnpm exec vitest run tests/conformance/sheet-operations.test.ts
```

Use the recorded seed/path/maxOperations rather than the illustrative values.

## Public loader resource contract

`loader-boundaries.test.ts` exercises both public loaders through buffered and
1/7/257-byte stream sources. Independently assembled inputs establish exact
cell/row/decompressed entry/total boundaries, compression-ratio rejection and
typed failures for DTDs, mismatched tags and invalid character references.
Large declared dimensions do not consume the content budget; actual rows and
cells do. ZIP byte caps are enabled by default, while cell/row caps require
explicit caller configuration. A cap is not a deadline, a compressed-input
size limit or complete validation of all OOXML. Applications still need their
own upload-size/time boundaries. Existing lower-level tests cover dishonest
ZIP sizes, cancellation and cross-sheet/query budget behavior.
## Independent Strict validation

Pass `{ conformance: 'strict' }` to the test-only oracle to validate original
Strict namespaces against the unmodified ECMA-376 Part 1 fifth-edition schemas.
The default remains Transitional. Schema hashes and archive provenance are
pinned in `schemas/strict/provenance.json`; validation runs offline. Production
namespace normalization is not involved. Numeric/ISO-date/boolean inputs and
single-fault type/style/shared-string controls calibrate the profile.

Both genuine Excel Strict fixtures contain unqualified `dateCompatibility`,
which this normative XSD rejects. Tests retain those original bytes and record
the precise discrepancy; tolerant production reading is a different contract.
Schema validity alone does not establish formula calculation, visual fidelity
or support for every Strict feature.

## Mac Excel feature profile

A separate, independently assembled fixture combines a filtered table, cached
pie chart, PNG image and print settings. After an unrelated cell edit, CI checks
both packages against XSD/OPC rules and the saved package with Open XML SDK.
It also calibrates the independent Python comparison with deliberately detached
relationships, changed chart/table references, moved anchors, changed image
bytes and changed print settings. These synthetic controls do not execute Excel.

```sh
pnpm qa:excel-features
pnpm qa:excel --prepare /tmp/excel-features-new-session --features
# Open the one file in input/ in desktop Excel without accepting any repair.
# Save As the same basename into output/, then CLOSE the workbook.
# Record the actual version, no-repair observation, UTC time and saved SHA256
# in observations.json, as for the cell profile above.
pnpm qa:excel --check /tmp/excel-features-new-session --features
```

Use a fresh directory and retain the input file. Closing the workbook removes
Excel's temporary lock file; an unexpected XLSX in output/ is rejected.
The comparator follows attached relationships rather than accepting orphan
parts. It checks table range/header/filter, pie category/value references and
caches, exact image bytes, cell anchors and stored page setup/margins/centering.
Omitted fit-width/height values use the official CT_PageSetup default of 1;
changing a nondefault value still fails. The feature manifest and
`excel-features-mac-observation.json` pin this one-case profile and an actual
Mac Excel 16.113.3 Open / Save As observation with no repair and a passing
independent comparison. The evidence matrix keeps it separate from the
nine-case cell profile. This does not establish all table/chart/image features,
rendered geometry, actual pagination, or Windows Excel compatibility.

### Pinned native Excel regression

`pnpm qa:native-excel` replays the synthetic feature file saved by Microsoft
Excel 16.113.3 on macOS. `fixtures/excel-mac/manifest.json` records the original
observation, source/output hashes, license and published fixture hash. The public
fixture replaces personal `cp:lastModifiedBy` metadata with `QA` and repacks the
ZIP; all other uncompressed parts retain Excel's bytes. CI checks the hash,
OPC/XSD, independent attached-feature projection, public edit/reopen and SDK
validation. Excel's omitted fit dimensions are interpreted with their schema
default of one. This replay preserves a real application regression sample; it
does not execute Excel continuously or cover all Excel versions.

## Isolated adversarial resource checks

Build first, then run `pnpm qa:resources-calibration` and `pnpm qa:resources`.
The built public model and streaming loaders receive independently assembled
8 MiB inflated XML, 100,000 empty rows, 100,000 cells and DTD inputs, plus a
positive single-cell workbook with a misleading full-grid dimension. Both
loaders must return the configured typed error, or the exact positive value.
Each case runs in its own child process with a 128 MiB V8 heap cap, 15-second
wall timeout and a 256 MiB measured peak RSS assertion. Input generation and
module startup are included in measurements; reports record Node/platform.
Peak RSS uses Node's documented KiB `process.resourceUsage().maxRSS` value.

Timeout, heap exhaustion, crashes, excessive RSS and absent reports fail QA;
calibration deliberately triggers these failure paths. A heap cap is not an
OS limit on all native/buffer memory, and measured RSS is checked after exit.
This bounds the test runner and detects regressions on the stated inputs, not
a universal loader CPU deadline, upload-size cap or every hostile ZIP grammar.
CI executes the profile on Node 22 / Ubuntu 24.04 and retains partial diagnostics.
Reference: [Node resource usage](https://nodejs.org/api/process.html#processresourceusage).
## Multiple-sheet feature interactions

`pnpm qa:feature-interactions` combines tables, cached pie charts, PNGs, anchors,
merges, cellIs conditional formatting, whole-number data validation and print
settings on two independently assembled sheets. Different formulas, validation
bounds, table names and print orientations establish per-sheet ownership after
unrelated edits, move and swap. Each saved state passes OPC/XSD and CI SDK checks;
reopening also checks that edits remain on the correct sheet. Single-fault
controls must detect changed formulas/orientation and a drawing attached to the
wrong sheet. This is a bounded combination, not every rule/chart/table feature,
copy/rename formula rewriting, rendering or pagination.
