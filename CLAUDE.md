# CLAUDE.md

Project-specific instructions for Claude (and humans) working in this repository.
This file is the load-bearing source of truth for "how do we write code here." If
something in here conflicts with what you would do by default, this file wins.

## Project overview

**@office-kit/xlsx** is a TypeScript library for reading and writing Excel `.xlsx`
workbooks from Node 22+ and modern browsers, with no runtime dependencies on
Python or Excel. Inspired by [openpyxl](https://openpyxl.readthedocs.io/).

- **Status**: pre-1.0 alpha. Core read / write / streaming pipeline is in place
  and round-trips real-world fixtures (including pivot tables and `.xlsm`),
  but APIs may shift before `1.0`.
- **License**: MIT, single tier, no Pro / paid features.
- **Audience**: OSS — code that external contributors and end users will read.

See `README.md` for the user-facing pitch and `CONTRIBUTING.md` for the
day-to-day contributor workflow.

## Tech stack

- TypeScript with strict mode + `exactOptionalPropertyTypes` +
  `noUncheckedIndexedAccess`. No `as any`, no `@ts-ignore`.
- Runtime: Node `>=22` (relies on built-in `Web Streams`, `Blob`, `fetch`).
  Modern browsers via the same APIs.
- Runtime dependencies: `fflate` (deflate / inflate), `saxes` (SAX XML),
  `fast-xml-parser`. Anything else is a build-time dependency.
- Tooling: pnpm, vitest, oxlint, knip, tsdown (bundler), size-limit, typedoc,
  changesets (release management), libxml2-utils (`xmllint`) for ECMA-376 XSD
  validation in CI.

## Repository layout

```
src/
  io/         loadWorkbook / saveWorkbook / workbookToBytes + Source/Sink
  node/       Node fs glue (fromFile / toFile / fromBuffer / ...)
  streaming/  read-only iter + write-only append
  workbook/   createWorkbook, addWorksheet, defined names
  worksheet/  setCell, getCell, mergeCells, tables, ...
  cell/       cell value-model + inline rich text
  styles/     fonts, fills, borders, alignment, number formats
  chart/      c: (legacy) + cx: (modern) chart kinds
  chartsheet/ standalone chartsheets
  drawing/    anchors, images, chart placement
  packaging/  OPC (Open Packaging Conventions)
  zip/        zip read/write, decompression-bomb defense
  xml/        XML reader/writer abstractions
  schema/     OOXML schema + types
  utils/      error hierarchy, shared helpers

tests/
  Mirrors src/ layout. Includes roundtrip tests against reference/openpyxl/
  fixtures (git submodule) and ECMA-376 conformance tests under conformance/.

reference/openpyxl/   git submodule — fixture corpus, do not edit
docs/                 documentation site source
.github/              CI workflows, issue/PR templates, template-compliance
.claude/skills/       workflow-specific guides Claude consults
```

## Operating context

- **Readers**: yourself in 6 months, a first-time contributor, an AI agent.
- OSS is not "code only I touch." Optimize for **a stranger not getting
  confused**, not for your own convenience.

## Core principles

Every line of code must be justified.

| Principle               | Meaning                                                                              |
| ----------------------- | ------------------------------------------------------------------------------------ |
| Simplicity              | No premature abstraction, no features for hypothetical futures. YAGNI.               |
| Consistency             | Match existing patterns. New patterns require an explicit reason.                    |
| Performance             | Don't write N+1 / O(n²) in the first place. Defend with code shape, not profilers.   |
| Security                | Validate at boundaries (external input, file I/O, external APIs). Watch OWASP.       |
| Maintainability         | Write code that you in 6 months and a new contributor can read and understand.       |
| Backwards compatibility | Public API is preserved unless a breaking change is genuinely justified. Pre-1.0 we still try to call out breaks clearly. |

## "One way to do one thing"

> A capability has exactly one canonical path through the public API.

Do not add a **parallel API** — a second path that produces the same result as an
existing one — just because it's shorter, more discoverable, or "feels nicer." Reasons:

1. Every reader pays the "which one should I use?" cost on every code review and onboarding.
2. Two APIs = two of everything: docs, tests, types, bundle size, compatibility surface.
3. If the README says "use X for Y" but the library also accepts Z, the docs are lying.

**The rule bends** in two cases:

- An existing path is **misleadingly named** — rename it, don't parallelize.
  (Pre-1.0: rename outright. Post-1.0: deprecate + remove on next major.)
- A capability is reachable but only at an abstraction so low that every real user
  re-implements the same wrapper — graduate the wrapper into the library and
  hide the low-level path behind an "escape hatch" subpath. @office-kit/xlsx's
  `@office-kit/xlsx/xml`, `/zip`, `/packaging`, `/schema` subpaths exist as those
  escape hatches; new high-level features should extend `io`, `streaming`,
  `workbook`, `worksheet`, `cell`, `styles`, `chart`, `drawing` instead.

Decision flow when evaluating a feature request:

1. Is the capability **already reachable** through the public API? → Yes: reject.
2. Is the unreachable capability **inside this project's scope**? → No: point to
   the appropriate other library (SheetJS for `.xls`/`.xlsb`/`.ods`/`.csv`,
   `read-excel-file` / `write-excel-file` for simple browser cases, etc.).
3. If it **replaces** an existing path, is the replacement **strictly better** on
   every dimension that matters (ergonomics, types, performance, size)? → If only
   some dimensions are better, you're proposing a parallel API → reject.

## Subpath imports — no root barrel

@office-kit/xlsx has **no root barrel export**. Every export lives behind a section
subpath (`@office-kit/xlsx/io`, `@office-kit/xlsx/streaming`, …). Each export has exactly one
home — no convenience re-exports.

Consequences when adding code:

- New public APIs live in the subpath that owns the area. If you're tempted to
  re-export from a second subpath "for ergonomics," stop — that's a parallel
  path.
- `@office-kit/xlsx/io` and `@office-kit/xlsx/streaming` have hard bundle budgets enforced by
  `size-limit`. New dependencies in those graphs need a budget conversation.
- `"sideEffects": false`. Anything you add must be tree-shakable.

## Defensive programming

**Yes at boundaries; no internally.**

| Situation                                    | Defend? | Example                                                |
| -------------------------------------------- | ------- | ------------------------------------------------------ |
| External input (xlsx bytes, user-supplied values) | **Yes** | Validate, fail loudly via `OpenXmlError` subclass      |
| External API / fs / stream I/O               | **Yes** | Error handling, timeouts, decompression-bomb limits    |
| Untrusted archives                           | **Yes** | `decompressionLimits` is on by default — keep it on    |
| Already validated upstream                   | No      | Don't re-validate — that's noise                       |
| Already guaranteed by the type system        | No      | Don't write redundant null checks / optional chaining  |
| Cases that are impossible by type definition | No      | No "just in case" guards                               |

"Just in case" checks pile up and **bury the real validation** under noise. If a
value has already been validated upstream or is guaranteed by the type system,
do not re-check it.

**Don't swallow exceptions.** Catching and silently returning `null` for
unexpected errors is anti-defensive — it hides bugs and security issues.

- **Internal logic threw** → that's a bug. Let it propagate so it surfaces fast.
- **External input was invalid** → throw an `OpenXmlError` subclass with a
  `cause` chain. Don't pretend the input was valid.

## Hard "no"s

- **N+1 access**: sequential `await` inside a loop (fs / network / streaming
  source). Bulk it.
- **O(n²) operations**: `find` / `filter` / linear search inside a loop over
  the same data. Build a `Map` / `Set` first.
- **Type-cast escape hatches**: `as any`, `as unknown as T` outside the
  irreducible WHATWG-stream / Buffer↔Uint8Array boundary in `src/io/`. Use
  validation to convert safely.
- **Redundant "just-in-case" checks**: re-validating values whose contract is
  already met.
- **Magic numbers / strings**: name them as constants. The OOXML constants live
  in `src/schema/`.
- **Dead code**: "might use this later," commented-out implementations. Delete.
- **Swallowed exceptions**: `catch` blocks that silently return null. Let errors
  propagate.
- **Plain `Error`**: throw an `OpenXmlError` subclass (`OpenXmlIoError`,
  `OpenXmlSchemaError`, `OpenXmlDecompressionBombError`, …). Chain causes via
  `{ cause }`.
- **Silent breaking changes to public API**: even pre-1.0, breaking changes
  belong in a changeset.

## Comments

Comments explain **why**, not what.

```ts
// Self-evident "what" — drop it.
/** Get a user. */
const getUser = (id: string) => { /* ... */ };

// "Why" — the code can't say this on its own.
// fflate's inflateSync has no abort hook, so we use the streaming Inflate
// class even on sync reads — that lets us bail out mid-flight on CD-lying
// zip bombs.
const inflater = new Inflate();

// About the current task — that belongs in the PR description.
// Added for issue #123
// TODO: also handle X later

// Constraints / pitfalls a future reader needs to know.
// ECMA-376 §18.2.20 allows empty sheet names; treat them as a valid
// (but unindexed) sheet rather than rejecting.
```

**Don't write a comment when**:

- The identifier name already says it
- It references the current task / PR / issue (Git history covers that, and these rot)
- It marks deleted code (`// removed: ...`)

**Do write a comment when**:

- The implementation looks weird and the reason is non-obvious (past bug,
  performance workaround, third-party API / spec quirk)
- There's an implicit constraint or invariant
- A future reader is likely to step on a specific rake

## Use the type system

| Pattern                      | What it buys                                                              |
| ---------------------------- | ------------------------------------------------------------------------- |
| Discriminated unions         | Cell values are `{ type: 'number'; value: number } \| { type: 'string'; … }`, never `any` |
| Exhaustiveness checks        | `switch` / `match` with `never` so adding a new variant fails to compile  |
| Parse, don't validate        | Convert input into a validated type once at the boundary; trust internally |
| `as unknown as T` only at WHATWG↔Node Buffer boundaries | The existing usages in `src/io/` are the template; new uses need justification |

## OSS-specific discipline

### Public API is a contract

Anything exported under a documented subpath is a contract. Users depend on the
**shape**, **name**, **behavior**, and **exceptions** it produces.

- **Adding** a new export: minor bump, fine.
- **Changing existing behavior**: breaking change → major bump (post-1.0) or
  clearly flagged minor (pre-1.0).
- **Removing or renaming**: breaking change.

Anything intentionally not part of the public API must be **explicitly internal**
(not re-exported from a subpath, or under an internal subpath). If users can
import it, they will, and you'll be on the hook for it.

### CHANGELOG is for users, not for you

Write from the user's perspective. @office-kit/xlsx uses
[Changesets](https://github.com/changesets/changesets) — run `pnpm changeset`
to add an entry.

- ❌ `refactor: extract helper function` — users don't care.
- ✅ `feat: add async loading option to loadWorkbook` — users want to know.
- ✅ `fix: loadWorkbook crashed on files with empty sheet names (#123)` — symptom-based.

A pure-internal-refactor PR doesn't need a changeset; if it does need one, mark
it `chore` so it doesn't show up as user-facing.

### Issues and PRs are a conversation, not a queue

- When asked a question, **point at existing docs** before writing code.
- Evaluate feature requests against the "one way to do one thing" rule (is it
  already reachable?).
- For bug reports, **write a failing test first**. Don't fix what you can't
  reproduce. xlsx-related bugs almost always need the offending `.xlsx`
  attached.

See `.claude/skills/issue-triage/SKILL.md` for the full workflow.

## Tests

- **Unit tests** live under `tests/<area>/`, mirroring `src/`.
- **Roundtrip tests** load a fixture, save it, and assert the output equals the
  input semantically. They're how we keep parity with Excel / LibreOffice /
  openpyxl output. Many use fixtures from the `reference/openpyxl/` submodule
  (run `git submodule update --init --recursive` if they're missing).
- **Property-based tests** use `fast-check` for invariants that should hold
  over arbitrary inputs.
- **ECMA-376 conformance** lives under `tests/conformance/`. New worksheet
  XML output should pass `validate.ts`'s OPC + XSD + semantic checks. CI
  installs `libxml2-utils` so XSD validation never silently skips.
- **Performance gates** under `tests/perf/` are off by default — set
  `PERF_GATE=1` to enforce. CI runs them on a dedicated job.

If a new test depends on a CLI tool not already in CI, install it explicitly
in `.github/workflows/ci.yml` and document the dependency in the test file.

## Security

This library reads adversarial input (xlsx files from untrusted sources).
Defense-in-depth is enforced:

- `decompressionLimits` is on by default in `loadWorkbook` / `loadWorkbookStream`
  (per-entry size cap, total archive cap, compression ratio cap). Keep it on.
- XML entity expansion / DOCTYPE is rejected by the parser.
- Path traversal in zip entries is rejected at the packaging layer.
- All error paths throw an `OpenXmlError` subclass — never expose raw
  third-party errors to consumers.

See `SECURITY.md` for the reporting process and full scope. Report
vulnerabilities **privately** via GitHub Security Advisories — never in a
public issue.

## The AI-slop era

**LLM-authored issues, PRs, and review comments are now common.** They tend to
be formally well-structured but substantively thin: not reproducible, already
addressed, out of scope, or copy-pasted from documentation.

This repository takes a hard line:

- **Templates are mandatory.** Issues and PRs that don't follow the templates
  are auto-closed by `.github/workflows/template-compliance.yml`.
- **Repeated low-effort AI submissions can lead to a ban.** This is stated in
  the templates so contributors (and the LLMs they use) know upfront.
- **Using AI is fine. Posting AI output without reading it is not.** Whatever a
  model produces, **a human is responsible** for whether it's worth a
  maintainer's time.

When Claude or any other LLM works in this repository, it must **re-read its
own output and ask: is this thin, generic, or templated?** before posting
anything visible to the public.

## Skills

The `.claude/skills/` directory contains workflow-specific guides. Use them
when the situation matches.

| Skill                | When to use                                                              |
| -------------------- | ------------------------------------------------------------------------ |
| `pr-workflow`        | Creating a PR                                                            |
| `full-code-review`   | Reviewing a branch from a maintainer's perspective before opening a PR   |
| `review-response`    | Responding to GitHub review comments                                     |
| `run-check-and-test` | Running quality checks and tests before commit / PR                      |
| `issue-triage`       | Classifying a GitHub issue and routing it to the right workflow          |

When you add a new skill, append it to this table.

## Quality gate

Before opening a PR, run the full gate locally — CI mirrors it.

```sh
pnpm typecheck   # tsc --noEmit under strict settings
pnpm lint        # oxlint (fix with pnpm lint:fix)
pnpm knip        # unused exports / files / deps
pnpm test        # unit + property + roundtrip (~2000 cases, <30s)
pnpm build       # produce dist/
pnpm size        # size-limit budgets
```

`pnpm prepublishOnly` chains the gate the way CI does. See `CONTRIBUTING.md` for
the full script reference.
