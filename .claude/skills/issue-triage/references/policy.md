# @office-kit/xlsx policy: one way to do one thing

Read this when evaluating any **feature request** against @office-kit/xlsx's
design philosophy. The policy decides whether a request is accepted or
rejected — not your aesthetic preference, not whether the proposed API
looks nice in isolation.

## The rule

> **There is exactly one way to do one thing.**

If a capability is already reachable through @office-kit/xlsx's existing API, we
do not add a second path to it — even if the second path is more
convenient, more discoverable, or shorter to type.

This is a deliberate trade-off, not a stylistic accident.

## Why

1. **Cognitive load on the reader.** Every parallel API forces the next
   developer to ask "which one should I use, and why are there two?"
   That question has a real cost across thousands of code reviews and
   onboarding sessions. One way means no question.

2. **Cognitive load on the writer.** Auto-complete that lists three
   functions doing similar things is worse than auto-complete that lists
   one. @office-kit/xlsx ships subpath imports specifically so each subpath's
   surface stays minimal.

3. **Maintenance multiplier.** Every public path is a forever
   commitment: tests, docs, type-stability under
   `exactOptionalPropertyTypes`, schema-conformance proof, bundle-size
   budget. Two paths = two of everything.

4. **Documentation truth.** If the README says "use X to do Y", and the
   library also accepts a second way, the README is silently lying.
   Either we keep the docs honest by removing the alternate path, or we
   double the doc surface. The former is cheaper.

5. **Bundle size.** @office-kit/xlsx holds tight bundle budgets (≤120 KB
   min+brotli for the full lib, ≤80 KB for streaming). A "convenience
   helper" that re-implements an existing primitive is bytes shipped to
   every user for the benefit of one.

## What counts as "the same thing"?

Two paths are the same thing if a reasonable user would, in the same
situation, pick either. Examples of paths that are the same thing:

- `getCellValue(ws, r, c)` and `getCell(ws, r, c).value`. Both return
  the value of a cell in a worksheet. Reject.
- `setRowHeight(ws, n, h)` as a convenience over a row-dim mutator that
  already exists. Reject.
- `loadWorkbookFromFile(path)` as a sugar over
  `loadWorkbook(fromFile(path))`. Reject — `@office-kit/xlsx/node` already
  composes this.
- A "fluent" chained builder mirroring an existing imperative API.
  Reject.

Two paths are **not** the same thing if they enable different shapes of
work:

- `loadWorkbook` (eager, full DOM) vs. `loadWorkbookStream` (iterator,
  fixed memory). Different memory profiles, different access patterns.
  Both kept.
- `setCell` (worksheet mutation) vs. `appendRow` (write-only streaming).
  Different write models. Both kept.

The deciding question: **does the new path do something the existing
path cannot?** If the answer is "no, just nicer", reject.

## Checklist for evaluating a feature request

Run through these in order. Stop at the first NO and reject.

1. **Does the request enable a capability that is currently
   unreachable?** If the user can already do it with the public API,
   reject.

2. **Is the unreachable capability inside @office-kit/xlsx's scope?** Scope is
   reading and writing OOXML `.xlsx` files (and `.xlsm` passthrough).
   Out of scope: `.xls` (BIFF), `.xlsb`, `.ods`, `.csv`, generic
   spreadsheet UI rendering. Out-of-scope requests are rejected with a
   pointer to the appropriate library (SheetJS for legacy formats, etc.)
   — see the README's "When NOT to use @office-kit/xlsx" section.

3. **If the request replaces an existing path, is the replacement
   strictly better?** "Strictly better" means: every existing caller
   would prefer the new path on every dimension that matters (ergonomics,
   types, performance, bundle size). If the new path is only better in
   *some* dimensions, you're proposing a parallel API — reject. If it's
   strictly better, the existing path must be removed in the same
   change (with a deprecation window if pre-1.0 churn matters).

4. **Does the new path round-trip?** @office-kit/xlsx's CI gate includes a
   3-tier validator (OPC structure + ECMA-376 XSD + semantic
   invariants). A new write API that produces files that don't parse
   back, or that modifies bytes that should be passthrough, is not
   shippable.

5. **Does it fit inside the bundle budget?** `@office-kit/xlsx/io` ≤ 120 KB,
   `@office-kit/xlsx/streaming` ≤ 80 KB. A feature that pushes a subpath over
   budget needs a story for what comes out — and "nothing comes out" is
   not a story.

If you reach the end with all YES answers, the request is **acceptable
in principle**. That doesn't mean implement immediately — write the spec
comment first (see `SKILL.md` §3b) and let the maintainer (the user)
sign off.

## How to phrase a rejection

Rejections are common and they should not feel hostile. Tone:
appreciative of the report, concrete about the existing path, honest
about the policy.

Template (translate to the issue's language):

> Thanks for the suggestion!
>
> @office-kit/xlsx already covers this via `<existing API>`:
>
> ```ts
> <minimal example>
> ```
>
> The library follows a "one way to do one thing" policy — adding a
> second path to the same capability would mean two APIs to learn,
> document, and maintain for the same outcome. Keeping the surface
> small is a deliberate design choice (see the README §"Subpath
> entries" for context).
>
> If `<existing API>` doesn't actually cover your use case, I missed
> something — could you share a code snippet of what you tried and what
> went wrong? @<op>

Don't apply `wontfix` yourself — leave that to the maintainer. The
comment is enough; the maintainer will close after reading it.

## When the policy bends

The policy is strict but not absolute. It bends when:

- An existing path is **misleading by name** and a new name is the
  fix — but in that case, the old name is renamed (or deprecated +
  removed pre-1.0), not parallelized.
- A capability is reachable but only at a level of abstraction so low
  that every real user would re-implement the same wrapper. In that
  case the wrapper graduates into the library — but the low-level path
  is hidden behind a subpath like `@office-kit/xlsx/xml` so it's clearly the
  escape hatch, not the primary API.

In both cases, the result is still **one path per thing** for the
typical user. If you find yourself arguing for two paths "because
different users want different things", that's the policy violation the
rule exists to prevent — most of the time, the right answer is to pick
one and document why.
