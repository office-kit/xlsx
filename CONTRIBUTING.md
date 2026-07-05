# Contributing to @office-kit/xlsx

Thanks for your interest in contributing. This document covers the day-to-day
workflow: how to set up a development environment, run the test suite, and
land a change.

## Reporting issues

- **Bug reports** — please use [`bug_report.yml`](.github/ISSUE_TEMPLATE/bug_report.yml).
  Include the @office-kit/xlsx version, a minimal reproduction, and the offending
  xlsx file (or at least a snippet that triggers it). Bugs in xlsx file
  output ("Excel says the file is corrupt") almost always need the file
  attached to be actionable.
- **Feature requests** — use [`feature_request.yml`](.github/ISSUE_TEMPLATE/feature_request.yml).
  Tell us the use case first; the API shape we land on usually depends on it.
- **Security issues** — please do **not** file a public issue. Follow the
  process in [`SECURITY.md`](SECURITY.md).

## Development environment

```sh
# Requires Node 22+ and pnpm.
git clone --recurse-submodules https://github.com/office-kit/xlsx
cd @office-kit/xlsx
pnpm install
```

The `--recurse-submodules` flag pulls the [`openpyxl`](reference/openpyxl)
reference checkout into `reference/`; many tests load fixture files from it.

### Useful scripts

| Command           | What it does                                                        |
| ----------------- | ------------------------------------------------------------------- |
| `pnpm typecheck`  | Run `tsc --noEmit` with the project's strict settings.              |
| `pnpm lint`       | Run `oxlint`. Fix with `pnpm lint:fix`.                             |
| `pnpm knip`       | Catch unused exports / files / dependencies.                        |
| `pnpm test`       | Vitest unit + property + roundtrip tests (~2000 cases, < 30s).      |
| `pnpm test:watch` | Same in watch mode.                                                 |
| `pnpm test:e2e`   | The slower end-to-end suite (writes real xlsx, requires more time). |
| `pnpm test:perf`  | Performance gates. Off by default — set `PERF_GATE=1` to enforce.   |
| `pnpm bench`      | Throughput micro-benchmarks (`tests/perf/throughput.bench.ts`).     |
| `pnpm build`      | Produce the publishable bundle in `dist/`.                          |
| `pnpm size`       | Run `size-limit` against `dist/` (gates against bundle bloat).      |
| `pnpm doc:api`    | Regenerate the TypeDoc JSON used by the docs site.                  |

The `prepublishOnly` script (`pnpm typecheck && pnpm lint && pnpm test &&
pnpm build && pnpm size`) is what CI mirrors. Run it locally before opening
a PR to catch most failures up front.

## Workflow

1. **Open or claim an issue.** For non-trivial work (new public API, breaking
   change, large refactor), please discuss the shape in an issue first so we
   can align before code is written.
2. **Branch from `main`.** Keep branches focused — one logical change per
   branch makes review and reverts straightforward.
3. **Write a test.** Bug fixes should add a regression test; features should
   ship with unit + integration coverage. The test layout under `tests/`
   mirrors `src/` — drop your test next to the matching module.
4. **Run the full gate locally.** `pnpm prepublishOnly` mirrors what CI runs.
5. **Add a changeset.** @office-kit/xlsx uses
   [Changesets](https://github.com/changesets/changesets) for release notes
   and version bumps. Run `pnpm changeset` and pick `patch` for fixes,
   `minor` for additive features, `major` only after coordination with the
   maintainer. Pre-1.0 we still try to call out breaking changes in `minor`
   bumps clearly.
6. **Open a PR.** Use the template; fill in what the change does and how
   you tested it. Link the issue you're closing.

## Code style

- **TypeScript strict mode is non-negotiable.** No `as any`, no
  `@ts-ignore`. Use `as unknown as T` only at irreducible JS↔WHATWG /
  Buffer↔Uint8Array boundaries — the existing usages in `src/io/` are the
  template.
- **Errors throw `OpenXmlError` subclasses.** Never throw bare strings or
  plain `Error`. Chain causes via `{ cause }`.
- **Pre-1.0 API is mutable.** Treat exports under `src/xml`, `src/zip`,
  `src/schema` as internal — they exist for advanced consumers but should
  rarely be the answer for new features. Prefer extending the public
  surfaces in `src/workbook`, `src/worksheet`, `src/io`, `src/streaming`.
- **No emoji in source files or commit messages.** The repository is
  plain-text only.
- **Comments explain *why*, not *what*.** Function names + types document
  *what*; comments earn their keep only when they capture invariants,
  trade-offs, or non-obvious history.

`oxlint` and `tsc` enforce the rest; please don't disable rules locally
without flagging it in the PR description.

## Tests

- **Unit tests** live next to the code under `tests/<area>/`.
- **Roundtrip tests** load a fixture, save it, and assert the output equals
  the input semantically. They're how we keep parity with Excel /
  LibreOffice / openpyxl output.
- **Property-based tests** use `fast-check` (`tests/perf/`,
  `tests/worksheet/`). Prefer them for invariants that should hold over
  arbitrary inputs.
- **ECMA-376 conformance** lives under `tests/conformance/`. New
  worksheet-level XML output should pass `validate.ts`'s OPC + XSD + semantic
  checks.

The CI matrix runs against Node 22 / 24 / 26 on Ubuntu. `xmllint` (from
`libxml2-utils`) is installed explicitly so schema validation never silently
skips. If your test depends on a CLI tool, document it in the test file and
add it to the install step in `.github/workflows/ci.yml`.

## Releases

@office-kit/xlsx publishes via npm Trusted Publishers (OIDC) — only the
[`release.yml`](.github/workflows/release.yml) workflow can publish. The flow:

1. PRs land on `main` with their associated changeset.
2. The Changesets action opens a "Version Packages" PR that bumps versions
   and consumes the pending changesets.
3. Merging that PR triggers the release workflow, which runs the full gate
   one more time and publishes to npm.

Maintainers: never publish from a local machine.

## Questions

If anything in this document is unclear, open an issue or start a discussion.
PRs welcome.
