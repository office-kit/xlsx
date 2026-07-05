# Changesets

This directory contains [changesets](https://github.com/changesets/changesets)
that drive version bumps, changelog entries, and npm publishes for `@office-kit/xlsx`.

## Workflow

1. Make a change that should ship in a release.
2. Run `pnpm changeset` and follow the prompts:
   - Select `@office-kit/xlsx` (the `@office-kit/xlsx-site` package is private and ignored).
   - Pick the bump type — `patch`, `minor`, or `major` — following [semver](https://semver.org/).
   - Write a short summary; this becomes the changelog entry.
3. Commit the generated `.md` file in `.changeset/` along with your code change.

## What happens on `main`

The `Release` workflow watches `main`:

- If unreleased changesets exist, it opens (or updates) a **Version Packages**
  PR that bumps `package.json`, regenerates `CHANGELOG.md`, and removes the
  consumed `.changeset/*.md` files.
- Merging that PR triggers a publish to npm with provenance and creates a
  matching GitHub Release.

No manual `npm publish` and no manual tagging — merge the PR and the release
ships.

## Tips

- Multiple changesets in one PR are fine; they are aggregated on release.
- For a release that should not bump the version (docs only, CI tweaks),
  simply do not add a changeset.
- To preview what the next version would be locally, run
  `pnpm changeset status --verbose`.
