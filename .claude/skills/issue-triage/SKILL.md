---
name: issue-triage
description: Triage GitHub issues for the @office-kit/xlsx repository — classify as bug / feature request / other, attach the matching label, and drive the right follow-up workflow (reproduce + fix-PR for bugs; policy check + spec + implement-PR for features; individual response for other). Invoke this whenever the user mentions an issue number, asks to "triage", "look at", "handle", "process" an issue, or pastes a GitHub issue URL — even if they don't say the word "triage".
---

# Issue Triage Skill (@office-kit/xlsx)

Triage a GitHub issue end-to-end: read it, classify it, label it, and run
the workflow that matches its kind. @office-kit/xlsx uses three kinds:

1. **Bug report** — something is broken or behaves contrary to docs.
2. **Feature request** — a new capability or API surface.
3. **Other** — questions, doc fixes, build/CI issues, dependency updates,
   meta discussion, etc.

The user invokes this skill by referencing an issue (number or URL) or by
asking to triage. The user's GitHub handle is the **operator** (`op`)
referred to throughout this skill.

## Operating principles

- **Stay on the issue you were given.** Don't sweep adjacent issues unless
  the user asks. One issue per invocation is the default.
- **Communicate before doing irreversible things.** Posting comments,
  applying labels, opening PRs, mentioning users — confirm with the user
  before the first action that's visible on GitHub. After the user
  approves the plan, you can execute the rest of that plan without
  re-confirming each step.
- **Never close the issue yourself.** Triage means routing, not deciding.
  Closing is the maintainer's call.
- **Don't @-mention anyone other than `op` (the issue author).**
  Specifically, don't tag random reviewers, the repo owner, or
  collaborators.
- **Write GitHub-visible text in the language the issue uses.** If the
  issue is in Japanese, your comment is in Japanese. If English, English.

## Step 0 — Locate the issue

Resolve what issue to operate on:

- Number form (`#42`) or URL (`https://github.com/office-kit/xlsx/issues/42`):
  fetch with `gh issue view <num> --repo office-kit/xlsx --json
  number,title,body,labels,author,state,comments`.
- If the user pastes the body inline, work from that and confirm the
  issue number with the user before posting anything.

Read existing labels and comments before doing anything else — somebody
may have already started triaging, and you should not duplicate work or
contradict an existing label without addressing the prior take.

## Step 1 — Classify

Decide one of: `bug`, `enhancement` (feature request), `other`.

| Signal | Likely class |
|---|---|
| "throws", "wrong output", "panics", repro steps + expected vs actual | bug |
| "would be nice if", "support for X", "add an option to …", new API | enhancement |
| "how do I …", "is this supported?" | other (question) |
| typo / wording / link rot in README or docs | other (documentation) |
| renovate / dependabot / lockfile | other (dependencies) |
| CI flake, build error in user environment | usually other (question), unless reproducible against a clean checkout — then bug |

Edge cases:

- A single issue can mix bug + feature. Pick the dominant frame and note
  the other in your triage comment.
- "It's slow" without numbers is a question until the reporter provides a
  benchmark or a fixture; don't escalate to bug yet.
- "It crashes on my file" without the file is a bug-shaped question.
  Apply `bug` only if you can reproduce. Otherwise apply `question` and
  ask `op` for the file.

## Step 2 — Apply the label

Available labels in this repo (see `gh label list --repo
office-kit/xlsx`):

- `bug`, `enhancement`, `documentation`, `question`, `duplicate`,
  `invalid`, `wontfix`, `good first issue`, `help wanted`, `dependencies`

Map class → label:

- bug → `bug`
- enhancement → `enhancement`
- other → the most specific of `documentation` / `question` /
  `dependencies`. Use `question` as the catch-all.

Apply with `gh issue edit <num> --repo office-kit/xlsx --add-label
"<label>"`. Don't remove existing labels unless they conflict; instead,
mention the conflict in your triage comment so the maintainer can adjust.

## Step 3 — Run the workflow for the class

Branch on the class. Each workflow has its own section below.

### 3a. Bug workflow

The goal is: a failing test that pins the bug, then a fix that flips the
test green, then a PR.

1. **Reproduce in a test first.** Add a failing test under `tests/` in the
   phase directory closest to the affected feature (or `tests/phase-0/`
   if you can't tell). Name it after the issue: `issue-<num>.test.ts`.
   Use the smallest fixture that reproduces; prefer in-memory workbooks
   over committing new `.xlsx` binaries.
   - Run with `pnpm vitest run tests/.../issue-<num>.test.ts`.
   - The first run **must fail** for the right reason (the symptom from
     the issue), not a setup error. If it passes immediately, the bug
     either doesn't reproduce or your test is wrong — investigate.

2. **If you reproduced it:** fix the underlying cause, not the symptom.
   Re-run the test to confirm green. Run `pnpm typecheck && pnpm lint &&
   pnpm test` for the affected area before opening the PR. Open a branch
   `fix/issue-<num>-<short-slug>`, commit with a message that links the
   issue (`fix: <summary> (#<num>)`), and open a PR whose body includes:
   - one-line summary of the bug
   - the root cause in 2–3 sentences
   - "Closes #<num>"
   - test command to verify locally

3. **If you couldn't reproduce:** do **not** apply `wontfix` or
   `invalid` — those are maintainer calls. Instead, post a single comment
   on the issue:
   - what you tried (versions, fixture you used, code path you exercised)
   - what specific information you need from `op` to make progress
     (minimal `.xlsx` fixture, exact @office-kit/xlsx version, Node version,
     code snippet, full stack trace)
   - @-mention `op` (the issue author) — and only `op`
   - Then stop. Don't open a speculative PR.

   Example comment skeleton (translate to the issue's language):
   > Thanks for the report! I tried to reproduce against `main` with
   > `<what you tried>` and it succeeded / produced `<actual>`. To move
   > this forward, could you share `<specific thing>`? @<op>

### 3b. Feature request (enhancement) workflow

@office-kit/xlsx's policy: **"There is exactly one way to do one thing."** Adding
a second way to do something — even a more convenient one — increases
cognitive load for every future reader and forces a choice on every
caller. See `references/policy.md` for the full rationale and the
checklist for evaluating a request against this rule.

1. **Read `references/policy.md`** before judging the request. The
   checklist there is the deciding criterion, not your aesthetic
   preference.

2. **Score the request against the policy.** Three outcomes:

   - **Reject (most common):** the request adds a parallel API to
     something @office-kit/xlsx already supports (e.g., "add a `getCellValue`
     helper" when `getCell(...).value` already works). Post a comment
     that:
     - thanks the reporter
     - shows the existing path that already covers their use case (with
       a code snippet pointing at the actual export, e.g.
       `@office-kit/xlsx/worksheet`'s `getCell`)
     - explains *why* @office-kit/xlsx does not add a second path — link to the
       policy section in the README/motivation if applicable
     - leaves the maintainer to close. Do not close yourself.

   - **Accept:** the request enables something that is genuinely
     impossible today, or replaces an existing path that's
     demonstrably worse. In this case, run the **spec → implement → PR**
     flow below.

   - **Unclear / needs more info:** post a question comment to `op` (and
     only `op`) asking for the specific use case that motivated the
     request. Stop until they answer.

3. **Spec → implement → PR (only for accepted requests).**

   a. **Spec the change** as a new comment on the original issue (not a
      separate issue — the original issue is the tracking issue). The
      spec covers:
      - the public API shape (subpath, exported names, types)
      - what existing path it replaces, if any (and a deprecation plan
        if relevant)
      - schema/conformance implications (@office-kit/xlsx validates against
        ECMA-376; new features must round-trip)
      - test plan (unit + at least one round-trip fixture)

   b. **Pause for user (operator) approval.** A spec comment is
      maintainer-visible and commits the project to a direction. Wait
      for the user to say "go" before posting it, and pause again before
      starting the implementation if the spec was non-trivial.

   c. **Implement** on a branch `feat/issue-<num>-<short-slug>`. Tests
      first; production code to make them pass. Run `pnpm typecheck &&
      pnpm lint && pnpm knip && pnpm test && pnpm build && pnpm size`
      before opening the PR — feature work has to clear the bundle-size
      gate, not just tests.

   d. **PR body** includes the spec (or a link to the spec comment),
      "Closes #<num>", and a manual round-trip check that the new path
      survives a load → mutate → save cycle.

### 3c. Other workflow

These are individual judgement calls. The pattern is: respond, don't
fix-by-default.

- **Question (`question`):** answer with a code snippet pointing at the
  existing API. If the answer is in the README or `docs/`, link the
  exact section. Don't open a PR unless the question reveals a doc gap
  the user explicitly asks you to fix.
- **Documentation (`documentation`):** if it's a small fix (typo, broken
  link, wrong code sample) and you've verified the correction, fix it
  and open a PR `docs: <summary> (#<num>)`. If it's a larger doc rewrite,
  treat it as a feature request and run 3b.
- **Dependencies (`dependencies`):** these are usually renovate /
  dependabot. Don't intervene unless the user asks — the bot has its own
  workflow and CI tells the maintainer if the bump is safe.
- **CI / build environment issue:** usually a question. Help the
  reporter narrow the cause, don't change CI yourself unless the user
  asks.
- **Spam / off-topic:** apply `invalid` only on explicit user
  instruction. Otherwise leave it for the maintainer.

## Step 4 — Report back to the user

End with a short status to the user (not a GitHub comment): one or two
sentences saying class, label applied, and what happened or what you're
waiting on. If you opened a PR, include the URL. If you posted a
comment, quote one line of it so the user can verify the tone before
the maintainer sees it.

## Quick command reference

```sh
# Read the issue
gh issue view <num> --repo office-kit/xlsx \
  --json number,title,body,labels,author,state,comments

# Apply a label
gh issue edit <num> --repo office-kit/xlsx --add-label "<label>"

# Comment (use HEREDOC to preserve formatting)
gh issue comment <num> --repo office-kit/xlsx --body "$(cat <<'EOF'
<message>
EOF
)"

# List labels
gh label list --repo office-kit/xlsx

# Run tests for a single file
pnpm vitest run tests/<phase>/issue-<num>.test.ts

# Full pre-PR gate
pnpm typecheck && pnpm lint && pnpm knip && pnpm test && pnpm build && pnpm size
```

## Reference files

- `references/policy.md` — "one way to do one thing" policy with a
  checklist for feature-request evaluation.
