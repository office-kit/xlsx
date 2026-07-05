# Security Policy

## Supported Versions

@office-kit/xlsx is pre-1.0. Security fixes ship on the latest `0.x` minor release.
Older minor versions do not receive backports — please upgrade to the latest
`0.x` to pick up patches. Once `1.0` ships, a longer support window will be
defined here.

| Version  | Supported          |
| -------- | ------------------ |
| 0.6.x    | :white_check_mark: |
| < 0.6    | :x:                |

## Reporting a Vulnerability

Please report vulnerabilities **privately** so users have time to upgrade
before the issue becomes public:

1. **Preferred:** open a [GitHub Private Security Advisory](https://github.com/office-kit/xlsx/security/advisories/new).
   This keeps the report invisible to the public, lets us coordinate the fix
   in a private fork, and produces a CVE on disclosure.
2. **Alternative:** email the maintainer at the address listed in
   `package.json` (`author` field). Use a subject line beginning with
   `[@office-kit/xlsx security]` so the report is triaged ahead of normal mail.

Please include:

- A description of the issue and the impact you believe it has.
- A minimal reproduction (a malformed xlsx, a code snippet, or a CVE-style
  attack scenario). For input-handling bugs, attach the offending file.
- The affected version range (e.g. `>= 0.5.0`, `all versions`).
- Any mitigations you have already verified.

### What to expect

- **Acknowledgement:** within 7 days.
- **Initial assessment:** within 14 days. We'll confirm whether the issue is
  in scope, ask for any missing reproduction details, and propose a
  disclosure timeline.
- **Fix + release:** target 30 days for high-severity issues, 90 days for
  lower-severity ones. Complex bugs that need a redesign may take longer; in
  that case we'll keep you informed.
- **Disclosure:** we publish a GitHub Security Advisory (with CVE when
  warranted) once a patched release is available. Reporters are credited
  unless they request anonymity.

## Scope

In scope:

- Decompression-bomb attacks (zip-bomb / nested archives) against the
  reader. `loadWorkbook` / `loadWorkbookStream` accept untrusted input;
  defaults applied through `decompressionLimits` should refuse to allocate
  unbounded memory on adversarial archives.
- XXE / external-entity attacks against the XML parser. DOCTYPE / ENTITY
  declarations are forbidden in OOXML payloads and must be rejected.
- Prototype-pollution via attacker-controlled keys arriving from XML attribute
  / element names.
- Path traversal during archive extraction (`zip-slip`). OOXML paths flow
  through Part-name resolution, not direct filesystem writes, but report
  anything that lets an attacker write outside the intended scope.
- Memory exhaustion / panics triggered by malformed (but well-formed-looking)
  xlsx files.
- Supply-chain issues in published artefacts (e.g. a tarball that doesn't
  match the source git tag).

Out of scope:

- Issues that require the attacker to control the source code or the host
  process (a malicious dependency added by the consumer, etc.).
- Bugs in our dependencies (`fflate`, `saxes`, `fast-xml-parser`). Please
  report those upstream; we'll backport mitigations once they ship.
- Excel correctness bugs that aren't security-relevant (wrong cell value,
  styling mismatch). Use the public issue tracker for those.
- Denial-of-service via legitimate-but-large inputs that stay within
  `decompressionLimits`. The library is designed to read large workbooks
  efficiently; if you can demonstrate a pathological case (e.g. O(n²)
  behaviour) we'll fix it, but it's a performance issue rather than a
  security issue.

## Hardening recommendations for consumers

If you process xlsx files from untrusted sources:

- Keep `decompressionLimits` at the defaults, or tighten them based on the
  largest file you expect from your users.
- Apply a hard timeout around `loadWorkbook` / `loadWorkbookStream` in
  addition to the size limits.
- Run the library in a process / worker isolated from sensitive state.
- Validate file size at the network / upload layer before invoking the
  library.

The defaults in `DEFAULT_DECOMPRESSION_LIMITS` reject pathological archives
without breaking legitimate xlsx files; review them against your threat
model.
