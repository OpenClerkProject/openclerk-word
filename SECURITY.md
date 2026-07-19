# Security Policy

OpenClerk is a client-side Microsoft Word add-in for legal citation work. It runs entirely inside
Word's WebView, has no backend server of its own, and collects no data (see the
[Privacy Policy](PRIVACY.md) and the README's
[Security & IT review](README.md#security--it-review)). Because it is used in legal work, we treat
citation-verification accuracy and safe hyperlink/HTML insertion as security properties, not just
correctness concerns.

## Reporting a vulnerability

**Please do not report security vulnerabilities through public GitHub issues, discussions, or pull
requests.** Public disclosure before a fix exists puts users at risk.

Instead, report privately through GitHub's private vulnerability reporting:

1. Open the [**Security** tab](https://github.com/OpenClerkProject/openclerk-word/security) of this
   repository.
2. Click **"Report a vulnerability"** to open an advisory visible only to the maintainers.

> **Maintainer note:** if the "Report a vulnerability" button isn't visible, enable it once under
> **Settings → Code security and analysis → Private vulnerability reporting**.

Please include enough detail to reproduce and assess the issue:

- the affected feature, file, and version or commit,
- steps to reproduce, or a proof of concept,
- the impact you believe it has (what data or action it exposes).

OpenClerk is maintained by an individual, in the open, on a **best-effort basis** — there is no paid
support line or guaranteed response time. You can expect an acknowledgment as soon as the maintainer
is able, followed by coordination on a fix and a disclosure timeline.

## Coordinated disclosure

Please give the maintainer a reasonable opportunity to release a fix before disclosing publicly.
Once a fix ships, the advisory can be published and credit given to the reporter (if wanted). There
is no bug-bounty program.

## Supported versions

OpenClerk is distributed as a sideloaded / AppSource Office add-in with **no auto-update mechanism** —
the version running is whatever was last installed. Security fixes are made against the latest
release only; please update to the latest before reporting.

| Version | Supported |
| --- | --- |
| Latest release | ✅ |
| Older releases | ❌ (update to the latest) |

## Scope

**In scope** — the add-in code in this repository (`src/`, `manifest.xml`, and the build/release
tooling under `scripts/`). The security-sensitive areas, specifically:

- **Citation verification / hallucination detection** — a check must never report a fabricated
  citation as "verified" (the project's core trust property).
- **Hyperlink and HTML insertion into the document** — URL-scheme validation and HTML escaping at
  every insertion sink (`src/taskpane/safeInsertion.ts` is the single guarded entry point).
- **Credential handling** for enterprise providers — credentials are held in memory only and sent
  over HTTPS to the provider the user chose (never persisted, never sent elsewhere).

**Out of scope:**

- **Third-party services** OpenClerk can talk to (CourtListener, Westlaw, LexisNexis, Bloomberg
  Law) — report issues in how *they* handle data to those vendors directly.
- **[`openclerk-core`](https://github.com/OpenClerkProject/openclerk-core)** — the shared citation /
  Bluebook logic lives in a separate repository; report issues in that logic there.

## Existing security posture

A point-in-time manual audit is recorded in [SECURITY_AUDIT.md](SECURITY_AUDIT.md), and the README's
[Security & IT review](README.md#security--it-review) section documents the add-in's data flow,
requested permissions, and dependency posture for IT reviewers.
